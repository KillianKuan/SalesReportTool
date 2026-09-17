"""launcher.py - PyInstaller entry point (Windows + macOS).

Starts Streamlit in a subprocess, shows a system tray / menu bar icon, and
opens the browser when the server is ready.  Single-instance protection
prevents a second copy from starting when one is already running.

All platform divergence lives in this module (``sys.platform`` checks); the
code under ``app/`` is fully shared and unaware of the platform.

| Concern        | Windows (unchanged)     | macOS (.app bundle)                                              |
|----------------|-------------------------|------------------------------------------------------------------|
| Log file       | next to the .exe        | ~/Library/Logs/SalesReportTool/salesreport.log                    |
| Streamlit app  | app/ next to the .exe   | ~/Library/Application Support/SalesReportTool/app (mirrored)      |
| Data folder    | data/ next to the .exe  | ~/Library/Application Support/SalesReportTool/data                |
| overrides.json | app/overrides.json      | ~/Library/Application Support/SalesReportTool/app/overrides.json  |
| settings.json  | app/settings.json       | ~/Library/Application Support/SalesReportTool/app/settings.json   |

A macOS ``.app`` bundle is read-only, so on every launch the bundled ``app/``
folder is staged into a scratch copy, persistent user-state files
(``overrides.json``, ``settings.json``) are restored into it, the result is
validated, and only then atomically swapped in to replace the live copy in
Application Support (see ``mirror_app_dir()``) - a failure partway through
never leaves a half-updated app/ behind.  Streamlit is started from that
copy.  ``utils.py`` then resolves ``data/`` (``APP_DIR.parent / "data"``) and
``overrides.json`` to writable locations without any macOS-specific code in
``app/``.

``aliases.json`` (bundled defaults, git-tracked) and ``settings.json``'s
``customer_aliases`` / ``fcst_customer_aliases`` (user overrides) are two
separate layers merged at read time by ``fcst_loader.py`` - the bundle can
freely update its defaults on every upgrade without ever touching a user's
custom mappings, since they live in a different file.
"""

from __future__ import annotations

import atexit
import json
import os
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path

APP_NAME = "SalesReportTool"
BASE_PORT = 8501
MAX_PORT = 8510
CHILD_MODE_ENV = "SALESREPORT_CHILD"
PORT_ENV = "SALESREPORT_STREAMLIT_PORT"
LOG_MAX_BYTES = 1 * 1024 * 1024   # 1 MB
LOG_KEEP_BYTES = 512 * 1024        # keep last 500 KB when trimming

DATA_SUBDIRS = ("Over the Years", "Current Year", "FCST")
# Files inside app/ that the running app writes to and must survive an upgrade.
# aliases.json is deliberately NOT in this list: it is a bundled default
# (git-tracked, shipped with the app) and is always refreshed from the
# bundle. User overrides on top of it live in settings.json's
# "customer_aliases" / "fcst_customer_aliases" keys instead, so the two
# layers - bundled defaults vs. user overrides - never collide on disk.
USER_STATE_FILES = ("overrides.json", "settings.json")
_USER_STATE_DEFAULTS = {"overrides.json": "[]", "settings.json": "{}"}

# Files that must exist (and be non-empty) in a mirrored app/ copy before it
# is safe to run Streamlit against it.
REQUIRED_RUNTIME_FILES = ("app.py", "utils.py", "charts.py", "fcst_loader.py")


class RuntimeMirrorError(RuntimeError):
    """Raised when the bundled app/ cannot be mirrored into a working, writable copy."""


IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

_LOCK_FILE = Path(tempfile.gettempdir()) / "salesreport.lock"

# Module-level so tray menu callback can read it without a closure
_log_path: Path | None = None
_log_file = None   # open file handle (parent only)


# ---------------------------------------------------------------------------
# Platform-aware path resolution
# ---------------------------------------------------------------------------

def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _source_dir() -> Path:
    """Directory of this file when running from source."""
    return Path(__file__).resolve().parent


def _exe_dir() -> Path:
    """Directory containing the frozen executable."""
    return Path(sys.executable).resolve().parent


def _meipass_dir() -> Path | None:
    meipass = getattr(sys, "_MEIPASS", None)
    return Path(meipass) if meipass else None


def macos_bundle_dir() -> Path | None:
    """Return the enclosing ``*.app`` directory, or None when not bundled."""
    if not (IS_MACOS and is_frozen()):
        return None
    exe = Path(sys.executable).resolve()
    for parent in exe.parents:
        if parent.suffix == ".app":
            return parent
    return None


def is_macos_bundle() -> bool:
    return macos_bundle_dir() is not None


def resource_dirs() -> list[Path]:
    """Read-only locations that may contain the bundled ``app/`` / ``assets/``."""
    dirs: list[Path] = []
    if is_frozen():
        dirs.append(_exe_dir())
        meipass = _meipass_dir()
        if meipass:
            dirs.append(meipass)
        bundle = macos_bundle_dir()
        if bundle:
            dirs.append(bundle / "Contents" / "Resources")
    else:
        dirs.append(_source_dir())
    return dirs


def _find_resource(relative: str) -> Path | None:
    for base in resource_dirs():
        candidate = base / relative
        if candidate.exists():
            return candidate
    return None


def macos_support_dir() -> Path:
    """Writable per-user root used instead of the read-only .app bundle."""
    return Path.home() / "Library" / "Application Support" / APP_NAME


def _get_log_path() -> Path:
    if IS_MACOS and is_frozen():
        log_dir = Path.home() / "Library" / "Logs" / APP_NAME
        log_dir.mkdir(parents=True, exist_ok=True)
        return log_dir / "salesreport.log"
    if is_frozen():
        return _exe_dir() / "salesreport.log"
    return _source_dir() / "salesreport.log"


def data_dir() -> Path:
    """Where the user's Excel/CSV input data lives.

    Windows keeps the historical exe-relative ``data/`` folder.  On macOS the
    .app bundle is read-only, so data lives in Application Support - which is
    exactly ``APP_DIR.parent / "data"`` for the mirrored app folder, so
    ``utils.py`` resolves it with no change.
    """
    if is_macos_bundle():
        return macos_support_dir() / "data"
    if is_frozen():
        return _exe_dir() / "data"
    return _source_dir() / "data"


def _copy_tree_fresh(source: Path, destination: Path) -> None:
    """Replace destination with a full recursive copy of source.

    Removing destination first (rather than ``dirs_exist_ok=True``) is what
    makes files removed from the bundle since the last version disappear
    from the mirrored copy instead of lingering forever.
    """
    if destination.exists():
        shutil.rmtree(destination)
    shutil.copytree(source, destination)


def _validate_runtime_dir(app_dir: Path) -> list[str]:
    """Return the required runtime files that are missing or empty under app_dir."""
    problems = []
    for name in REQUIRED_RUNTIME_FILES:
        path = app_dir / name
        if not path.is_file() or path.stat().st_size == 0:
            problems.append(name)
    return problems


def _restore_user_state(active_dir: Path, staging_dir: Path) -> None:
    """Carry persistent user-state files from the live app/ into the staged copy.

    ``staging_dir`` was just built fresh from the bundle, so it currently
    holds the bundled placeholder versions of ``overrides.json`` /
    ``settings.json``. Overwrite those with the user's real files whenever
    the app has run before; otherwise leave the staged default in place (or
    create one) so the file always exists.
    """
    for name in USER_STATE_FILES:
        source = active_dir / name
        destination = staging_dir / name
        if source.is_file():
            try:
                shutil.copy2(source, destination)
                continue
            except Exception as exc:
                print(f"WARNING: could not preserve {name} from active app dir: {exc}")
        if not destination.exists():
            try:
                destination.write_text(_USER_STATE_DEFAULTS.get(name, "{}"), encoding="utf-8")
            except Exception as exc:
                print(f"WARNING: could not create default {destination}: {exc}")


def _backup_dir_for(active_dir: Path) -> Path:
    return active_dir.parent / (active_dir.name + ".previous")


def _recover_interrupted_swap(active_dir: Path) -> None:
    """Recover from a process crash between the two renames of a previous
    atomic swap: active_dir would be missing while the backup still holds
    the last known-good copy. Restore it before attempting a new mirror."""
    backup_dir = _backup_dir_for(active_dir)
    if not active_dir.exists() and backup_dir.exists():
        os.rename(backup_dir, active_dir)


def _activate_staged_app(staging_dir: Path, active_dir: Path) -> None:
    """Atomically swap staging_dir into active_dir's place.

    The current active_dir (if any) is moved aside to a backup path first,
    then staging_dir is renamed into active_dir's place. Both are simple
    directory renames on the same filesystem, so each step either fully
    succeeds or fully fails - there is no window where active_dir is a
    stale/new file mix. If the second rename fails, the backup is moved
    back so active_dir is never left missing.
    """
    backup_dir = _backup_dir_for(active_dir)
    if backup_dir.exists():
        shutil.rmtree(backup_dir, ignore_errors=True)

    moved_old = False
    if active_dir.exists():
        os.rename(active_dir, backup_dir)
        moved_old = True

    try:
        os.rename(staging_dir, active_dir)
    except Exception:
        if moved_old:
            os.rename(backup_dir, active_dir)   # roll back so active_dir still works
        raise

    if moved_old:
        shutil.rmtree(backup_dir, ignore_errors=True)


def mirror_app_dir(source_dir: Path | None, active_dir: Path) -> Path:
    """Stage a fresh copy of the bundled app/, merge in user state, validate,
    then atomically activate it as active_dir. Runs on every launch so
    app-code updates ship with the .app.

    On failure, the previously-active copy is left untouched and used
    instead (with a warning logged) whenever it is itself still valid;
    RuntimeMirrorError is raised only when there is no usable copy at all,
    so the caller can show an actionable startup error.
    """
    active_dir.parent.mkdir(parents=True, exist_ok=True)
    _recover_interrupted_swap(active_dir)

    if source_dir is None or source_dir.resolve() == active_dir.resolve():
        # Nothing to mirror from (dev/source run, or already the live copy).
        active_dir.mkdir(parents=True, exist_ok=True)
        for name in USER_STATE_FILES:
            path = active_dir / name
            if not path.exists():
                try:
                    path.write_text(_USER_STATE_DEFAULTS.get(name, "{}"), encoding="utf-8")
                except Exception as exc:
                    print(f"WARNING: could not create {path}: {exc}")
        return active_dir

    staging_dir = active_dir.parent / (active_dir.name + ".staging")

    try:
        _copy_tree_fresh(source_dir, staging_dir)
        _restore_user_state(active_dir, staging_dir)
        problems = _validate_runtime_dir(staging_dir)
        if problems:
            raise RuntimeMirrorError(
                "Bundled app copy is missing required file(s): " + ", ".join(problems)
            )
        _activate_staged_app(staging_dir, active_dir)
    except Exception as exc:
        shutil.rmtree(staging_dir, ignore_errors=True)
        if not _validate_runtime_dir(active_dir):
            print(f"WARNING: could not update app copy, keeping previous version: {exc}")
            return active_dir
        raise RuntimeMirrorError(
            f"Could not prepare a working application copy at {active_dir}: {exc}"
        ) from exc

    return active_dir


def _mirror_app_dir() -> Path:
    """Resolve the bundled source and live target, then mirror (macOS only)."""
    target = macos_support_dir() / "app"
    source = _find_resource("app/app.py")
    source_dir = source.parent if source is not None else None
    return mirror_app_dir(source_dir, target)


def _seed_data_dir(target: Path) -> None:
    """Create the data/ layout and seed it from bundled samples when empty."""
    for sub in DATA_SUBDIRS:
        (target / sub).mkdir(parents=True, exist_ok=True)

    sample = _find_resource("data")
    if sample is None or sample.resolve() == target.resolve():
        return

    for sub in DATA_SUBDIRS:
        src, dst = sample / sub, target / sub
        if not src.is_dir() or any(dst.iterdir()):
            continue
        for item in src.iterdir():
            try:
                if item.is_dir():
                    shutil.copytree(item, dst / item.name, dirs_exist_ok=True)
                else:
                    shutil.copy2(item, dst / item.name)
            except Exception as exc:
                print(f"WARNING: could not seed {item}: {exc}")


def prepare_runtime_dirs() -> None:
    """First-launch setup for the read-only macOS bundle.  No-op elsewhere.

    Raises RuntimeMirrorError if no working app/ copy could be prepared at
    all; callers should show an actionable error rather than starting
    Streamlit against a broken or missing directory.
    """
    if not is_macos_bundle():
        return
    app_dir = _mirror_app_dir()
    _seed_data_dir(data_dir())
    print(f"App dir:  {app_dir}", flush=True)
    print(f"Data dir: {data_dir()}", flush=True)


# ---------------------------------------------------------------------------
# Log file setup  (parent process only)
# ---------------------------------------------------------------------------

def _trim_log(path: Path) -> None:
    """If the log exceeds LOG_MAX_BYTES, keep only the last LOG_KEEP_BYTES."""
    try:
        size = path.stat().st_size
    except FileNotFoundError:
        return
    if size <= LOG_MAX_BYTES:
        return
    with open(path, "rb") as f:
        f.seek(-LOG_KEEP_BYTES, 2)
        tail = f.read()
    with open(path, "wb") as f:
        f.write(b"[... log trimmed ...]\n")
        f.write(tail)


def setup_logging() -> None:
    """Redirect parent stdout/stderr to the log file and write a launch header."""
    global _log_path, _log_file

    _log_path = _get_log_path()
    _trim_log(_log_path)

    _log_file = open(_log_path, "a", encoding="utf-8", errors="replace", buffering=1)

    header = f"\n--- Launch: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n"
    _log_file.write(header)
    _log_file.flush()

    sys.stdout = _log_file
    sys.stderr = _log_file

    atexit.register(_close_log)


def _close_log() -> None:
    global _log_file
    if _log_file is not None:
        try:
            _log_file.flush()
            _log_file.close()
        except Exception:
            pass
        _log_file = None


def _open_log() -> None:
    """Open the log file in the OS default text viewer."""
    if _log_path is None or not _log_path.exists():
        return
    if IS_WINDOWS:
        os.startfile(str(_log_path))
    elif IS_MACOS:
        subprocess.Popen(["open", str(_log_path)])
    else:
        subprocess.Popen(["xdg-open", str(_log_path)])


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def is_port_in_use(port: int) -> bool:
    try:
        with socket.create_connection(("localhost", port), timeout=0.3):
            return True
    except OSError:
        return False


def find_free_port() -> int:
    for port in range(BASE_PORT, MAX_PORT + 1):
        if not is_port_in_use(port):
            return port
    return BASE_PORT


def wait_for_server(url: str, max_wait: int = 60) -> bool:
    for _ in range(max_wait * 2):
        try:
            with urllib.request.urlopen(url, timeout=1):
                return True
        except Exception:
            time.sleep(0.5)
    return False


def get_app_path() -> Path:
    """Path to app/app.py that Streamlit should run."""
    candidates: list[Path] = []
    if is_macos_bundle():
        # Run from the writable mirror so Streamlit can write overrides.json
        candidates.append(macos_support_dir() / "app" / "app.py")
    candidates += [base / "app" / "app.py" for base in resource_dirs()]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    searched = "\n".join(str(p) for p in candidates)
    raise FileNotFoundError(f"app.py not found. Searched:\n{searched}")


def is_child_mode() -> bool:
    return os.environ.get(CHILD_MODE_ENV) == "1"


# ---------------------------------------------------------------------------
# Single-instance lock file helpers
# ---------------------------------------------------------------------------

def read_lock() -> dict | None:
    try:
        return json.loads(_LOCK_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def write_lock(pid: int, port: int) -> None:
    _LOCK_FILE.write_text(json.dumps({"pid": pid, "port": port}), encoding="utf-8")


def remove_lock() -> None:
    try:
        _LOCK_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def check_single_instance() -> None:
    """If another instance is already serving, focus its browser tab and exit.

    Uses port liveness rather than PID signals - os.kill(pid, 0) is unreliable
    on Windows.  The same check works unchanged on macOS; the lock file lives
    in TMPDIR on both platforms.
    """
    try:
        lock = read_lock()
        if lock is None:
            return

        port = int(lock["port"])
    except Exception:
        # Malformed lock - remove and continue
        remove_lock()
        return

    if is_port_in_use(port):
        # Streamlit is still serving on that port: another instance is live.
        print(f"Instance already running on port {port}, opening browser and exiting")
        webbrowser.open(f"http://localhost:{port}")
        sys.exit(0)
    else:
        # Port is free: the previous instance is gone, lock is stale.
        print(f"Stale lock found (port {port} not in use), removing lock")
        remove_lock()


# ---------------------------------------------------------------------------
# System tray / menu bar icon
# ---------------------------------------------------------------------------

def _icon_candidates() -> list[Path]:
    """Icon files to try, in order.  macOS prefers .icns, Windows .ico."""
    names = ("app.icns", "app.png", "app.ico") if IS_MACOS else ("app.ico", "app.png")
    return [base / "assets" / name for base in resource_dirs() for name in names]


def _make_icon_image():
    """Return a PIL Image: bundled icon if present, else a bar-chart fallback."""
    from PIL import Image, ImageDraw

    for path in _icon_candidates():
        if path.exists():
            try:
                return Image.open(path).convert("RGBA").resize((64, 64))
            except Exception:
                pass

    # Fallback: draw a 3-bar ascending chart on a transparent background
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    bar_w = 14
    gap = 4
    x0 = 10
    bottom = 58   # baseline y

    bars = [
        (x0,                     int(64 * 0.70), (74,  144, 217, 255)),   # 30 % height
        (x0 + bar_w + gap,       int(64 * 0.40), (91,  160, 233, 255)),   # 60 % height
        (x0 + 2 * (bar_w + gap), int(64 * 0.10), (122, 184, 245, 255)),   # 90 % height
    ]

    for bx, top_y, color in bars:
        # Fill
        draw.rectangle([bx, top_y, bx + bar_w - 1, bottom], fill=color)
        # 1-px white outline
        draw.rectangle([bx, top_y, bx + bar_w - 1, bottom], outline=(255, 255, 255, 180))

    return img


def build_tray_icon(port: int, proc: "subprocess.Popen[bytes]"):
    """Build and return a pystray.Icon (not yet running)."""
    try:
        import pystray
        from pystray import MenuItem as Item
    except ImportError:
        _show_fatal("pystray is not installed.\nRebuild with: pip install pystray Pillow")
        sys.exit(1)

    url = f"http://localhost:{port}"

    def on_open(icon, item):
        webbrowser.open(url)

    def on_open_log(icon, item):
        _open_log()

    def on_quit(icon, item):
        icon.stop()
        _terminate_child(proc)
        remove_lock()

    menu = pystray.Menu(
        Item("Open Browser", on_open, default=True),
        Item("Open Log", on_open_log),
        Item("Quit", on_quit),
    )

    image = _make_icon_image()
    icon = pystray.Icon(
        "SalesReportTool",
        image,
        f"Sales Report Tool \u2014 port {port}",
        menu,
    )
    return icon


def _terminate_child(proc: "subprocess.Popen[bytes]") -> None:
    if proc.poll() is not None:
        return
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()


def _show_fatal(msg: str) -> None:
    """Show an error dialog even without a console."""
    if IS_WINDOWS:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "Sales Report Tool - Error", 0x10)
            return
        except Exception:
            pass
    elif IS_MACOS:
        try:
            script = (
                f"display dialog {json.dumps(msg)} "
                'with title "Sales Report Tool" buttons {"OK"} with icon stop'
            )
            subprocess.Popen(["osascript", "-e", script])
            return
        except Exception:
            pass
    print(msg, file=sys.stderr)


# ---------------------------------------------------------------------------
# Child-mode entry point (runs Streamlit in-process)
# ---------------------------------------------------------------------------

def run_streamlit_child() -> None:
    app_path = get_app_path()

    port = os.environ.get(PORT_ENV, str(BASE_PORT))
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = port
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    sys.argv = ["streamlit", "run", str(app_path)]
    from streamlit.web import cli as stcli
    stcli.main()


def build_child_command() -> list[str]:
    if is_frozen():
        return [sys.executable]
    return [sys.executable, str(_source_dir() / "launcher.py")]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    # Child mode must be detected before log redirect (Streamlit owns its I/O)
    if is_child_mode():
        run_streamlit_child()
        return

    # Single-instance check BEFORE log redirect: the second instance must exit
    # quickly without competing for the log file.
    check_single_instance()

    # Redirect parent stdout/stderr to log file
    setup_logging()

    try:
        _main_parent()
    except RuntimeMirrorError as exc:
        traceback.print_exc()
        _show_fatal(
            "Sales Report Tool could not prepare its application files.\n\n"
            f"{exc}\n\n"
            "Try reinstalling the app. If the problem persists, check the log:\n"
            f"{_log_path}"
        )
        raise
    except Exception:
        traceback.print_exc()
        _show_fatal(
            "Sales Report Tool crashed at startup.\n"
            f"See log for details:\n{_log_path}"
        )
        raise


def _main_parent() -> None:
    # ---- Single-instance check ----
    check_single_instance()

    # ---- macOS: mirror app/ + create data folders outside the .app bundle ----
    prepare_runtime_dirs()

    # ---- Find free port and write lock ----
    port = find_free_port()
    url = f"http://localhost:{port}"
    print(f"Starting on port {port}", flush=True)

    write_lock(os.getpid(), port)
    atexit.register(remove_lock)

    def _sig_handler(signum, frame):
        remove_lock()
        sys.exit(0)

    try:
        signal.signal(signal.SIGTERM, _sig_handler)
        signal.signal(signal.SIGINT, _sig_handler)
    except (OSError, ValueError):
        pass

    # ---- Start Streamlit child, piping its output into the log file ----
    cmd = build_child_command()
    env = dict(os.environ, **{CHILD_MODE_ENV: "1", PORT_ENV: str(port)})

    proc = subprocess.Popen(
        cmd,
        env=env,
        stdout=_log_file,
        stderr=_log_file,
    )
    print(f"Child PID: {proc.pid}", flush=True)

    # ---- Background thread: wait for server -> open browser -> watch child ----
    def _startup_thread(icon_ref: list):
        ready = wait_for_server(url)
        if ready:
            print("Server ready, opening browser", flush=True)
            webbrowser.open(url)
            for _ in range(20):
                if icon_ref[0] is not None:
                    try:
                        icon_ref[0].notify("Sales Report Tool is running")
                    except Exception:
                        pass
                    break
                time.sleep(0.25)
        else:
            print("WARNING: server did not respond within timeout", flush=True)

        rc = proc.wait()
        print(f"Child exited with code {rc}", flush=True)
        if rc not in (0, None, -15):   # -15 = SIGTERM (normal quit)
            print(f"WARNING: non-zero exit code {rc} - check log for errors", flush=True)
        remove_lock()
        if icon_ref[0] is not None:
            try:
                icon_ref[0].stop()
            except Exception:
                pass

    icon_ref: list = [None]
    threading.Thread(target=_startup_thread, args=(icon_ref,), daemon=True).start()

    # ---- Build and run tray icon on main thread ----
    try:
        icon = build_tray_icon(port, proc)
    except Exception:
        traceback.print_exc()
        _show_fatal(
            "Failed to create system tray icon.\n"
            f"See log:\n{_log_path}"
        )
        proc.wait()
        return

    icon_ref[0] = icon
    print("Tray icon starting", flush=True)
    icon.run()   # blocks until icon.stop()

    # ---- Cleanup after tray exits ----
    _terminate_child(proc)
    remove_lock()
    print("Shutdown complete", flush=True)


if __name__ == "__main__":
    main()

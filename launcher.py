"""launcher.py - PyInstaller entry point.

Starts Streamlit in a hidden subprocess, opens the browser when ready,
and shuts the app down after the browser tab stops sending heartbeats.
"""

import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

BASE_PORT = 8501
MAX_PORT = 8510  # try up to 10 ports
HEARTBEAT_PORT = 8502
HEARTBEAT_TIMEOUT = 10
CREATE_NO_WINDOW = 0x08000000

_last_heartbeat = time.time()


def is_port_in_use(port: int) -> bool:
    """Check if a port is already in use."""
    try:
        with socket.create_connection(("localhost", port), timeout=0.3):
            return True
    except OSError:
        return False


def find_free_port() -> int:
    """Find the first available port."""
    for port in range(BASE_PORT, MAX_PORT + 1):
        if not is_port_in_use(port):
            return port
    return BASE_PORT  # fallback


def wait_for_server(url: str, max_wait: int = 30) -> bool:
    """Wait until the Streamlit HTTP endpoint responds."""
    for _ in range(max_wait * 2):
        try:
            with urllib.request.urlopen(url, timeout=1):
                return True
        except Exception:
            time.sleep(0.5)
    return False


def get_app_path() -> Path:
    """Get the absolute path to app.py in both frozen and source mode."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    return base / "app" / "app.py"


class _HeartbeatHandler(BaseHTTPRequestHandler):
    """Tiny local HTTP handler that records browser heartbeats."""

    def do_GET(self) -> None:
        global _last_heartbeat
        _last_heartbeat = time.time()
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def do_POST(self) -> None:
        self.do_GET()

    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

    def log_message(self, format: str, *args) -> None:
        del format, args


def run_heartbeat_server() -> None:
    """Start the local heartbeat listener."""
    server = HTTPServer(("127.0.0.1", HEARTBEAT_PORT), _HeartbeatHandler)
    server.serve_forever()


def monitor_heartbeat(proc: subprocess.Popen[bytes]) -> None:
    """Stop Streamlit after the browser stops pinging for too long."""
    while proc.poll() is None:
        time.sleep(2)
        if time.time() - _last_heartbeat <= HEARTBEAT_TIMEOUT:
            continue
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
        break


def main() -> None:
    app_path = get_app_path()
    if not app_path.exists():
        raise FileNotFoundError(f"app.py not found: {app_path}")

    port = find_free_port()
    url = f"http://localhost:{port}"

    threading.Thread(target=run_heartbeat_server, daemon=True).start()

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app_path),
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
    ]
    env = dict(
        os.environ,
        STREAMLIT_GLOBAL_DEVELOPMENT_MODE="false",
        APP_HEARTBEAT_PORT=str(HEARTBEAT_PORT),
    )
    proc = subprocess.Popen(
        cmd,
        creationflags=CREATE_NO_WINDOW,
        env=env,
    )

    wait_for_server(url)
    webbrowser.open(url)
    threading.Thread(target=monitor_heartbeat, args=(proc,), daemon=True).start()
    proc.wait()


if __name__ == "__main__":
    main()



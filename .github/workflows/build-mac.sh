#!/usr/bin/env bash
# macOS (Apple Silicon) release build.
#
# Produces a double-clickable dist/SalesReportTool.app for arm64. PyInstaller
# cannot cross-compile, so the Windows .exe is still built by build.bat or
# .github/workflows/build-windows.yml on a Windows runner.
#
# This script is the single source of truth for the macOS PyInstaller flags:
# .github/workflows/build-macos.yml calls it with --skip-deps.
#
# Usage:
#   ./build-mac.sh              # install deps, then build
#   ./build-mac.sh --skip-deps  # CI: dependencies already installed
set -euo pipefail
cd "$(dirname "$0")"

SKIP_DEPS=0
for arg in "$@"; do
    case "$arg" in
        --skip-deps) SKIP_DEPS=1 ;;
        *) echo "Unknown option: $arg" >&2; exit 2 ;;
    esac
done

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "ERROR: this script builds a macOS .app and must run on macOS." >&2
    exit 1
fi
if [[ "$(uname -m)" != "arm64" ]]; then
    echo "WARNING: not running on Apple Silicon (uname -m = $(uname -m));" >&2
    echo "         the produced .app will not be arm64." >&2
fi

if [[ "$SKIP_DEPS" -eq 0 ]]; then
    echo "[1/5] Installing dependencies..."
    pip install -r requirements.txt
    pip install pyinstaller
else
    echo "[1/5] Skipping dependency install (--skip-deps)"
fi

echo "[2/5] Preparing assets/app.icns..."
if [[ ! -f assets/app.icns ]]; then
    WORKDIR="$(mktemp -d)"
    ICONSET="$WORKDIR/app.iconset"
    BASE_PNG="$WORKDIR/app.png"
    mkdir -p "$ICONSET"
    sips -s format png assets/app.ico --out "$BASE_PNG" > /dev/null
    for size in 16 32 64 128 256 512; do
        sips -z "$size" "$size" "$BASE_PNG" \
            --out "$ICONSET/icon_${size}x${size}.png" > /dev/null
        double=$((size * 2))
        sips -z "$double" "$double" "$BASE_PNG" \
            --out "$ICONSET/icon_${size}x${size}@2x.png" > /dev/null
    done
    iconutil -c icns "$ICONSET" -o assets/app.icns
    rm -rf "$WORKDIR"
    echo "      generated assets/app.icns from assets/app.ico"
else
    echo "      assets/app.icns already present"
fi

echo "[3/5] Cleaning old build artifacts..."
rm -rf dist build launcher.spec

# PyInstaller flags mirror build.bat / build-windows.yml. Keep them in sync.
# macOS-only differences: --windowed (.app bundle), .icns icon, bundle id,
# arm64 target, and --add-data so app/ + assets/ ship inside the bundle.
echo "[4/5] Building SalesReportTool.app (arm64)..."
python -m PyInstaller \
    --name SalesReportTool \
    --onedir \
    --windowed \
    --icon assets/app.icns \
    --osx-bundle-identifier com.killiankuan.salesreporttool \
    --target-architecture arm64 \
    --add-data "app:app" \
    --add-data "assets:assets" \
    --collect-all streamlit \
    --copy-metadata streamlit \
    --hidden-import streamlit.web.cli \
    --hidden-import streamlit.web.bootstrap \
    --hidden-import streamlit.runtime.scriptrunner \
    --hidden-import streamlit.runtime.caching \
    --hidden-import streamlit.runtime.secrets \
    --hidden-import pkg_resources \
    --collect-data altair \
    --collect-data pydeck \
    --collect-data packaging \
    --hidden-import python_calamine \
    --hidden-import pystray \
    --hidden-import PIL \
    --collect-all pystray \
    launcher.py

echo "[5/5] Verifying..."
if [[ ! -d dist/SalesReportTool.app ]]; then
    echo "ERROR: build failed, dist/SalesReportTool.app not found" >&2
    exit 1
fi

echo
echo "============================================"
echo " Build complete: dist/SalesReportTool.app"
echo
echo " The .app is unsigned: on first launch use"
echo " right-click -> Open to bypass Gatekeeper."
echo
echo " Data folder : ~/Library/Application Support/SalesReportTool/data"
echo " Log file    : ~/Library/Logs/SalesReportTool/salesreport.log"
echo "============================================"

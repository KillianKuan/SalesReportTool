#!/usr/bin/env bash
# Local smoke-test build for macOS / Linux ONLY.
#
# PyInstaller cannot cross-compile: running this produces a binary for the HOST
# OS (macOS/Linux), NOT a distributable Windows .exe. Use it only to verify that
# the PyInstaller build succeeds locally.
#
# To produce the Windows .exe for end users, push a vX.Y.Z tag and let the
# GitHub Actions workflow (.github/workflows/build-windows.yml) build it, or run
# build.bat on Windows.
#
# Flags below mirror build.bat so local and CI builds stay in sync.
set -euo pipefail
cd "$(dirname "$0")"

echo "[1/4] Installing dependencies..."
pip install -r requirements.txt
pip install pyinstaller

echo "[2/4] Cleaning old build artifacts..."
rm -rf dist build launcher.spec

echo "[3/4] Building (host-OS binary, smoke-test only)..."
python -m PyInstaller \
    --name SalesReportTool \
    --onedir \
    --noconsole \
    --icon assets/app.ico \
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

echo "[4/4] Copying app/ and data/ into dist/SalesReportTool/..."
cp -R app dist/SalesReportTool/app
if [ -d data ]; then
    cp -R data dist/SalesReportTool/data
else
    mkdir -p dist/SalesReportTool/data
fi

echo
echo "Done. Host-OS binary at dist/SalesReportTool/ (NOT a Windows .exe; smoke-test only)."

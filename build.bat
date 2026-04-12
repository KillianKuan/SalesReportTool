@echo off
chcp 65001 > nul
echo ============================================
echo  SalesReportTool Build Script
echo ============================================

echo [1/5] Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller
if errorlevel 1 (
    echo X Package installation failed
    pause
    exit /b 1
)

echo [2/5] Cleaning old files...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist launcher.spec del launcher.spec

echo [3/5] Building (this may take a few minutes)...
python -m PyInstaller ^
    --name SalesReportTool ^
    --onedir ^
    --noconsole ^
    --icon NONE ^
    --collect-all streamlit ^
    --copy-metadata streamlit ^
    --hidden-import streamlit.web.cli ^
    --hidden-import streamlit.web.bootstrap ^
    --hidden-import streamlit.runtime.scriptrunner ^
    --hidden-import streamlit.runtime.caching ^
    --hidden-import streamlit.runtime.secrets ^
    --hidden-import pkg_resources ^
    --collect-data altair ^
    --collect-data pydeck ^
    --collect-data packaging ^
    --hidden-import python_calamine ^
    launcher.py
if errorlevel 1 (
    echo.
    echo X PyInstaller build failed! Check the error messages above.
    pause
    exit /b 1
)

echo.
echo Checking build output...
if not exist "dist\SalesReportTool\SalesReportTool.exe" (
    echo.
    echo X Build failed! .exe not found
    pause
    exit /b 1
)
echo OK .exe generated

echo [4/5] Preparing output folder...
xcopy /e /i /y app dist\SalesReportTool\app
if exist data (
    xcopy /e /i /y data dist\SalesReportTool\data
    echo OK data\ folder copied to output
) else (
    mkdir "dist\SalesReportTool\data"
    echo WARN data\ folder not found, created empty folder
)

echo [5/5] Verifying...
dir "dist\SalesReportTool\SalesReportTool.exe"

echo.
echo ============================================
echo  Build complete! Output: dist\SalesReportTool\
echo  Make sure data\ contains year subfolders (e.g. 2024\, 2025\)
echo  Distribute the entire folder to end users.
echo ============================================
pause

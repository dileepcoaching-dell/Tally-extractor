@echo off
title Tally Data Extractor — Launcher
color 0A

echo.
echo  ============================================
echo    Tally Data Extractor  ^|  Starting...
echo  ============================================
echo.

:: ── Check Python is available ──────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo  [ERROR] Python is not installed or not in PATH.
    echo  Please install Python 3.9+ from https://python.org
    echo.
    pause
    exit /b 1
)

:: ── Install / upgrade dependencies silently ────────────────────
echo  [1/2] Checking dependencies...
pip install -r "%~dp0requirements.txt" --quiet --disable-pip-version-check
if errorlevel 1 (
    color 0C
    echo  [ERROR] Failed to install dependencies.
    echo  Try running:  pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)
echo        Dependencies OK.
echo.

:: ── Launch Streamlit app ───────────────────────────────────────
echo  [2/2] Launching Tally Data Extractor...
echo        The app will open in your browser at http://localhost:8501
echo.
echo  ============================================
echo   Press Ctrl+C in this window to stop the app
echo  ============================================
echo.

cd /d "%~dp0"
streamlit run tally_extractor.py --server.headless false --browser.gatherUsageStats false

:: ── If streamlit exits (e.g. user closed it) ──────────────────
echo.
echo  App has stopped. Press any key to close this window.
pause >nul

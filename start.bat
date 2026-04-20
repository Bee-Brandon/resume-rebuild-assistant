@echo off
title Resume Rebuild Assistant
echo.
echo  ========================================
echo   Resume Rebuild Assistant - Starting...
echo  ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python is not installed or not on PATH.
    echo  Download from: https://www.python.org/downloads/
    echo  Make sure to check "Add to PATH" during install.
    pause
    exit /b 1
)

:: Check if venv exists, create if not
if not exist ".venv" (
    echo  First run - setting up environment...
    echo  This may take a minute.
    echo.
    pip install uv >nul 2>&1
    uv venv
    uv pip install -r requirements.txt
    echo.
    echo  Generating default template...
    uv run python create_template.py
    echo.
)

:: Check for Poppler (optional - only needed for scanned PDFs)
where pdftoppm >nul 2>&1
if errorlevel 1 (
    echo  NOTE: Poppler not found. Scanned PDFs won't work.
    echo  Text-based PDFs, Word docs, and images will work fine.
    echo.
)

:: Launch
echo  Starting app... (opening browser)
echo  Press Ctrl+C in this window to stop.
echo.
uv run streamlit run app.py

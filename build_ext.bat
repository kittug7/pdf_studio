@echo off
setlocal
cd /d "%~dp0"
title PDF Studio - Build EXE

echo ============================================
echo   PDF Studio - building PDFStudio.exe
echo ============================================
echo.

rem --- find a REAL Python (the Microsoft Store shortcut does not count) ---
set PY=
py -3 --version >nul 2>nul && set PY=py -3
if not defined PY (
    python --version >nul 2>nul && set PY=python
)
if not defined PY (
    echo [ERROR] Python is not installed ^(or only the Microsoft Store shortcut exists^).
    echo.
    echo   1. Download Python 3.10+ from https://www.python.org/downloads/windows/
    echo   2. In the installer, TICK "Add python.exe to PATH", then click Install.
    echo   3. Close this window, then double-click build_exe.bat again.
    echo.
    echo   Optional: Settings ^> Apps ^> Advanced app settings ^> App execution aliases
    echo   and turn OFF the python.exe / python3.exe entries.
    pause
    exit /b 1
)
echo Using Python: %PY%
%PY% --version

if not exist "main.py" (
    echo [ERROR] main.py not found. Run this file from the PDFStudio folder.
    pause
    exit /b 1
)

if not exist "venv\Scripts\activate.bat" (
    echo [1/4] Creating virtual environment...
    %PY% -m venv venv
    if errorlevel 1 goto :fail
) else (
    echo [1/4] Using existing virtual environment.
)
call "venv\Scripts\activate.bat"

echo [2/4] Installing dependencies...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 goto :fail

echo [3/4] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist PDFStudio.spec del /q PDFStudio.spec

echo [4/4] Building the executable (this can take a few minutes)...
set ICON=
if exist app.ico set ICON=--icon app.ico
pyinstaller --noconfirm --windowed --name PDFStudio %ICON% ^
    --collect-all pdf2docx ^
    --collect-all pymupdf ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    main.py
if errorlevel 1 goto :fail

echo.
echo ============================================
echo   Build complete!
echo   Your app: dist\PDFStudio\PDFStudio.exe
echo   Share the whole dist\PDFStudio folder.
echo ============================================
start "" "dist\PDFStudio"
pause
exit /b 0

:fail
echo.
echo [ERROR] Build failed. Scroll up to see the error message.
pause
exit /b 1
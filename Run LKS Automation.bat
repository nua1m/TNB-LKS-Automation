@echo off
setlocal
cd /d "%~dp0"

if exist "updater.exe" (
    "updater.exe" --launch launcher.exe
) else if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
    "%PYTHON_EXE%" updater.py --launch launcher.py
) else (
    set "PYTHON_EXE=python"
    "%PYTHON_EXE%" updater.py --launch launcher.py
)
pause

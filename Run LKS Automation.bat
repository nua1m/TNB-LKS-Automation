@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "REQ_MARKER=%VENV_DIR%\requirements.sha256"
set "BOOTSTRAP_PY="

if not exist "%VENV_PY%" (
    call :find_python
    if not defined BOOTSTRAP_PY (
        echo Python 3.11 or 3.10 was not found.
        echo Install Python, then run this launcher again.
        goto :end
    )

    echo Creating virtual environment...
    !BOOTSTRAP_PY! -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Failed to create the virtual environment.
        goto :end
    )
)

call :sync_requirements
if errorlevel 1 goto :end

"%VENV_PY%" updater.py --launch launcher.py

:end
pause
exit /b 0

:find_python
where py >nul 2>nul
if not errorlevel 1 (
    py -3.11 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_PY=py -3.11"
        exit /b 0
    )

    py -3.10 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_PY=py -3.10"
        exit /b 0
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    set "BOOTSTRAP_PY=python"
)
exit /b 0

:sync_requirements
for /f "tokens=* delims=" %%H in ('certutil -hashfile "requirements.txt" SHA256 ^| findstr /R "^[0-9A-F][0-9A-F]"') do (
    set "CURRENT_REQ_HASH=%%H"
    goto :have_hash
)

echo Failed to calculate the requirements hash.
exit /b 1

:have_hash

set "INSTALLED_REQ_HASH="
if exist "%REQ_MARKER%" (
    set /p INSTALLED_REQ_HASH=<"%REQ_MARKER%"
)

if /I "%CURRENT_REQ_HASH%"=="%INSTALLED_REQ_HASH%" (
    exit /b 0
)

echo Installing Python dependencies...
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install requirements.
    exit /b 1
)

>"%REQ_MARKER%" echo %CURRENT_REQ_HASH%
exit /b 0

@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "REQ_MARKER=%VENV_DIR%\requirements.sha256"
set "BOOTSTRAP_MODE="

if not exist "%VENV_PY%" (
    call :find_python
    if not defined BOOTSTRAP_MODE (
        echo Python 3.11 or 3.10 was not found.
        echo Install Python, then run this launcher again.
        goto :end
    )

    echo Creating virtual environment...
    call :create_venv
    if errorlevel 1 (
        echo Failed to create the virtual environment.
        goto :end
    )

    if not exist "%VENV_PY%" (
        echo Virtual environment was not created successfully.
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
        set "BOOTSTRAP_MODE=PY311"
        exit /b 0
    )

    py -3.10 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_MODE=PY310"
        exit /b 0
    )

    py -3.12 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_MODE=PY312"
        exit /b 0
    )

    py -3.13 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_MODE=PY313"
        exit /b 0
    )

    py -3.14 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_MODE=PY314"
        exit /b 0
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    set "BOOTSTRAP_MODE=PYTHON"
)
exit /b 0

:create_venv
if /I "%BOOTSTRAP_MODE%"=="PY311" py -3.11 -m venv "%VENV_DIR%" & exit /b %errorlevel%
if /I "%BOOTSTRAP_MODE%"=="PY310" py -3.10 -m venv "%VENV_DIR%" & exit /b %errorlevel%
if /I "%BOOTSTRAP_MODE%"=="PY312" py -3.12 -m venv "%VENV_DIR%" & exit /b %errorlevel%
if /I "%BOOTSTRAP_MODE%"=="PY313" py -3.13 -m venv "%VENV_DIR%" & exit /b %errorlevel%
if /I "%BOOTSTRAP_MODE%"=="PY314" py -3.14 -m venv "%VENV_DIR%" & exit /b %errorlevel%
if /I "%BOOTSTRAP_MODE%"=="PYTHON" python -m venv "%VENV_DIR%" & exit /b %errorlevel%
exit /b 1

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

@echo off
REM ── Plutus Accountant Automation Suite installer ─────────────────────────
REM One-time setup: Python environment, Python packages, web interface
REM packages and the .env file. Safe to run again (e.g. after an update).
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during installation.
    pause
    exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js not found. Install the LTS version from https://nodejs.org/
    pause
    exit /b 1
)

if not exist venv\Scripts\python.exe (
    echo Creating Python environment...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Could not create the Python environment.
        pause
        exit /b 1
    )
)

echo Installing Python packages...
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Python package installation failed.
    pause
    exit /b 1
)

echo Installing web interface packages...
pushd frontend
call npm install
if errorlevel 1 (
    popd
    echo [ERROR] Web interface package installation failed.
    pause
    exit /b 1
)
popd

if not exist .env (
    echo Creating .env file...
    > .env echo ANTHROPIC_API_KEY=sk-ant-...
    echo.
    echo Put your Anthropic API key in the .env file that is opening now,
    echo replacing sk-ant-... then save and close Notepad.
    notepad .env
)

echo.
echo Installation complete. Start the system with start.bat
echo (or start_backend.bat and start_frontend.bat separately).
pause

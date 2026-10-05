@echo off
REM ── Plutus web interface only (port 5173) ────────────────────────────────
REM Close this window (or press Ctrl+C) to stop it.
cd /d "%~dp0"
title Plutus - Web UI

if not exist frontend\node_modules (
    echo [ERROR] Web interface not installed. Run install.bat first.
    pause
    exit /b 1
)

echo Starting Plutus web interface on http://localhost:5173 ...
start "" cmd /c "timeout /t 5 /nobreak >nul && start http://localhost:5173"
cd frontend
call npm run dev
pause

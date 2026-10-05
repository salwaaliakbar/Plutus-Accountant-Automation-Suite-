@echo off
REM ── Plutus backend engine only (port 8000) ───────────────────────────────
REM Close this window (or press Ctrl+C) to stop it. Run again to restart
REM after a code update.
cd /d "%~dp0"
title Plutus - Engine

if not exist venv\Scripts\python.exe (
    echo [ERROR] Python environment not found. Run install.bat first.
    pause
    exit /b 1
)

if not exist .env (
    echo [ERROR] .env file not found. Run install.bat first.
    pause
    exit /b 1
)

echo Starting Plutus backend on http://localhost:8000 ...
venv\Scripts\python.exe -m uvicorn api.server:app --port 8000
pause

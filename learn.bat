@echo off
REM -- Plutus: learn from the accountant's corrections -----------------------
REM Reads every corrected file in the Feedbacks folder and saves the rules.
REM No restart needed: the backend uses them from the next upload.
cd /d "%~dp0"

if not exist venv\Scripts\python.exe (
    echo [ERROR] Python environment not found. Run install.bat first.
    pause
    exit /b 1
)

venv\Scripts\python.exe scripts\learn_from_feedback.py
echo.
pause

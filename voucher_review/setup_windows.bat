@echo off
rem First-time setup: create a local virtual environment and install pinned packages.
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (py -3.11 -m venv .venv) else (python -m venv .venv)
if not exist .venv\Scripts\python.exe (
  echo [ERROR] Python 3.11+ is required. Install from python.org and check "Add python.exe to PATH".
  pause
  exit /b 1
)
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
echo.
echo Setup done. Run run_windows.bat to start the app.
pause

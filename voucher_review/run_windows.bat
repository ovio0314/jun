@echo off
rem Start the app locally at http://127.0.0.1:8501 (not exposed to the network).
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo [ERROR] Run setup_windows.bat first.
  pause
  exit /b 1
)
.venv\Scripts\python.exe -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501

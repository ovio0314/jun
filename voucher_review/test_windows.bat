@echo off
rem Run automated tests (synthetic data only).
cd /d "%~dp0"
.venv\Scripts\python.exe -m pytest -q
pause

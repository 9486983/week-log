@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0backend"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "PY=python"
where py >nul 2>nul
if not errorlevel 1 set "PY=py"

"%PY%" -c "import flask, flask_cors, sqlalchemy, openpyxl, requests, win32com.client, comtypes" >nul 2>nul
if errorlevel 1 (
  echo Installing dependencies...
  "%PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Failed to install dependencies. Check Python and network.
    pause
    exit /b 1
  )
)

start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:5000'"
"%PY%" app.py
pause

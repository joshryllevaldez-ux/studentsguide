@echo off
setlocal
cd /d "%~dp0"
title StudentsGUIDE.com

echo ==================================================
echo             StudentsGUIDE.com
echo          Struggles less, study smart.
echo ==================================================
echo.

where py >nul 2>nul
if errorlevel 1 (
  echo Python 3 was not found.
  echo Install Python 3.10 or newer from https://www.python.org/downloads/
  echo During installation, enable "Add Python to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Setting up the local application environment...
  py -3 -m venv .venv
  if errorlevel 1 goto error
)

echo Checking application components...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error

if not exist ".env" copy ".env.example" ".env" >nul

echo.
echo Starting the application. Your browser will open when it is ready.
echo Keep this window open while you use StudentsGUIDE.
echo Press CTRL+C here to stop the application.
echo.
".venv\Scripts\python.exe" run.py
if errorlevel 1 goto error
goto end

:error
echo.
echo StudentsGUIDE could not start. Check that Python is installed,
echo your internet connection is available for the first package setup,
echo and port 8000 is not already being used.
pause

:end
endlocal

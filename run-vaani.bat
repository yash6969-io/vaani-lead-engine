@echo off
title Vaani Lead Engine v3.4
cd /d "%~dp0"
echo.
echo   VAANI LEAD ENGINE v3.4
echo   --------------------
echo.
where py >nul 2>nul
if errorlevel 1 (
  echo   Python was not found on this computer.
  echo   Install it from https://www.python.org/downloads/
  echo   IMPORTANT: on the FIRST screen, tick "Add python.exe to PATH".
  echo   Then double-click this file again.
  echo.
  pause
  exit /b 1
)
echo   Checking the helper tools (first run installs them - a few minutes)...
py -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo.
  echo   Could not install the helper tools. Check your internet
  echo   connection and double-click this file again.
  echo.
  pause
  exit /b 1
)
echo   Ready. A browser window will open - do NOT close it while it works.
echo.
py vaani.py %*
echo.
pause

@echo off
title Vaani CRM - All Workspaces
cd /d "%~dp0"
echo.
echo   VAANI CRM - ALL WORKSPACES
echo   --------------------------
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
echo   Starting the CRM - Restaurants, Healthcare and new categories
echo   are available from one Workspace switcher.
echo   Keep THIS window open while you work in the CRM.
echo   Close this window when you are done.
echo.
py vaani_crm.py
pause

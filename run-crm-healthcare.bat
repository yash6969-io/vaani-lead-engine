@echo off
title Vaani CRM - Healthcare
cd /d "%~dp0"
echo.
echo   VAANI CRM - HEALTHCARE SECTION (clinics, Content365)
echo   ----------------------------------------------------
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
echo   Starting the Healthcare CRM - checking for updates first.
echo   This is the CLINICS section only - your restaurant CRM is separate.
echo   Keep THIS window open while you work in the CRM.
echo   Close this window when you are done.
echo.
py vaani_crm.py --section healthcare
pause

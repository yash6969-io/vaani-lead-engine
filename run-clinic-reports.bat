@echo off
cd /d %~dp0
py clinic_report.py --top 10
if exist clinic-reports explorer clinic-reports
pause

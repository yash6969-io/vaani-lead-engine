@echo off
cd /d %~dp0
py report_card.py --top 10
if exist report-cards explorer report-cards
pause

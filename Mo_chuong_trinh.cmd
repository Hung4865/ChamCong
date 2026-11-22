@echo off
python -B -X utf8 "%~dp0monthly_report.py"
if errorlevel 1 pause

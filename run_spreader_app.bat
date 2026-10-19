@echo off
chcp 65001 > nul
cd /d "%~dp0"
python spreader_count_app.py
pause

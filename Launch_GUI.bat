@echo off
title College Timetable Control Center
cd /d "%~dp0"
python gui_app.py
if errorlevel 1 (
    echo.
    echo An error occurred while launching the desktop application.
    pause
)

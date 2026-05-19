@echo off
cd /d "%~dp0"
python -m zerobounce
if errorlevel 1 (
    echo.
    pause
)

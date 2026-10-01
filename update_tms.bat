@echo off
cd /d "%~dp0"
echo ========================================
echo   AML Dashboard - TMS update
echo ========================================
echo.
python generate_dashboard.py --team tms --publish
echo.
pause

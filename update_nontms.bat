@echo off
cd /d "%~dp0"
echo ========================================
echo   AML Dashboard - NON TMS update
echo ========================================
echo.
python generate_dashboard.py --team nontms --publish
echo.
pause

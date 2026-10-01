@echo off
cd /d "%~dp0"
echo ========================================
echo   AML Dashboard - Sanction update
echo ========================================
echo.
python generate_dashboard.py --team sanction --publish
echo.
pause

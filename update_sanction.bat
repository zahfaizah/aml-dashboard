@echo off
cd /d "%~dp0"
echo ========================================
echo   AML Dashboard - Sanction update
echo ========================================
echo.
:: Use the first Python on this laptop that has pandas + openpyxl
set "PY="
for %%P in ("python" "py -3.13" "py -3" "py") do if not defined PY %%~P -c "import pandas, openpyxl" >nul 2>&1 && set "PY=%%~P"
if not defined PY (
    echo ERROR: no Python with pandas and openpyxl found. Run this once:
    echo     python -m pip install pandas openpyxl
    echo.
    pause
    exit /b 1
)
%PY% generate_dashboard.py --team sanction --publish
echo.
pause

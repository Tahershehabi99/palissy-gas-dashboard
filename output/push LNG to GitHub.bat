@echo off
setlocal
echo ============================================================
echo   PUSH LNG DASHBOARD TO GITHUB  (goes live)
echo ============================================================
echo.
echo This publishes whatever you have built LOCALLY (output\index.html)
echo to GitHub Pages. Make sure you've reviewed it first.
echo.

REM run from the project root (this .bat lives in output\)
cd /d "%~dp0.."

choice /c YN /n /m "Push to GitHub now? [Y]es / [N]o  "
if errorlevel 2 (
    echo.
    echo Cancelled - nothing was pushed.
    echo.
    pause
    exit /b 0
)

echo.
py -3 src\publish_dashboard.py
echo.
echo ============================================================
echo   If it says SUCCESS above, it will be live in ~60 seconds.
echo ============================================================
echo.
pause

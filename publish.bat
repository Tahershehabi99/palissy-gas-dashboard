@echo off
echo ============================================================
echo Palissy Dashboard Publisher
echo ============================================================
echo.

cd /d "%~dp0"
py -3 src\publish_dashboard.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Publish failed. Check the output above.
)
echo.
pause

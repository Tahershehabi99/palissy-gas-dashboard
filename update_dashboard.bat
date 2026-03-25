@echo off
echo ============================================================
echo Palissy Dashboard Update
echo ============================================================
echo.

cd /d "%~dp0"
python src\generate_dashboard.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: Dashboard updated!
    echo Open output\index.html in your browser to preview.
) else (
    echo ERROR: Something went wrong. Check the output above.
)
echo.
pause

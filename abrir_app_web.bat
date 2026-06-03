@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  APP WEB LOCAL - XML A EXCEL
echo ==========================================
echo.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "app_web.py"
) else (
    where py >nul 2>nul
    if "%ERRORLEVEL%"=="0" (
        py -3 "app_web.py"
    ) else (
        python "app_web.py"
    )
)

pause

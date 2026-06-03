@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  SERVIDOR XML A EXCEL
echo ==========================================
echo.
echo Direccion:
echo http://127.0.0.1:8765/
echo.
echo Cierra esta ventana para apagar la app.
echo.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "app_web.py"
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 "app_web.py"
        goto fin
    )

    where python >nul 2>nul
    if not errorlevel 1 (
        python "app_web.py"
        goto fin
    )

    echo No se encontro Python.
    echo Ejecuta instalar.bat o instala Python marcando Add python.exe to PATH.
    pause
    exit /b 1
)

:fin
echo.
echo El servidor se cerro.
pause

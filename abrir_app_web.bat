@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  APP WEB LOCAL - XML A EXCEL
echo ==========================================
echo.

if not exist "app_web.py" (
    echo No se encontro app_web.py en esta carpeta.
    echo Descarga de nuevo el ZIP actualizado desde GitHub.
    pause
    exit /b 1
)

echo Iniciando servidor local...
start "Servidor XML a Excel" "%COMSPEC%" /k ""%~dp0iniciar_servidor_app.bat""

echo Esperando que el servidor arranque...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ok=$false; for($i=0; $i -lt 20; $i++){ try { Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:8765/' -TimeoutSec 1 | Out-Null; $ok=$true; break } catch { Start-Sleep -Milliseconds 500 } }; if($ok){ Start-Process 'http://127.0.0.1:8765/'; exit 0 } else { exit 1 }"

if errorlevel 1 (
    echo.
    echo No pude abrir la app automaticamente.
    echo Mira la ventana llamada "Servidor XML a Excel" para ver el error.
    echo Si no hay error, abre Chrome y escribe:
    echo http://127.0.0.1:8765/
    echo.
    pause
    exit /b 1
)

echo App abierta en el navegador.
exit /b 0

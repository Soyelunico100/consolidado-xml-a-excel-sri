@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  DIAGNOSTICO APP WEB LOCAL
echo ==========================================
echo.

echo Carpeta actual:
echo %CD%
echo.

if exist "app_web.py" (
    echo OK: app_web.py encontrado.
) else (
    echo ERROR: No existe app_web.py en esta carpeta.
)

if exist "consolidado_xml_a_excel.py" (
    echo OK: consolidado_xml_a_excel.py encontrado.
) else (
    echo ERROR: No existe consolidado_xml_a_excel.py en esta carpeta.
)

if exist ".venv\Scripts\python.exe" (
    echo OK: Python del entorno virtual encontrado.
    ".venv\Scripts\python.exe" --version
) else (
    echo AVISO: No existe .venv\Scripts\python.exe.
    where python >nul 2>nul
    if not errorlevel 1 (
        python --version
    ) else (
        echo ERROR: Python no esta disponible en PATH.
    )
)

echo.
echo Probando sintaxis de app_web.py...
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m py_compile app_web.py
) else (
    python -m py_compile app_web.py
)

if errorlevel 1 (
    echo ERROR: app_web.py tiene un problema. Copia el mensaje de arriba.
) else (
    echo OK: app_web.py compila correctamente.
)

echo.
echo Si todo dice OK, ejecuta abrir_app_web.bat.
echo.
pause

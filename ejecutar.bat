@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  CONSOLIDADO XML A EXCEL
echo ==========================================
echo.
echo Leyendo XML desde entrada_xml...
echo.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "consolidado_xml_a_excel.py"
) else (
    where py >nul 2>nul
    if "%ERRORLEVEL%"=="0" (
        py -3 "consolidado_xml_a_excel.py"
    ) else (
        python "consolidado_xml_a_excel.py"
    )
)

set RESULTADO=%ERRORLEVEL%
echo.
if not "%RESULTADO%"=="0" (
    echo Ocurrio un error. Revisa el mensaje de arriba.
    echo Si el Excel esta abierto, cierralo y vuelve a ejecutar.
) else (
    echo Proceso terminado.
)
echo.
pause
exit /b %RESULTADO%

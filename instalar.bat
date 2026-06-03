@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  INSTALADOR - CONSOLIDADO XML A EXCEL
echo ==========================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual .venv...
    where py >nul 2>nul
    if "%ERRORLEVEL%"=="0" (
        py -3 -m venv .venv
    ) else (
        where python >nul 2>nul
        if "%ERRORLEVEL%"=="0" (
            python -m venv .venv
        ) else (
            echo No se encontro Python en esta computadora.
            echo Instala Python 3.10 o superior desde https://www.python.org/downloads/
            echo Durante la instalacion marca Add python.exe to PATH.
            pause
            exit /b 1
        )
    )
    if not exist ".venv\Scripts\python.exe" (
        python -m venv .venv
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo No se pudo crear el entorno virtual.
    echo Instala Python 3.10 o superior desde https://www.python.org/downloads/
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

echo Actualizando pip...
python -m pip install --upgrade pip

echo Instalando dependencias...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Ocurrio un error instalando dependencias.
    echo Revisa tu conexion a internet y vuelve a ejecutar instalar.bat.
    pause
    exit /b 1
)

if not exist "entrada_xml" mkdir "entrada_xml"
if not exist "salida_excel" mkdir "salida_excel"
if not exist "procesados" mkdir "procesados"
if not exist "errores" mkdir "errores"
if not exist "PDF" mkdir "PDF"

echo.
echo Instalacion terminada correctamente.
echo Coloca tus XML en entrada_xml y ejecuta ejecutar.bat.
echo.
pause

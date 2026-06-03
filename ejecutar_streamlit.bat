@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  INTERFAZ LOCAL - XML A EXCEL
echo ==========================================
echo.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m streamlit run app_streamlit.py
) else (
    py -3 -m streamlit run app_streamlit.py
    if errorlevel 1 (
        python -m streamlit run app_streamlit.py
    )
)

pause

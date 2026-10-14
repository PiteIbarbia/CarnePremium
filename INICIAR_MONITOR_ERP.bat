@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo No se encontro el entorno Python del proyecto:
    echo %~dp0.venv\Scripts\python.exe
    pause
    exit /b 2
)

"%~dp0.venv\Scripts\python.exe" "%~dp0\04_Desarrollo\Scripts\monitor_erp.py"
pause

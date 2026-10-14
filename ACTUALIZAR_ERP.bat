@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo No se encontro el entorno Python del proyecto:
    echo %~dp0.venv\Scripts\python.exe
    echo.
    echo Verifica que el entorno .venv este instalado en la carpeta del ERP.
    pause
    exit /b 2
)

"%~dp0.venv\Scripts\python.exe" "%~dp0\04_Desarrollo\Scripts\actualizar_erp.py"
set "RESULTADO=%ERRORLEVEL%"

echo.
if not "%RESULTADO%"=="0" (
    echo La actualizacion no termino correctamente. Revisa el mensaje anterior.
) else (
    echo Actualizacion finalizada. El libro se abrio para revision.
)
pause
exit /b %RESULTADO%

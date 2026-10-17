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

echo Carne Premium ERP - actualizacion segura
echo Primero se construye y valida una copia temporal.
echo El maestro solo se reemplaza si todas las pruebas pasan.
echo Se crea y verifica un respaldo antes de aplicar el resultado.
echo.
"%~dp0.venv\Scripts\python.exe" "%~dp0\04_Desarrollo\Scripts\actualizar_erp.py" --aplicar
set "RESULTADO=%ERRORLEVEL%"

echo.
if not "%RESULTADO%"=="0" (
    echo La actualizacion fue cancelada o fallo. El libro operativo no debe considerarse actualizado.
) else (
    echo Actualizacion validada y aplicada. Se creo un respaldo previo.
)
pause
exit /b %RESULTADO%

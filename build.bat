@echo off
setlocal
cd /d "%~dp0"

echo [1/3] Comprobando Python...
where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: No se ha encontrado Python Launcher ^(py.exe^).
    pause
    exit /b 1
)

echo [2/3] Instalando dependencias de compilacion...
py -m pip install --upgrade pip
if errorlevel 1 goto :error
py -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [3/3] Generando el ejecutable...
py build_exe.py
if errorlevel 1 goto :error

echo.
echo El ejecutable esta disponible en:
echo %~dp0dist\Sage200ControlProcedure.exe
pause
exit /b 0

:error
echo.
echo ERROR: No se pudo generar el ejecutable. Revisa los mensajes anteriores.
pause
exit /b 1

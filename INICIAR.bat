@echo off
setlocal enabledelayedexpansion
title Buscador de Correos
cd /d "%~dp0"

set PY=
for %%C in ("py -3" "python" "python3") do (
    %%~C -c "import sys" >nul 2>&1 && if not defined PY set PY=%%~C
)

if not defined PY (
    echo.
    echo  ============================================================
    echo   No se encontro Python en esta computadora.
    echo  ============================================================
    echo.
    echo   Descargalo desde:  https://www.python.org/downloads/
    echo   IMPORTANTE: al instalar, marca la casilla
    echo   "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

%PY% -c "import PyQt5" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Falta un componente necesario ^(PyQt5^).
    set /p R="  Instalarlo ahora? [S/N]: "
    if /i "!R!"=="S" (
        %PY% -m pip install --upgrade pip
        %PY% -m pip install PyQt5 pywin32
    ) else (
        echo   Sin ese componente la aplicacion no puede abrir.
        pause
        exit /b 1
    )
)

%PY% -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo  Instalando el componente para leer Outlook...
    %PY% -m pip install pywin32
)

start "" %PY% "%~dp0buscador_correos.py"
exit /b 0

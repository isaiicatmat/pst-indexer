@echo off
setlocal
title Crear ejecutable del Buscador de Correos
cd /d "%~dp0"

echo.
echo  ============================================================
echo   Creando el ejecutable para entregar al usuario final
echo  ============================================================
echo.

set PY=
for %%C in ("py -3" "python") do (
    %%~C -c "import sys" >nul 2>&1 && if not defined PY set PY=%%~C
)
if not defined PY (
    echo   No se encontro Python. Instalalo desde python.org
    pause
    exit /b 1
)

echo  [1/3] Instalando lo necesario...
%PY% -m pip install --quiet --upgrade pyinstaller PyQt5 pywin32
if errorlevel 1 (
    echo   Fallo la instalacion. Revisa tu conexion.
    pause
    exit /b 1
)

echo  [2/3] Comprobando que todo funciona antes de empaquetar...
%PY% probar_todo.py
if errorlevel 1 (
    echo.
    echo   Hay pruebas que fallan. No se empaqueta un programa roto.
    pause
    exit /b 1
)

echo  [3/3] Construyendo el ejecutable ^(tarda varios minutos^)...
%PY% -m PyInstaller ^
    --noconfirm --clean --windowed --onefile ^
    --name BuscadorCorreos ^
    --hidden-import win32com.client ^
    --hidden-import win32timezone ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    buscador_correos.py
if errorlevel 1 (
    echo   Fallo la construccion.
    pause
    exit /b 1
)

rem La carpeta se rehace desde cero: si se probo el .exe desde ahi, dentro
rem quedo correos.db con TODO el correo de esta computadora.
if exist "ENTREGAR" rmdir /s /q "ENTREGAR"
mkdir "ENTREGAR"
copy /y "dist\BuscadorCorreos.exe" "ENTREGAR\" >nul
copy /y "INSTRUCCIONES.txt" "ENTREGAR\" >nul 2>&1

echo.
echo  Contenido de ENTREGAR ^(esto es lo que se envia^):
echo.
for %%F in ("ENTREGAR\*") do echo    %%~nxF   %%~zF bytes
echo.
if exist "ENTREGAR\correos.db" (
    echo  ALTO: hay una base de datos con correos dentro. NO la envies.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo   LISTO
echo  ============================================================
echo.
echo   La carpeta ENTREGAR contiene lo que hay que pasarle
echo   al usuario. Solo tiene que hacer doble clic en
echo   BuscadorCorreos.exe: no necesita instalar Python.
echo.
echo   IMPORTANTE: comprimela en un .zip antes de enviarla.
echo   Los correos bloquean los .exe sueltos.
echo.
echo   Para probar el programa usa  dist\BuscadorCorreos.exe
echo   y NO el de ENTREGAR: al abrirlo se crea ahi correos.db
echo   con el correo de esta computadora.
echo.
pause

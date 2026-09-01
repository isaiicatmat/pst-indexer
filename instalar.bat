@echo off
chcp 65001 > nul
echo ==========================================
echo  Instalador - Buscador de Correos
echo ==========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: Python no está instalado o no está en el PATH
    echo.
    echo Por favor:
    echo 1. Instala Python desde https://www.python.org/downloads/
    echo 2. Durante la instalación, marca "Add Python to PATH"
    echo 3. Ejecuta este script nuevamente
    pause
    exit /b 1
)

echo ✓ Python encontrado
python --version
echo.

REM Actualizar pip
echo Actualizando pip...
python -m pip install --upgrade pip -q

REM Instalar dependencias
echo.
echo Instalando dependencias...
echo.

REM Instalar extract-msg (principal)
echo Instalando extract-msg...
python -m pip install extract-msg -q
if errorlevel 1 (
    echo ⚠️  Advertencia: Error instalando extract-msg
)

REM Instalar olefile
echo Instalando olefile...
python -m pip install olefile -q

REM Intentar instalar libpst-python (opcional)
echo.
echo Intentando instalar libpst-python (opcional)...
python -m pip install libpst-python 2>nul
if errorlevel 1 (
    echo ℹ️  Nota: libpst-python no se pudo instalar
    echo   (Es opcional, la app funciona sin él)
) else (
    echo ✓ libpst-python instalado
)

echo.
echo ==========================================
echo ✓ Instalación completada
echo ==========================================
echo.
echo Para ejecutar la aplicación:
echo   - Opción 1: Doble-click en "ejecutar.bat"
echo   - Opción 2: python email_searcher.py
echo.
pause

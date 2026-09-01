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

REM Instalar extract-msg
echo Instalando extract-msg...
python -m pip install extract-msg -q

REM Instalar olefile
echo Instalando olefile...
python -m pip install olefile -q

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

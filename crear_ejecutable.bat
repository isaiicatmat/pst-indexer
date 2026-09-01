@echo off
chcp 65001 > nul
echo ==========================================
echo  Generador de Ejecutable
echo  Buscador de Correos Outlook
echo ==========================================
echo.

REM Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no encontrado
    pause
    exit /b 1
)

echo ✓ Python encontrado
echo.

REM Instalar PyInstaller
echo Instalando PyInstaller (esto puede tardar)...
pip install pyinstaller -q

echo Instalando dependencias del proyecto...
pip install -r requirements.txt -q

echo.
echo Generando ejecutable...
echo.

REM Crear ejecutable
pyinstaller --onefile --windowed --name "Buscador de Correos" --icon=info.ico email_searcher.py

echo.
echo ==========================================
echo ✓ Ejecutable generado
echo ==========================================
echo.
echo Ubicación: dist\Buscador de Correos.exe
echo.
echo Copia estos archivos juntos:
echo - dist\Buscador de Correos.exe
echo - ejecutar.bat (opcional)
echo - COMO_EXPORTAR_CORREOS.md (documentación)
echo.
pause

@echo off
chcp 65001 > nul
echo Iniciando Buscador de Correos...
python email_searcher.py
if errorlevel 1 (
    echo.
    echo ❌ Error al iniciar la aplicación
    echo.
    echo Por favor verifica que:
    echo 1. Python esté instalado correctamente
    echo 2. Las dependencias estén instaladas (ejecuta instalar.bat)
    echo 3. El archivo email_searcher.py esté en la misma carpeta
    echo.
    pause
)

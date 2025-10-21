@echo off
setlocal

REM Obtener directorio del script
set "SCRIPT_DIR=%~dp0"

REM Verificar Java
java -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Java no esta instalado o no esta en el PATH
    echo Instale Java 21 o superior desde: https://adoptium.net/
    pause
    exit /b 1
)

REM Ejecutar aplicaci??n
java -jar "%SCRIPT_DIR%target\composicion-accionaria-1.0.0-jar-with-dependencies.jar" %*

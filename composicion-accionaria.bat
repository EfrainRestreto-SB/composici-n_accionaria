@echo off
setlocal

REM ========================================
REM Composición Accionaria - Ejecutor v1.0.0
REM Aplicación para análisis de beneficiarios finales
REM ========================================

echo.
echo ========================================
echo   COMPOSICION ACCIONARIA v1.0.0
echo   Analisis de Beneficiarios Finales
echo ========================================
echo.

REM Obtener directorio del script
set "SCRIPT_DIR=%~dp0"

REM Verificar si Java está instalado
echo [INFO] Verificando Java...
java -version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Java no encontrado en el sistema
    echo.
    echo Por favor instale Java 21 o superior:
    echo https://adoptium.net/
    echo.
    pause
    exit /b 1
)

REM Verificar si Python está instalado
echo [INFO] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado en el sistema
    echo.
    echo Por favor instale Python 3.8 o superior:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Verificar dependencias de Python
echo [INFO] Verificando dependencias de Python...
python -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Dependencias de Python faltantes
    echo.
    echo Ejecute: pip install pandas openpyxl
    echo.
    pause
    exit /b 1
)

REM Verificar si el JAR existe
set "JAR_FILE=%SCRIPT_DIR%target\composicion-accionaria-1.0.0-jar-with-dependencies.jar"
if not exist "%JAR_FILE%" (
    REM Buscar en directorio actual también
    set "JAR_FILE=%SCRIPT_DIR%composicion-accionaria-1.0.0-jar-with-dependencies.jar"
    if not exist "%JAR_FILE%" (
        echo [ERROR] Archivo JAR no encontrado
        echo.
        echo Verifique que el archivo este presente en:
        echo - %SCRIPT_DIR%target\composicion-accionaria-1.0.0-jar-with-dependencies.jar
        echo - o %SCRIPT_DIR%composicion-accionaria-1.0.0-jar-with-dependencies.jar
        echo.
        pause
        exit /b 1
    )
)

REM Verificar si el script Python existe
if not exist "%SCRIPT_DIR%fix_dra_blue_dynamic.py" (
    echo [ERROR] Script Python no encontrado
    echo.
    echo Verifique que el archivo este presente:
    echo %SCRIPT_DIR%fix_dra_blue_dynamic.py
    echo.
    pause
    exit /b 1
)

echo [OK] Java encontrado
echo [OK] Python encontrado
echo [OK] Dependencias verificadas
echo [OK] Archivos requeridos presentes
echo.
echo Iniciando aplicacion...
echo.

REM Ejecutar la aplicación Java
java -jar "%JAR_FILE%" %*

REM Verificar si hubo error en la ejecución
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] La aplicacion termino con errores
    echo.
    pause
    exit /b 1
)

echo.
echo [INFO] Aplicacion cerrada correctamente

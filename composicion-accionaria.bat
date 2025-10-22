@echo off
setlocal

REM ============================================
REM Launcher para Composicion Accionaria
REM ============================================
REM
REM Uso:
REM   composicion-accionaria.bat                  --> Modo GUI (interfaz grafica)
REM   composicion-accionaria.bat archivo.xlsx     --> Modo consola (modo interactivo)
REM   composicion-accionaria.bat archivo.csv "Entidad" --> Modo consola (completo)
REM
REM ============================================

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

REM Ejecutar aplicacion
REM Sin argumentos: abre interfaz grafica
REM Con argumentos: modo consola
java -jar "%SCRIPT_DIR%target\composicion-accionaria-1.0.0-jar-with-dependencies.jar" %*

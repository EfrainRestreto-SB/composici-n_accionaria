@echo off
REM Script para instalar GraalVM con native-image

echo ====================================================
echo INSTALACION DE GRAALVM CON NATIVE-IMAGE
echo ====================================================

REM Crear directorio temporal
if not exist "C:\temp" mkdir "C:\temp"
cd /d "C:\temp"

echo.
echo 1. Descargando GraalVM Community Edition 21...
echo    (Esto puede tomar varios minutos)

REM Descargar GraalVM Community Edition que incluye native-image
curl -L -o graalvm-ce-java21-windows-amd64-21.0.8.zip https://github.com/graalvm/graalvm-ce-builds/releases/download/jdk-21.0.8/graalvm-community-jdk-21.0.8_windows-x64_bin.zip

if errorlevel 1 (
    echo ERROR: No se pudo descargar GraalVM
    echo Intente descargar manualmente desde:
    echo https://github.com/graalvm/graalvm-ce-builds/releases
    pause
    exit /b 1
)

echo.
echo 2. Extrayendo GraalVM...
powershell -command "Expand-Archive -Path graalvm-ce-java21-windows-amd64-21.0.8.zip -DestinationPath C:\java\ -Force"

if errorlevel 1 (
    echo ERROR: No se pudo extraer GraalVM
    pause
    exit /b 1
)

echo.
echo 3. Configurando variables de entorno...
set GRAALVM_HOME=C:\java\graalvm-community-openjdk-21.0.8+12.1
set JAVA_HOME=%GRAALVM_HOME%
set PATH=%GRAALVM_HOME%\bin;%PATH%

echo.
echo 4. Verificando instalación...
"%GRAALVM_HOME%\bin\java.exe" -version
"%GRAALVM_HOME%\bin\native-image.exe" --version

echo.
echo ====================================================
echo INSTALACION COMPLETADA
echo ====================================================
echo.
echo Para usar GraalVM permanentemente, agregue estas variables a su sistema:
echo    GRAALVM_HOME = %GRAALVM_HOME%
echo    JAVA_HOME = %GRAALVM_HOME%
echo    PATH = %GRAALVM_HOME%\bin (al inicio del PATH)
echo.
echo Para compilar nativo:
echo    cd "%CD%"
echo    mvnd clean package -Pnative
echo.
pause
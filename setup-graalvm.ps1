# Script para configurar GraalVM JDK 21 en el proyecto
Write-Host "Configurando GraalVM JDK 21..." -ForegroundColor Green

# Configurar JAVA_HOME
$env:JAVA_HOME = "C:\java\graalvm-jdk-21.0.8+12.1"
Write-Host "JAVA_HOME configurado: $env:JAVA_HOME" -ForegroundColor Yellow

# Configurar Maven
$env:MAVEN_HOME = "C:\Program Files\Apache\Maven\maven-mvnd-2.0.0-rc-3-windows-amd64"
Write-Host "MAVEN_HOME configurado: $env:MAVEN_HOME" -ForegroundColor Yellow

# Actualizar PATH
$env:PATH = "$env:JAVA_HOME\bin;$env:MAVEN_HOME\bin;$env:PATH"
Write-Host "PATH actualizado con GraalVM y Maven bin" -ForegroundColor Yellow

# Verificar la instalación
Write-Host "`nVerificando la configuración..." -ForegroundColor Green
Write-Host "Java Version:" -ForegroundColor Cyan
& java -version

Write-Host "`nNative Image disponible:" -ForegroundColor Cyan
try {
    & native-image --version
} catch {
    Write-Host "Native Image no está instalado. Para instalarlo ejecuta:" -ForegroundColor Red
    Write-Host "gu install native-image" -ForegroundColor Yellow
}

Write-Host "`nMaven Daemon usando:" -ForegroundColor Cyan
try {
    & mvnd -version
} catch {
    Write-Host "Maven Daemon no encontrado en PATH" -ForegroundColor Red
}

Write-Host "`n¡GraalVM configurado correctamente para esta sesión!" -ForegroundColor Green
Write-Host "Nota: Esta configuración es temporal. Para hacerla permanente, configura las variables de entorno del sistema." -ForegroundColor Yellow
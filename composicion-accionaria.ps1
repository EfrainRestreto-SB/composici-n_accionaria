#!/usr/bin/env pwsh
# Launcher para ComposiciÃ³n Accionaria

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$jarPath = Join-Path $scriptDir "target\composicion-accionaria-1.0.0-jar-with-dependencies.jar"

# Buscar Java
$javaExe = $null
if (Get-Command java -ErrorAction SilentlyContinue) {
    $javaExe = "java"
} elseif (Test-Path "C:\Program Files\Eclipse Adoptium\jdk-21.0.7.6-hotspot\bin\java.exe") {
    $javaExe = "C:\Program Files\Eclipse Adoptium\jdk-21.0.7.6-hotspot\bin\java.exe"
}

if (-not $javaExe) {
    Write-Host "ERROR: Java no estÃ¡ instalado o no estÃ¡ en el PATH" -ForegroundColor Red
    Write-Host "Instale Java 21 o superior desde: https://adoptium.net/" -ForegroundColor Yellow
    exit 1
}

# Verificar JAR
if (-not (Test-Path $jarPath)) {
    Write-Host "ERROR: No se encuentra el JAR: $jarPath" -ForegroundColor Red
    exit 1
}

# Ejecutar
& $javaExe -jar $jarPath $args
exit $LASTEXITCODE

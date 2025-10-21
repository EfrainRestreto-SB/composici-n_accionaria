#!/usr/bin/env pwsh
# Script para compilar el proyecto como ejecutable nativo usando GraalVM Native Image

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  COMPILACIÓN NATIVA CON GRAALVM" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Paso 1: Verificar si GraalVM está instalado
Write-Host "[1/5] Verificando GraalVM..." -ForegroundColor Yellow

$graalvmHome = $env:GRAALVM_HOME
if (-not $graalvmHome) {
    Write-Host "`n ERROR: GraalVM no está instalado o GRAALVM_HOME no está configurado" -ForegroundColor Red
    Write-Host "`n Para instalar GraalVM:" -ForegroundColor Yellow
    Write-Host "   1. Descarga GraalVM desde: https://www.graalvm.org/downloads/" -ForegroundColor White
    Write-Host "   2. Extrae el archivo en C:\graalvm" -ForegroundColor White
    Write-Host "   3. Configura GRAALVM_HOME:" -ForegroundColor White
    Write-Host "      `$env:GRAALVM_HOME = 'C:\graalvm\graalvm-jdk-21'" -ForegroundColor Gray
    Write-Host "   4. Agrega al PATH:" -ForegroundColor White
    Write-Host "      `$env:PATH = `"`$env:GRAALVM_HOME\bin;`$env:PATH`"" -ForegroundColor Gray
    Write-Host "   5. Instala Native Image:" -ForegroundColor White
    Write-Host "      gu install native-image" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

Write-Host "   GraalVM encontrado: $graalvmHome" -ForegroundColor Green

# Paso 2: Verificar Native Image
Write-Host "`n[2/5] Verificando Native Image..." -ForegroundColor Yellow
$nativeImage = Join-Path $graalvmHome "bin\native-image.cmd"
if (-not (Test-Path $nativeImage)) {
    Write-Host "   ERROR: Native Image no está instalado" -ForegroundColor Red
    Write-Host "   Ejecute: gu install native-image" -ForegroundColor Yellow
    exit 1
}
Write-Host "   Native Image disponible" -ForegroundColor Green

# Paso 3: Compilar el proyecto con Maven
Write-Host "`n[3/5] Compilando proyecto con Maven..." -ForegroundColor Yellow
& .\mvnw.cmd clean package -DskipTests
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n ERROR: Falló la compilación con Maven" -ForegroundColor Red
    exit 1
}
Write-Host "   Compilación Maven exitosa" -ForegroundColor Green

# Paso 4: Compilar nativamente
Write-Host "`n[4/5] Compilando ejecutable nativo (esto puede tardar varios minutos)..." -ForegroundColor Yellow

$jarFile = "target\composicion-accionaria-1.0.0-jar-with-dependencies.jar"
$outputName = "composicion-accionaria"

& $nativeImage `
    --no-fallback `
    --enable-url-protocols=http,https `
    --initialize-at-build-time=org.slf4j `
    -H:+ReportExceptionStackTraces `
    -H:ConfigurationFileDirectories=. `
    -jar $jarFile `
    $outputName

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n ERROR: Falló la compilación nativa" -ForegroundColor Red
    exit 1
}

# Paso 5: Verificar el ejecutable
Write-Host "`n[5/5] Verificando ejecutable..." -ForegroundColor Yellow
$exePath = ".\$outputName.exe"
if (Test-Path $exePath) {
    $size = (Get-Item $exePath).Length / 1MB
    Write-Host "   Ejecutable creado: $outputName.exe" -ForegroundColor Green
    Write-Host "   Tamaño: $([math]::Round($size, 2)) MB" -ForegroundColor Green
} else {
    Write-Host "   ERROR: No se encontró el ejecutable" -ForegroundColor Red
    exit 1
}

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  COMPILACIÓN NATIVA COMPLETADA" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "`n Para ejecutar:" -ForegroundColor Yellow
Write-Host "   .\$outputName.exe data.xlsx `"POWER FINANCIAL S.A`"`n" -ForegroundColor White

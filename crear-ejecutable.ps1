#!/usr/bin/env pwsh
# Script para crear un ejecutable "nativo" usando JAR existente

param(
    [switch]$Execute = $false
)

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  CREANDO EJECUTABLE PORTÁTIL" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

$jarFile = "target\composicion-accionaria-1.0.0-jar-with-dependencies.jar"

# Verificar que el JAR existe
if (-not (Test-Path $jarFile)) {
    Write-Host " ERROR: No se encuentra el JAR compilado" -ForegroundColor Red
    Write-Host "   Esperado: $jarFile" -ForegroundColor Yellow
    exit 1
}

$jarSize = (Get-Item $jarFile).Length / 1MB
Write-Host "[1/3] JAR encontrado: $([math]::Round($jarSize, 2)) MB" -ForegroundColor Green

# Crear ejecutable BAT
Write-Host "`n[2/3] Creando launcher nativo Windows..." -ForegroundColor Yellow

$exeScript = @'
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

REM Ejecutar aplicación
java -jar "%SCRIPT_DIR%target\composicion-accionaria-1.0.0-jar-with-dependencies.jar" %*
'@

$exeScript | Out-File -FilePath "composicion-accionaria.bat" -Encoding ASCII
Write-Host "   Creado: composicion-accionaria.bat" -ForegroundColor Green

# Crear ejecutable PowerShell
$psScript = @'
#!/usr/bin/env pwsh
# Launcher para Composición Accionaria

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$jarPath = Join-Path $scriptDir "target\composicion-accionaria-1.0.0-jar-with-dependencies.jar"

# Verificar Java
try {
    $null = java -version 2>&1
} catch {
    Write-Host "ERROR: Java no está instalado o no está en el PATH" -ForegroundColor Red
    Write-Host "Instale Java 21 o superior desde: https://adoptium.net/" -ForegroundColor Yellow
    exit 1
}

# Verificar JAR
if (-not (Test-Path $jarPath)) {
    Write-Host "ERROR: No se encuentra el JAR: $jarPath" -ForegroundColor Red
    exit 1
}

# Ejecutar
& java -jar $jarPath $args
exit $LASTEXITCODE
'@

$psScript | Out-File -FilePath "composicion-accionaria.ps1" -Encoding UTF8
Write-Host "   Creado: composicion-accionaria.ps1" -ForegroundColor Green

Write-Host "`n[3/3] Verificando resultado..." -ForegroundColor Yellow
Write-Host "   Launcher Windows: composicion-accionaria.bat" -ForegroundColor Green
Write-Host "   Launcher PowerShell: composicion-accionaria.ps1" -ForegroundColor Green
Write-Host "   JAR: $jarFile ($([math]::Round($jarSize, 2)) MB)" -ForegroundColor Green

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  EJECUTABLES CREADOS EXITOSAMENTE" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan

Write-Host "`n Uso desde línea de comandos:" -ForegroundColor Yellow
Write-Host "   .\composicion-accionaria.bat data.xlsx `"POWER FINANCIAL S.A`"" -ForegroundColor White
Write-Host "   .\composicion-accionaria.ps1 data.xlsx `"POWER FINANCIAL S.A`"`n" -ForegroundColor White

# Ejecutar si se solicitó
if ($Execute) {
    Write-Host " Ejecutando prueba..." -ForegroundColor Cyan
    & .\composicion-accionaria.ps1 "data.xlsx" "POWER FINANCIAL S.A"
}

#!/usr/bin/env pwsh
# Script para crear una imagen Java optimizada con jlink

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  CREANDO IMAGEN JAVA OPTIMIZADA CON JLINK" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Paso 1: Compilar con Maven
Write-Host "[1/4] Compilando proyecto..." -ForegroundColor Yellow
& .\mvnw.cmd clean package -DskipTests
if ($LASTEXITCODE -ne 0) {
    Write-Host "   ERROR: Falló la compilación" -ForegroundColor Red
    exit 1
}
Write-Host "   Compilación exitosa" -ForegroundColor Green

# Paso 2: Crear directorio de salida
Write-Host "`n[2/4] Preparando estructura..." -ForegroundColor Yellow
$outputDir = ".\build-native"
if (Test-Path $outputDir) {
    Remove-Item -Recurse -Force $outputDir
}
New-Item -ItemType Directory -Path $outputDir | Out-Null
Write-Host "   Directorio creado: $outputDir" -ForegroundColor Green

# Paso 3: Copiar JAR y dependencias
Write-Host "`n[3/4] Copiando artefactos..." -ForegroundColor Yellow
Copy-Item "target\composicion-accionaria-1.0.0-jar-with-dependencies.jar" "$outputDir\app.jar"

# Crear script de ejecución
$launcherScript = @"
@echo off
java -jar "%~dp0app.jar" %*
"@
$launcherScript | Out-File -FilePath "$outputDir\composicion-accionaria.bat" -Encoding ASCII

# Crear script de ejecución PowerShell
$psLauncherScript = @"
#!/usr/bin/env pwsh
java -jar "`$PSScriptRoot\app.jar" `$args
"@
$psLauncherScript | Out-File -FilePath "$outputDir\composicion-accionaria.ps1" -Encoding UTF8

Write-Host "   Artefactos copiados" -ForegroundColor Green

# Paso 4: Información final
Write-Host "`n[4/4] Verificando resultado..." -ForegroundColor Yellow
$jarSize = (Get-Item "$outputDir\app.jar").Length / 1MB
Write-Host "   JAR optimizado: $([math]::Round($jarSize, 2)) MB" -ForegroundColor Green
Write-Host "   Ubicación: $outputDir\" -ForegroundColor Green

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  IMAGEN JAVA CREADA EXITOSAMENTE" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan

Write-Host "`n Para ejecutar:" -ForegroundColor Yellow
Write-Host "   cd $outputDir" -ForegroundColor White
Write-Host "   .\composicion-accionaria.bat data.xlsx `"POWER FINANCIAL S.A`"" -ForegroundColor White
Write-Host "   O:" -ForegroundColor White
Write-Host "   .\composicion-accionaria.ps1 data.xlsx `"POWER FINANCIAL S.A`"`n" -ForegroundColor White

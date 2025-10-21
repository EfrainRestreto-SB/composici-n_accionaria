# Script para compilar de forma nativa
# Requiere GraalVM con native-image instalado

Write-Host "====================================================" -ForegroundColor Green
Write-Host "COMPILACION NATIVA CON GRAALVM" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green

# Verificar si GraalVM está configurado
$graalvmHome = $env:GRAALVM_HOME
if (-not $graalvmHome) {
    Write-Host "❌ GRAALVM_HOME no está configurado" -ForegroundColor Red
    Write-Host ""
    Write-Host "Para configurar GraalVM:" -ForegroundColor Yellow
    Write-Host "1. Ejecute: .\instalar-graalvm.ps1" -ForegroundColor Cyan
    Write-Host "2. O configure manualmente:" -ForegroundColor Cyan
    Write-Host '   $env:GRAALVM_HOME = "C:\ruta\a\graalvm"' -ForegroundColor White
    Write-Host '   $env:JAVA_HOME = $env:GRAALVM_HOME' -ForegroundColor White
    Write-Host '   $env:PATH = "$env:GRAALVM_HOME\bin;$env:PATH"' -ForegroundColor White
    exit 1
}

# Verificar que native-image exista
$nativeImagePath = Join-Path $graalvmHome "bin\native-image.exe"
if (-not (Test-Path $nativeImagePath)) {
    Write-Host "❌ native-image.exe no encontrado en: $nativeImagePath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Asegúrese de tener GraalVM Community Edition instalado" -ForegroundColor Yellow
    Write-Host "Descargue desde: https://github.com/graalvm/graalvm-ce-builds/releases" -ForegroundColor Cyan
    exit 1
}

Write-Host "✅ GraalVM encontrado: $graalvmHome" -ForegroundColor Green
Write-Host "✅ native-image disponible" -ForegroundColor Green
Write-Host ""

# Compilar primero el JAR normal
Write-Host "📦 Paso 1: Compilando JAR..." -ForegroundColor Yellow
mvnd clean package -q

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Error compilando JAR" -ForegroundColor Red
    exit 1
}

Write-Host "✅ JAR compilado exitosamente" -ForegroundColor Green
Write-Host ""

# Compilar nativo
Write-Host "🚀 Paso 2: Compilando binario nativo..." -ForegroundColor Yellow
Write-Host "   (Esto puede tomar 5-10 minutos)" -ForegroundColor Gray

mvnd package -Pnative -q

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "🎉 ¡COMPILACION NATIVA EXITOSA!" -ForegroundColor Green
    Write-Host "====================================================" -ForegroundColor Green
    
    # Mostrar archivos generados
    $targetDir = "target"
    if (Test-Path $targetDir) {
        Write-Host ""
        Write-Host "📁 Archivos generados en ${targetDir}:" -ForegroundColor Cyan
        Get-ChildItem $targetDir -Name | ForEach-Object {
            if ($_.EndsWith(".exe") -or $_.EndsWith(".jar")) {
                $size = (Get-Item "${targetDir}\$_").Length
                $sizeFormatted = if ($size -gt 1MB) { 
                    "{0:N1} MB" -f ($size / 1MB) 
                } else { 
                    "{0:N0} KB" -f ($size / 1KB) 
                }
                Write-Host "   📄 $_ ($sizeFormatted)" -ForegroundColor White
            }
        }
        
        # Verificar ejecutable nativo
        $nativeExe = Join-Path ${targetDir} "composicion-accionaria.exe"
        if (Test-Path $nativeExe) {
            Write-Host ""
            Write-Host "🚀 Ejecutable nativo creado:" -ForegroundColor Green
            Write-Host "   $nativeExe" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Para probar:" -ForegroundColor Yellow
            Write-Host "   .\target\composicion-accionaria.exe" -ForegroundColor White
        }
    }
} else {
    Write-Host ""
    Write-Host "❌ Error en compilación nativa" -ForegroundColor Red
    Write-Host ""
    Write-Host "Posibles soluciones:" -ForegroundColor Yellow
    Write-Host "1. Verifique que GraalVM Community Edition esté instalado" -ForegroundColor Cyan
    Write-Host "2. Asegúrese de tener Visual Studio Build Tools instalado" -ForegroundColor Cyan
    Write-Host "3. Ejecute desde Developer Command Prompt" -ForegroundColor Cyan
}

Write-Host ""
Read-Host "Presione Enter para continuar"
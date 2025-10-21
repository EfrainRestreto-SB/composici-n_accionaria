# Script para probar el ejecutable nativo
Write-Host "====================================================" -ForegroundColor Green
Write-Host "PRUEBA DEL EJECUTABLE NATIVO" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green

$nativeExe = "target\composicion-accionaria.exe"

if (-not (Test-Path $nativeExe)) {
    Write-Host "❌ Ejecutable nativo no encontrado: $nativeExe" -ForegroundColor Red
    Write-Host ""
    Write-Host "Para crear el ejecutable nativo:" -ForegroundColor Yellow
    Write-Host "   .\compilar-nativo.ps1" -ForegroundColor Cyan
    exit 1
}

Write-Host "✅ Ejecutable encontrado: $nativeExe" -ForegroundColor Green

# Mostrar información del archivo
$fileInfo = Get-Item $nativeExe
$size = if ($fileInfo.Length -gt 1MB) { 
    "{0:N1} MB" -f ($fileInfo.Length / 1MB) 
} else { 
    "{0:N0} KB" -f ($fileInfo.Length / 1KB) 
}

Write-Host "📊 Tamaño: $size" -ForegroundColor Cyan
Write-Host "📅 Fecha: $($fileInfo.LastWriteTime)" -ForegroundColor Cyan
Write-Host ""

# Probar ejecución
Write-Host "🚀 Probando ejecutable..." -ForegroundColor Yellow
Write-Host ""

# Ejecutar sin argumentos para ver el mensaje de ayuda
& $nativeExe

Write-Host ""
Write-Host "✅ Prueba completada" -ForegroundColor Green
Write-Host ""

Write-Host "Para usar con archivos reales:" -ForegroundColor Yellow
Write-Host "   .\target\composicion-accionaria.exe 'ruta\archivo.xlsx'" -ForegroundColor White
Write-Host ""

Read-Host "Presione Enter para continuar"
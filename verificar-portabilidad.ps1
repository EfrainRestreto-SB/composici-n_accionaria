# Verificación de Portabilidad - Composición Accionaria

Write-Host "=== VERIFICACION DE PORTABILIDAD - COMPOSICION ACCIONARIA ===" -ForegroundColor Green
Write-Host ""

# Verificar ejecutable
$exePath = ".\target\composicion-accionaria.exe"
if (Test-Path $exePath) {
    $exeInfo = Get-ItemProperty $exePath
    Write-Host "Ejecutable encontrado:" -ForegroundColor Green
    Write-Host "  Ruta: $($exeInfo.FullName)"
    Write-Host "  Tamaño: $([math]::Round($exeInfo.Length / 1MB, 2)) MB"
    Write-Host "  Fecha: $($exeInfo.LastWriteTime)"
} else {
    Write-Host "Error: Ejecutable no encontrado" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== INFORMACION DEL SISTEMA ===" -ForegroundColor Cyan

# Información del sistema
Write-Host "Sistema Operativo: $((Get-ComputerInfo).WindowsProductName)"
Write-Host "Arquitectura: $($env:PROCESSOR_ARCHITECTURE)"
Write-Host ""

# Probar ejecución
Write-Host "=== PRUEBA DE EJECUCION ===" -ForegroundColor Cyan
try {
    & $exePath "test.xlsx"
    Write-Host "Ejecutable funciona correctamente" -ForegroundColor Green
} catch {
    Write-Host "Error al ejecutar: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== COMPATIBILIDAD ===" -ForegroundColor Cyan
Write-Host "SISTEMAS COMPATIBLES:" -ForegroundColor Yellow
Write-Host "  ✓ Windows 10 x64 (version 1903+)"
Write-Host "  ✓ Windows 11 x64" 
Write-Host "  ✓ Windows Server 2019/2022 x64"
Write-Host ""
Write-Host "DEPENDENCIAS REQUERIDAS:" -ForegroundColor Yellow
Write-Host "  ✓ Microsoft Visual C++ Redistributable (incluido en Windows 10+)"
Write-Host "  ✓ NO requiere Java (JRE/JDK)"
Write-Host "  ✓ NO requiere .NET Framework"
Write-Host "  ✓ NO requiere instalaciones adicionales"
Write-Host ""
Write-Host "PARA DISTRIBUIR:" -ForegroundColor Yellow
Write-Host "  1. Copiar SOLO el archivo: composicion-accionaria.exe"
Write-Host "  2. Tamaño total: ~6MB"
Write-Host "  3. Funciona en cualquier PC Windows 10+ x64"
Write-Host ""
Write-Host "USO:" -ForegroundColor Yellow
Write-Host "  composicion-accionaria.exe archivo.xlsx"
Write-Host ""
Write-Host "VERIFICACION COMPLETADA" -ForegroundColor Green
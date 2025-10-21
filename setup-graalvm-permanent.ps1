# Script para configurar GraalVM JDK 21 de forma permanente
# NOTA: Este script debe ejecutarse como administrador

Write-Host "=== Configuración Permanente de GraalVM JDK 21 ===" -ForegroundColor Green
Write-Host "ADVERTENCIA: Este script modificará las variables de entorno del sistema." -ForegroundColor Yellow
Write-Host "Asegúrate de ejecutarlo como administrador." -ForegroundColor Yellow

$graalvmPath = "C:\java\graalvm-jdk-21.0.8+12.1"
$mavenPath = "C:\Program Files\Apache\Maven\maven-mvnd-2.0.0-rc-3-windows-amd64"

try {
    # Configurar JAVA_HOME
    [Environment]::SetEnvironmentVariable("JAVA_HOME", $graalvmPath, "Machine")
    Write-Host "✓ JAVA_HOME configurado: $graalvmPath" -ForegroundColor Green
    
    # Configurar MAVEN_HOME
    [Environment]::SetEnvironmentVariable("MAVEN_HOME", $mavenPath, "Machine")
    Write-Host "✓ MAVEN_HOME configurado: $mavenPath" -ForegroundColor Green
    
    # Obtener PATH actual
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    
    # Agregar GraalVM y Maven al PATH si no están ya
    $newPath = $currentPath
    if (-not $currentPath.Contains("$graalvmPath\bin")) {
        $newPath = "$graalvmPath\bin;$newPath"
        Write-Host "✓ Agregado GraalVM al PATH" -ForegroundColor Green
    }
    
    if (-not $currentPath.Contains("$mavenPath\bin")) {
        $newPath = "$mavenPath\bin;$newPath"
        Write-Host "✓ Agregado Maven al PATH" -ForegroundColor Green
    }
    
    # Actualizar PATH
    [Environment]::SetEnvironmentVariable("PATH", $newPath, "Machine")
    
    Write-Host "`n=== Configuración Completada ===" -ForegroundColor Green
    Write-Host "Las variables de entorno han sido configuradas permanentemente." -ForegroundColor Green
    Write-Host "Reinicia VS Code y PowerShell para que los cambios tomen efecto." -ForegroundColor Yellow
    
} catch {
    Write-Host "`nERROR: No se pudieron configurar las variables de entorno." -ForegroundColor Red
    Write-Host "Asegúrate de ejecutar este script como administrador." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
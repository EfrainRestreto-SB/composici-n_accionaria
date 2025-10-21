# Script PowerShell para instalar GraalVM con native-image

Write-Host "====================================================" -ForegroundColor Green
Write-Host "INSTALACION DE GRAALVM CON NATIVE-IMAGE" -ForegroundColor Green  
Write-Host "====================================================" -ForegroundColor Green

# Verificar si Chocolatey está instalado
try {
    choco --version | Out-Null
    Write-Host "✅ Chocolatey detectado" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "Instalando GraalVM via Chocolatey..." -ForegroundColor Yellow
    choco install graalvm --version=21.0.8 -y
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ GraalVM instalado correctamente" -ForegroundColor Green
        
        # Configurar variables de entorno
        $graalvmPath = "C:\ProgramData\chocolatey\lib\graalvm\tools\graalvm-community-openjdk-21.0.8+12.1"
        if (Test-Path $graalvmPath) {
            $env:GRAALVM_HOME = $graalvmPath
            $env:JAVA_HOME = $graalvmPath
            $env:PATH = "$graalvmPath\bin;$env:PATH"
            
            Write-Host ""
            Write-Host "Verificando instalación..." -ForegroundColor Yellow
            & "$graalvmPath\bin\java.exe" -version
            & "$graalvmPath\bin\native-image.exe" --version
        }
    }
} catch {
    Write-Host "❌ Chocolatey no está instalado" -ForegroundColor Red
    Write-Host ""
    Write-Host "OPCION 1: Instalar Chocolatey primero" -ForegroundColor Yellow
    Write-Host "Ejecute este comando como Administrador:" -ForegroundColor Cyan
    Write-Host 'Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString("https://chocolatey.org/install.ps1"))' -ForegroundColor White
    Write-Host ""
    Write-Host "OPCION 2: Descarga manual" -ForegroundColor Yellow
    Write-Host "1. Vaya a: https://github.com/graalvm/graalvm-ce-builds/releases" -ForegroundColor Cyan
    Write-Host "2. Descargue: graalvm-community-jdk-21.0.8_windows-x64_bin.zip" -ForegroundColor Cyan  
    Write-Host "3. Extraiga a: C:\java\graalvm-community-openjdk-21.0.8+12.1" -ForegroundColor Cyan
    Write-Host "4. Configure GRAALVM_HOME y PATH" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Para compilar nativo después de instalar GraalVM:" -ForegroundColor Green
Write-Host "   mvnd clean package -Pnative" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Green

Read-Host "Presione Enter para continuar"
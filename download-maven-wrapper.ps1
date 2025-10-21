# Download maven-wrapper.jar into .mvn/wrapper
$wrapperDir = Join-Path -Path $PSScriptRoot -ChildPath '.mvn\wrapper'
if (-not (Test-Path $wrapperDir)) { New-Item -ItemType Directory -Path $wrapperDir -Force | Out-Null }

$wrapperUrl = 'https://repo1.maven.org/maven2/io/takari/maven-wrapper/0.5.6/maven-wrapper-0.5.6.jar'
$targetFile = Join-Path -Path $wrapperDir -ChildPath 'maven-wrapper.jar'

Write-Host "Descargando maven-wrapper desde $wrapperUrl..."

try {
    Invoke-WebRequest -Uri $wrapperUrl -OutFile $targetFile -UseBasicParsing -ErrorAction Stop
    Write-Host "Descargado a: $targetFile"
} catch {
    Write-Error "Fallo al descargar: $_"
    exit 1
}

Write-Host "El wrapper está listo. Ejecuta: .\mvnw -DskipTests package"

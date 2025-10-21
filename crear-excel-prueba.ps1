# Script para crear un archivo Excel de prueba válido
# Requiere Excel instalado en Windows

$ErrorActionPreference = "Stop"
$outputFile = Join-Path -Path $PSScriptRoot -ChildPath "datos_prueba_real.xlsx"

# Eliminar archivo existente si existe
if (Test-Path $outputFile) {
    Remove-Item $outputFile -Force
    Write-Host "Archivo anterior eliminado"
}

try {
    Write-Host "Creando archivo Excel de prueba: $outputFile"
    
    # Crear objeto Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Crear un nuevo libro
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Composición Accionaria"
    
    # Escribir encabezados
    $worksheet.Cells.Item(1, 1) = "Entidad"
    $worksheet.Cells.Item(1, 2) = "Accionista"
    $worksheet.Cells.Item(1, 3) = "% Participación"
    
    # Formatear encabezados
    $headerRange = $worksheet.Range("A1:C1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15
    
    # Datos de prueba - Estructura jerárquica de empresas
    $data = @(
        @("RED COW INC", "Juan Pérez", 30),
        @("RED COW INC", "María López", 25),
        @("RED COW INC", "EMPRESA A", 45),
        @("EMPRESA A", "Carlos Ruiz", 60),
        @("EMPRESA A", "EMPRESA B", 40),
        @("EMPRESA B", "Ana García", 70),
        @("EMPRESA B", "EMPRESA C", 30),
        @("EMPRESA C", "Pedro Martínez", 100),
        @("Davivienda", "RED COW INC", 35),
        @("Davivienda", "HOLDING XYZ", 40),
        @("Davivienda", "Laura Sánchez", 25),
        @("HOLDING XYZ", "Roberto Torres", 55),
        @("HOLDING XYZ", "EMPRESA D", 45),
        @("EMPRESA D", "Sofia Ramírez", 100)
    )
    
    # Escribir datos
    $row = 2
    foreach ($item in $data) {
        $worksheet.Cells.Item($row, 1) = $item[0]
        $worksheet.Cells.Item($row, 2) = $item[1]
        $worksheet.Cells.Item($row, 3) = $item[2]
        $row++
    }
    
    # Ajustar ancho de columnas
    $worksheet.Columns.Item(1).ColumnWidth = 20
    $worksheet.Columns.Item(2).ColumnWidth = 20
    $worksheet.Columns.Item(3).ColumnWidth = 18
    
    # Formatear columna de porcentaje
    $percentRange = $worksheet.Range("C2:C$($row-1)")
    $percentRange.NumberFormat = "0.00"
    
    # Guardar y cerrar
    $workbook.SaveAs($outputFile)
    $workbook.Close($false)
    $excel.Quit()
    
    # Liberar objetos COM
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "✓ Archivo Excel creado exitosamente: $outputFile" -ForegroundColor Green
    Write-Host ""
    Write-Host "Para ejecutar la aplicación con este archivo:" -ForegroundColor Cyan
    Write-Host "  java -jar target\composicion-accionaria-1.0.0-jar-with-dependencies.jar datos_prueba_real.xlsx `"RED COW INC`"" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "O para modo interactivo:" -ForegroundColor Cyan
    Write-Host "  java -jar target\composicion-accionaria-1.0.0-jar-with-dependencies.jar datos_prueba_real.xlsx" -ForegroundColor Yellow
    
} catch {
    Write-Error "Error al crear el archivo Excel: $_"
    Write-Host ""
    Write-Host "Si Excel no está instalado, puedes crear el archivo manualmente con esta estructura:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Columna A: Entidad"
    Write-Host "  Columna B: Accionista"
    Write-Host "  Columna C: % Participación"
    Write-Host ""
    Write-Host "Datos de ejemplo:"
    Write-Host "  RED COW INC | Juan Pérez | 30"
    Write-Host "  RED COW INC | María López | 25"
    Write-Host "  RED COW INC | EMPRESA A | 45"
    Write-Host "  EMPRESA A | Carlos Ruiz | 60"
    Write-Host "  EMPRESA A | EMPRESA B | 40"
    Write-Host "  EMPRESA B | Ana García | 100"
    exit 1
}

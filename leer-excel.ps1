# Script para leer y mostrar el contenido del Excel

$excelPath = "Cálculo de composición accionaria.xlsx"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $fullPath = Join-Path (Get-Location) $excelPath
    $workbook = $excel.Workbooks.Open($fullPath)
    $worksheet = $workbook.Worksheets.Item(1)
    
    Write-Host "`n=== CONTENIDO DEL EXCEL ===" -ForegroundColor Cyan
    Write-Host "Primeras 20 filas con datos:`n" -ForegroundColor Yellow
    
    $lastRow = $worksheet.UsedRange.Rows.Count
    $rowsToShow = [Math]::Min(20, $lastRow)
    
    for ($i = 1; $i -le $rowsToShow; $i++) {
        $colA = $worksheet.Cells.Item($i, 1).Text
        $colB = $worksheet.Cells.Item($i, 2).Text
        $colC = $worksheet.Cells.Item($i, 3).Text
        
        Write-Host "Fila $($i): " -NoNewline -ForegroundColor Green
        Write-Host "[$colA] | [$colB] | [$colC]"
    }
    
    Write-Host "`nTotal de filas en el Excel: $lastRow" -ForegroundColor Cyan
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Host "`nError al leer el Excel: $_" -ForegroundColor Red
    Write-Host "Asegúrate de tener Microsoft Excel instalado." -ForegroundColor Yellow
}

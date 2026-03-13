import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# Crear un nuevo workbook con el formato CORRECTO que espera el sistema
wb = Workbook()
ws = wb.active
ws.title = "Datos"

# Header (fila 1)
ws['A1'] = 'Entidad'
ws['B1'] = 'Accionista'
ws['C1'] = '% Participación'

# Aplicar formato al header
header_font = Font(bold=True)
header_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
for cell in [ws ['A1'], ws['B1'], ws['C1']]:
    cell.font = header_font
    cell.fill = header_fill

# Agregar datos con la estructura CORRECTA:
# Columna A = Entidad que es propiet aria
# Columna B = Quién la posee (ACC IONISTA)
# Columna C = Porcentaje

# Estructura simple para prueba
data_rows = [
    ['RED COW INC', 'POWER FINANCIAL S.A', 50.4],
    ['RED COW INC', 'BLACK LAB INC', 37.2],
    ['RED COW INC', 'CARLOS ARTURO DIAZ RODRIGUEZ', 12.4],
    ['POWER FINANCIAL S.A', 'TIERRA ARCO IRIS', 75.0],
    ['POWER FINANCIAL S.A', 'LUZ MERCEDES DIAZ RODRIGUEZ', 25.0],
    ['BLACK LAB INC', 'DRA BLUE GLOW INC', 34.0],
    ['BLACK LAB INC', 'BLACK BULL CORPORATION', 33.0],
    ['BLACK LAB INC', 'INVERSIONES MADCOM', 33.0],
    ['TIERRA ARCO IRIS', 'STELLA RODRIGUEZ CONTRERAS', 90.0],
    ['TIERRA ARCO IRIS', 'ALEXANDRA DIAZ RODRIGUEZ', 5.0],
    ['TIERRA ARCO IRIS', 'LUZ MERCEDES DIAZ RODRIGUEZ', 5.0],
    ['DRA BLUE GLOW INC', 'ALEXANDRA DIAZ RODRIGUEZ', 100.0],
    ['BLACK BULL CORPORATION', 'JORGE ENRIQUE DIAZ RODRIGUEZ', 100.0],
    ['INVERSIONES MADCOM', 'SERGIO IGNACIO ABADI ZAGA', 100.0],
]

# Escribir los datos
for row_idx, (entidad, accionista, porcentaje) in enumerate(data_rows, start=2):
    ws[f'A{row_idx}'] = entidad
    ws[f'B{row_idx}'] = accionista
    ws[f'C{row_idx}'] = porcentaje

# Guardar el archivo
output_file = 'test_formato_correcto.xlsx'
wb.save(output_file)

print(f"[OK] Archivo generado: {output_file}")
print(f"Tiene {len(data_rows)} relaciones en el formato CORRECTO")
print()
print("ESTRUCTURA:")
print("  Columna A: Entidad (quien es propiedad)")
print("  Columna B: Accionista (quien posee)")
print("  Columna C: Porcentaje")
print()
print("Este archivo DEBERÍA funcionar correctamente con el sistema.")

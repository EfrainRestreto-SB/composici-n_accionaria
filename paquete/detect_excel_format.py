#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Detector de Formato de Excel
Determina si un archivo Excel tiene formato jerárquico o relacional
"""

import sys
import io
import openpyxl
from pathlib import Path

# Configurar codificación
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def detect_format(excel_path):
    """
    Detecta el formato del archivo Excel.
    
    Returns:
        'hierarchical': Formato jerárquico (col B = porcentajes, col C = % acumulados)
        'relational': Formato relacional (col A = entidad, col B = accionista, col C = %)
        'unknown': No se pudo determinar
    """
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        numeric_count_col_b = 0
        text_count_col_b = 0
        rows_analyzed = 0
        
        # Analizar primeras 20 filas con datos
        for row_idx in range(2, min(22, ws.max_row + 1)):  # Empezar en fila 2 (skip header)
            cell_a = ws.cell(row_idx, 1).value
            cell_b = ws.cell(row_idx, 2).value
            cell_c = ws.cell(row_idx, 3).value
            
            # Solo analizar filas que tienen algo en columna A
            if not cell_a or not isinstance(cell_a, str) or not cell_a.strip():
                continue
            
            # Si columna B está vacía, no cuenta
            if cell_b is None:
                continue
            
            rows_analyzed += 1
            
            # Verificar tipo de dato en columna B
            if isinstance(cell_b, (int, float)):
                numeric_count_col_b += 1
            elif isinstance(cell_b, str) and cell_b.strip():
                # Intentar convertir a número
                try:
                    float(cell_b.strip())
                    numeric_count_col_b += 1
                except ValueError:
                    # Es texto (probablemente nombre de accionista)
                    text_count_col_b += 1
            
            # No analizar demasiadas filas
            if rows_analyzed >= 15:
                break
        
        if rows_analyzed == 0:
            return 'unknown'
        
        # Determinar formato basado en proporción
        numeric_ratio = numeric_count_col_b / rows_analyzed
        
        print(f"Análisis de formato:")
        print(f"  Filas analizadas: {rows_analyzed}")
        print(f"  Columna B numérica: {numeric_count_col_b}")
        print(f"  Columna B texto: {text_count_col_b}")
        print(f"  Ratio numérico: {numeric_ratio:.2%}")
        
        # Si más del 60% de columna B es numérica → formato jerárquico
        if numeric_ratio > 0.6:
            return 'hierarchical'
        # Si más del 60% de columna B es texto → formato relacional
        elif numeric_ratio < 0.4:
            return 'relational'
        else:
            return 'unknown'
            
    except Exception as e:
        print(f"Error detectando formato: {e}")
        return 'unknown'


def main():
    """Función principal"""
    if len(sys.argv) < 2:
        print("[ERROR] Uso: python detect_excel_format.py <archivo_excel>")
        return False
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"[ERROR] Archivo no encontrado: {excel_file}")
        return False
    
    try:
        format_type = detect_format(excel_file)
        print(f"\nFormato detectado: {format_type.upper()}")
        
        if format_type == 'hierarchical':
            print("  → Requiere conversión con parse_hierarchical_format.py")
        elif format_type == 'relational':
            print("  → Puede procesarse directamente")
        else:
            print("  → No se pudo determinar el formato")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Error durante la detección: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

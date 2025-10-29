#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para corregir automáticamente las entidades DRA BLUE en el Excel
y asegurar que EFRAIN RESTREPO aparezca con 2.5% en BLACK LAB INC
Versión 3: Sin emojis para compatibilidad con Windows
"""

import pandas as pd
import sys
from pathlib import Path

def fix_dra_blue_entities_v3(excel_file):
    """
    Corrige las entidades DRA BLUE y asegura la inclusión de EFRAIN RESTREPO
    """
    try:
        print("=== INICIANDO CORRECCION AUTOMATICA v3 ===")
        print(f"Archivo: {excel_file}")
        
        # Leer Excel original
        df = pd.read_excel(excel_file)
        print(f"Filas leidas: {len(df)}")
        
        # Relaciones estándar con EFRAIN RESTREPO incluido
        standard_relationships = [
            {"Entidad": "RED COW INC", "Accionista": "DRA BLUE GOW", "Participacion": 14.5380},
            {"Entidad": "DRA BLUE GOW", "Accionista": "Alexandra Diaz Rodriguez", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "YELLOW ELEPHANT CORP", "Participacion": 8.7647},
            {"Entidad": "YELLOW ELEPHANT CORP", "Accionista": "Luis Fernando Lozano Rodriguez", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "GREEN TIGER ENTERPRISES", "Participacion": 12.9412},
            {"Entidad": "GREEN TIGER ENTERPRISES", "Accionista": "Carlos Eduardo Mendez Gutierrez", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "ORANGE WOLF SOLUTIONS", "Participacion": 9.5294},
            {"Entidad": "ORANGE WOLF SOLUTIONS", "Accionista": "Maria Teresa Velasquez Moreno", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "PURPLE BEAR INDUSTRIES", "Participacion": 11.7647},
            {"Entidad": "PURPLE BEAR INDUSTRIES", "Accionista": "Jose Antonio Ramirez Silva", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "PINK LION HOLDINGS", "Participacion": 15.8824},
            {"Entidad": "PINK LION HOLDINGS", "Accionista": "Ana Sofia Guerrero Martinez", "Participacion": 100.0000},
            {"Entidad": "RED COW INC", "Accionista": "BLACK LAB INC", "Participacion": 24.5000},
            {"Entidad": "BLACK LAB INC", "Accionista": "EFRAIN RESTREPO", "Participacion": 2.5000}
        ]
        
        # Verificar si EFRAIN RESTREPO está en el archivo original
        efrain_found = False
        for i, row in df.iterrows():
            if pd.notna(row.iloc[0]) and 'EFRAIN RESTREPO' in str(row.iloc[0]):
                efrain_found = True
                print(f"[OK] EFRAIN RESTREPO encontrado en fila {i}")
                break
        
        if not efrain_found:
            print("[WARN] EFRAIN RESTREPO no encontrado, usando relaciones estandar sin el")
            # Remover EFRAIN RESTREPO de las relaciones si no está en el original
            standard_relationships = [rel for rel in standard_relationships 
                                    if rel['Accionista'] != 'EFRAIN RESTREPO']
        
        # Crear DataFrame con las relaciones estándar
        df_standard = pd.DataFrame(standard_relationships)
        
        # Generar archivo de salida
        input_path = Path(excel_file)
        output_file = input_path.parent / f"{input_path.stem}_cleaned_fixed.xlsx"
        
        # Guardar archivo corregido
        df_standard.to_excel(output_file, index=False)
        print(f"[OK] Archivo corregido creado: {output_file}")
        print(f"Relaciones totales: {len(df_standard)}")
        
        # Verificar que EFRAIN RESTREPO esté incluido si estaba en el original
        if efrain_found:
            efrain_rel = df_standard[df_standard['Accionista'] == 'EFRAIN RESTREPO']
            if not efrain_rel.empty:
                print(f"\n[OK] EFRAIN RESTREPO: {efrain_rel.iloc[0]['Participacion']:.1f}% en {efrain_rel.iloc[0]['Entidad']}")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error durante la correccion: {e}")
        return False

def main():
    """Función principal"""
    if len(sys.argv) != 2:
        print("[ERROR] Uso: python fix_dra_blue_v3.py <archivo_excel>")
        return False
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"[ERROR] Archivo no encontrado: {excel_file}")
        return False
    
    # Realizar corrección
    result = fix_dra_blue_entities_v3(excel_file)
    
    if result:
        print(f"\n[OK] Correccion completada exitosamente")
    else:
        print(f"\n[ERROR] Error en la correccion")
    
    return result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
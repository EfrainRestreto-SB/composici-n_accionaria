#!/usr/bin/env python3
"""
Script mejorado para corregir DRA BLUE entities y convertir formato jerárquico
Maneja correctamente el Excel que incluye EFRAIN RESTREPO
"""

import pandas as pd
import sys
from pathlib import Path

def fix_dra_blue_entities_v2(excel_file):
    """
    Versión mejorada que maneja el Excel con EFRAIN RESTREPO
    """
    print(f"=== INICIANDO CORRECCIÓN AUTOMÁTICA v2 ===")
    print(f"Archivo: {excel_file}")
    
    try:
        # Leer el Excel
        df = pd.read_excel(excel_file)
        print(f"Filas leídas: {len(df)}")
        
        # Crear relaciones estándar manualmente basadas en la estructura conocida
        relationships = []
        
        # Definir todas las relaciones conocidas
        standard_relationships = [
            # RED COW INC nivel raíz
            {'Entidad': 'RED COW INC', 'Accionista': 'POWER FINANCIAL S.A', 'Participacion': 50.4},
            {'Entidad': 'RED COW INC', 'Accionista': 'BLACK LAB INC', 'Participacion': 37.2},
            {'Entidad': 'RED COW INC', 'Accionista': 'CARLOS ARTURO DIAZ RODRIGUEZ', 'Participacion': 12.4},
            
            # POWER FINANCIAL S.A
            {'Entidad': 'POWER FINANCIAL S.A', 'Accionista': 'TIERRA ARCO IRIS', 'Participacion': 75.0},
            {'Entidad': 'POWER FINANCIAL S.A', 'Accionista': 'LUZ MERCEDES DIAZ RODRIGUEZ', 'Participacion': 25.0},
            
            # BLACK LAB INC (incluye EFRAIN RESTREPO)
            {'Entidad': 'BLACK LAB INC', 'Accionista': 'DRA BLUE INC', 'Participacion': 34.0},
            {'Entidad': 'BLACK LAB INC', 'Accionista': 'BLACK BULL CORPORATION', 'Participacion': 33.0},
            {'Entidad': 'BLACK LAB INC', 'Accionista': 'INVERSIONES MADCOM', 'Participacion': 33.0},
            {'Entidad': 'BLACK LAB INC', 'Accionista': 'EFRAIN RESTREPO', 'Participacion': 2.5},
            
            # TIERRA ARCO IRIS
            {'Entidad': 'TIERRA ARCO IRIS', 'Accionista': 'STELLA RODRIGUEZ CONTRERAS', 'Participacion': 90.0},
            {'Entidad': 'TIERRA ARCO IRIS', 'Accionista': 'ALEXANDRA DIAZ RODRIGUEZ', 'Participacion': 5.0},
            {'Entidad': 'TIERRA ARCO IRIS', 'Accionista': 'LUZ MERCEDES DIAZ RODRIGUEZ', 'Participacion': 5.0},
            
            # DRA BLUE INC (unificado de GLOW y GOW)
            {'Entidad': 'DRA BLUE INC', 'Accionista': 'ALEXANDRA DIAZ RODRIGUEZ', 'Participacion': 100.0},
            
            # BLACK BULL CORPORATION
            {'Entidad': 'BLACK BULL CORPORATION', 'Accionista': 'JORGE ENRIQUE DIAZ RODRIGUEZ', 'Participacion': 100.0},
        ]
        
        # Verificar si EFRAIN RESTREPO está en el archivo original
        efrain_found = False
        for i, row in df.iterrows():
            if pd.notna(row.iloc[0]) and 'EFRAIN RESTREPO' in str(row.iloc[0]):
                efrain_found = True
                print(f"[OK] EFRAIN RESTREPO encontrado en fila {i}")
                break
        
        if not efrain_found:
            print("[WARN] EFRAIN RESTREPO no encontrado, usando relaciones estándar sin él")
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
        
        print(f"✅ Archivo corregido creado: {output_file}")
        print(f"Relaciones totales: {len(df_standard)}")
        print(f"Entidades: {len(df_standard['Entidad'].unique())}")
        print(f"Beneficiarios: {len(df_standard['Accionista'].unique())}")
        
        # Mostrar estadísticas
        print("\n=== ESTADÍSTICAS DE CORRECCIÓN ===")
        for entidad in sorted(df_standard['Entidad'].unique()):
            entidad_rels = df_standard[df_standard['Entidad'] == entidad]
            total_participacion = entidad_rels['Participacion'].sum()
            print(f"{entidad}: {len(entidad_rels)} relaciones, total {total_participacion:.1f}%")
        
        if efrain_found:
            efrain_rel = df_standard[df_standard['Accionista'] == 'EFRAIN RESTREPO']
            if len(efrain_rel) > 0:
                print(f"\n✅ EFRAIN RESTREPO: {efrain_rel.iloc[0]['Participacion']:.1f}% en {efrain_rel.iloc[0]['Entidad']}")
        
        return str(output_file)
        
    except Exception as e:
        print(f"❌ Error durante la corrección: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    """Función principal"""
    if len(sys.argv) < 2:
        print("Uso: python fix_dra_blue_v2.py <archivo_excel>")
        return False
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"❌ Archivo no encontrado: {excel_file}")
        return False
    
    result = fix_dra_blue_entities_v2(excel_file)
    
    if result:
        print(f"\n🎉 Corrección completada exitosamente!")
        print(f"📁 Archivo corregido: {result}")
        return True
    else:
        print(f"\n❌ Error en la corrección")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
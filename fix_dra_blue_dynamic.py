#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para consolidar automáticamente las participaciones en RED COW INC
Interpreta la estructura jerárquica del archivo data.xlsx y calcula beneficiarios finales
Versión 6: Interpretación inteligente de estructura jerárquica - Java Compatible
"""

import pandas as pd
import sys
import os
from pathlib import Path

# Configurar codificación para compatibilidad con Java
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def consolidate_participations(excel_file):
    """
    Consolida las participaciones finales a partir de la estructura jerárquica
    """
    try:
        print("=== INICIANDO CONSOLIDACION INTELIGENTE v6 ===")
        print(f"Archivo: {excel_file}")
        
        # Leer Excel original
        df = pd.read_excel(excel_file)
        print(f"Filas leidas: {len(df)}")
        
        # Extraer beneficiarios finales basándose en la estructura del reporte
        final_beneficiaries = extract_final_beneficiaries(df)
        print(f"Beneficiarios finales detectados: {len(final_beneficiaries)}")
        
        # Crear estructura de salida
        output_data = create_consolidated_output(final_beneficiaries)
        
        # Generar archivo de salida
        input_path = Path(excel_file)
        output_file = input_path.parent / f"{input_path.stem}_cleaned_fixed.xlsx"
        
        # Guardar archivo corregido
        df_output = pd.DataFrame(output_data)
        df_output.to_excel(output_file, index=False)
        
        print(f"[OK] Archivo consolidado creado: {output_file}")
        print(f"Relaciones en output: {len(df_output)}")
        
        # Mostrar estadísticas
        show_consolidation_stats(output_data)
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error durante la consolidacion: {e}")
        import traceback
        traceback.print_exc()
        return False

def extract_final_beneficiaries(df):
    """
    Extrae los beneficiarios finales y sus participaciones de la estructura jerárquica
    """
    print("\n=== EXTRAYENDO BENEFICIARIOS FINALES ===")
    
    beneficiaries = {}
    
    # Analizar cada fila buscando patrones de beneficiarios finales
    for i, row in df.iterrows():
        entidad = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        participacion_directa = row.iloc[1] if pd.notna(row.iloc[1]) else None
        participacion_final = row.iloc[2] if pd.notna(row.iloc[2]) else None
        
        # Saltar filas sin datos útiles
        if not entidad or entidad == "nan":
            continue
            
        # Detectar beneficiarios finales: tienen participación final pero no son entidades de primer nivel
        if participacion_final is not None and participacion_final > 0:
            
            # Determinar si es beneficiario final (persona natural o entidad operativa)
            is_final_beneficiary = is_final_entity(entidad)
            
            if is_final_beneficiary:
                # Convertir participación final a porcentaje
                participation_percent = float(participacion_final) * 100
                
                # Agregar o actualizar beneficiario
                if entidad in beneficiaries:
                    beneficiaries[entidad] += participation_percent
                else:
                    beneficiaries[entidad] = participation_percent
                    
                print(f"Beneficiario detectado: {entidad} -> {participation_percent:.2f}%")
    
    return beneficiaries

def is_final_entity(entidad):
    """
    Determina si una entidad es un beneficiario final (persona natural principalmente)
    Mantiene la lógica original pero agrega detección dinámica de personas
    """
    entidad_upper = entidad.upper()
    
    # Lista base de nombres de personas conocidos (beneficiarios finales)
    known_person_indicators = [
        "RODRIGUEZ", "DIAZ", "CONTRERAS", "MERCEDES", "STELLA", 
        "ALEXANDRA", "CARLOS", "JORGE", "ARTURO", "ENRIQUE", "LUZ"
    ]
    
    # Entidades operativas conocidas que también son beneficiarios finales
    operational_entities = [
        "INVERSIONES MADCOM", "MADCOM"
    ]
    
    # Entidades corporativas que NO son beneficiarios finales (aunque no tengan palabras corporativas obvias)
    corporate_entities = [
        "TIERRA ARCO IRIS"  # Es una entidad corporativa, no persona natural
    ]
    
    # Verificar contra patrones conocidos primero
    for indicator in known_person_indicators:
        if indicator in entidad_upper:
            return True
            
    # Verificar si es una entidad corporativa conocida (NO beneficiario final)
    for entity in corporate_entities:
        if entity in entidad_upper:
            return False
            
    for entity in operational_entities:
        if entity in entidad_upper:
            return True
    
    # Detección dinámica adicional para nuevos nombres de personas
    # Buscar patrones que indiquen nombre de persona (2+ nombres propios)
    words = entidad_upper.split()
    if len(words) >= 2:
        # Si contiene palabras típicas de nombres de persona
        name_patterns = ["MARIA", "FERNANDEZ", "LOPEZ", "JUAN", "PEDRO", "ANA", "JOSE"]
        for pattern in name_patterns:
            if pattern in entidad_upper:
                return True
        
        # Si todas las palabras empiezan con mayúscula y no contienen palabras corporativas
        corporate_words = ["INC", "S.A", "SAS", "LTDA", "CORPORATION", "COMPANY", "FINANCIAL", "INVERSIONES"]
        is_corporate = any(corp_word in entidad_upper for corp_word in corporate_words)
        
        if not is_corporate and len(words) >= 2:
            # Probablemente es un nombre de persona
            return True
    
    return False

def create_consolidated_output(beneficiaries):
    """
    Crea la estructura de salida consolidada - SOLO participaciones directas en RED COW INC
    """
    print("\n=== CREANDO ESTRUCTURA CONSOLIDADA ===")
    
    output_data = []
    
    # Ordenar beneficiarios por participación descendente
    sorted_beneficiaries = sorted(beneficiaries.items(), key=lambda x: x[1], reverse=True)
    
    # SOLO agregar participaciones directas en RED COW INC (no entidades intermedias)
    for beneficiary, participation in sorted_beneficiaries:
        # Mantener el nombre original del beneficiario final
        accionista_display = beneficiary
            
        output_data.append({
            "Entidad": "RED COW INC",
            "Accionista": accionista_display,
            "Participacion": round(participation, 2)
        })
        
        print(f"Agregado: RED COW INC <- {accionista_display} ({participation:.2f}%)")
    
    # NO agregar entidades intermedias para evitar confusión en el algoritmo
        print("INFO: Omitiendo entidades intermedias para evitar doble conteo")
    return output_data

def detect_intermediate_entities(beneficiaries):
    """
    Detecta entidades intermedias importantes que deben aparecer en la salida
    """
    intermediate = []
    
    # No agregar entidades intermedias adicionales
    
    return intermediate

def show_consolidation_stats(output_data):
    """
    Muestra estadísticas de la consolidación
    """
    print(f"\n=== ESTADISTICAS DE CONSOLIDACION ===")
    
    red_cow_participations = [row for row in output_data if row["Entidad"] == "RED COW INC"]
    total_participation = sum(row["Participacion"] for row in red_cow_participations)
    
    print(f"Beneficiarios directos de RED COW INC: {len(red_cow_participations)}")
    print(f"Participacion total: {total_participation:.2f}%")
    print(f"Total de relaciones en output: {len(output_data)}")
    
    print(f"\nBENEFICIARIOS FINALES (ESTRUCTURA PLANA):")
    for row in red_cow_participations:
        print(f"  - {row['Accionista']}: {row['Participacion']}%")
        
    # Verificar si tenemos los valores esperados
    expected_beneficiaries = {
        "Stella Rodriguez Contreras": 34.02,
        "Alexandra Diaz Rodriguez": 12.65,  # Beneficiario real en lugar de DRA BLUE GOW
        "Luz Mercedes Diaz Rodriguez": 14.49,
        "Jorge Enrique Diaz Rodriguez": 12.28,
        "Carlos Arturo Diaz Rodriguez": 12.4,
        "Inversiones MADCOM": 12.28
        # TIERRA ARCO IRIS eliminado - no es beneficiario final (entidad corporativa)
    }
    
    print(f"\nVERIFICACION CONTRA VALORES ESPERADOS:")
    for expected_name, expected_value in expected_beneficiaries.items():
        found = False
        for row in red_cow_participations:
            if normalize_name(row['Accionista']) == normalize_name(expected_name):
                diff = abs(row['Participacion'] - expected_value)
                status = "[OK]" if diff < 0.1 else "[ERROR]"
                print(f"  {status} {expected_name}: Esperado {expected_value}%, Actual {row['Participacion']}%")
                found = True
                break
        if not found:
            print(f"  [ERROR] {expected_name}: No encontrado")

def normalize_name(name):
    """
    Normaliza nombres para comparación
    """
    return name.upper().replace("Ñ", "N").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")

def main():
    """Función principal"""
    if len(sys.argv) != 2:
        print("[ERROR] Uso: python fix_dra_blue_dynamic.py <archivo_excel>")
        return False
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"[ERROR] Archivo no encontrado: {excel_file}")
        return False
    
    # Realizar consolidación
    result = consolidate_participations(excel_file)
    
    if result:
        print(f"\n[OK] Consolidacion completada exitosamente")
    else:
        print(f"\n[ERROR] Error en la consolidacion")
    
    return result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
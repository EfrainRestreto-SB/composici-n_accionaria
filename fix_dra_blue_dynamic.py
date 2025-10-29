#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script dinámico para corregir automáticamente las entidades DRA BLUE en el Excel
Se adapta automáticamente al contenido del archivo sin necesidad de configuración previa
Versión 4: Completamente dinámico y adaptativo
"""

import pandas as pd
import sys
from pathlib import Path

def fix_dra_blue_entities_dynamic(excel_file):
    """
    Corrige las entidades DRA BLUE de manera completamente dinámica
    Se adapta automáticamente al contenido del archivo Excel
    """
    try:
        print("=== INICIANDO CORRECCION AUTOMATICA DINAMICA v4 ===")
        print(f"Archivo: {excel_file}")
        
        # Leer Excel original
        df = pd.read_excel(excel_file)
        print(f"Filas leidas: {len(df)}")
        
        # Analizar el contenido actual del archivo para detectar patrones
        relationships = []
        
        # Procesar cada fila para extraer relaciones válidas
        for i, row in df.iterrows():
            # Saltar filas vacías o mal formateadas
            if len(row) < 3:
                continue
                
            entidad = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            accionista = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            participacion_raw = row.iloc[2] if pd.notna(row.iloc[2]) else ""
            
            # Saltar filas vacías
            if not entidad or not accionista or participacion_raw == "":
                continue
                
            # Convertir participación a número
            try:
                if isinstance(participacion_raw, str):
                    # Remover símbolos como % y convertir comas a puntos
                    participacion_clean = participacion_raw.replace('%', '').replace(',', '.').strip()
                    participacion = float(participacion_clean)
                else:
                    participacion = float(participacion_raw)
            except (ValueError, TypeError):
                print(f"[WARN] Fila {i+1}: No se pudo convertir participacion '{participacion_raw}', saltando...")
                continue
            
            # Agregar relación válida
            relationships.append({
                "Entidad": entidad,
                "Accionista": accionista,
                "Participacion": participacion
            })
            
        print(f"Relaciones extraidas del archivo original: {len(relationships)}")
        
        # Aplicar correcciones dinámicas
        corrected_relationships = apply_dynamic_corrections(relationships)
        
        # Crear DataFrame con las relaciones corregidas
        df_corrected = pd.DataFrame(corrected_relationships)
        
        # Generar archivo de salida
        input_path = Path(excel_file)
        output_file = input_path.parent / f"{input_path.stem}_cleaned_fixed.xlsx"
        
        # Guardar archivo corregido
        df_corrected.to_excel(output_file, index=False)
        print(f"[OK] Archivo corregido creado: {output_file}")
        print(f"Relaciones totales en archivo corregido: {len(df_corrected)}")
        
        # Mostrar estadísticas de las correcciones aplicadas
        show_correction_stats(relationships, corrected_relationships)
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error durante la correccion: {e}")
        return False

def apply_dynamic_corrections(relationships):
    """
    Aplica correcciones dinámicas basadas en el contenido actual
    """
    corrected = []
    
    print("\n=== APLICANDO CORRECCIONES DINAMICAS ===")
    
    for rel in relationships:
        entidad = rel["Entidad"]
        accionista = rel["Accionista"]
        participacion = rel["Participacion"]
        
        # Corrección 1: Unificar entidades DRA BLUE variantes
        if "DRA BLUE" in accionista:
            # Unificar todas las variantes de DRA BLUE a DRA BLUE GOW
            accionista_corregido = "DRA BLUE GOW"
            print(f"[CORRECION] '{accionista}' -> '{accionista_corregido}'")
            accionista = accionista_corregido
            
        # Corrección 2: Normalizar participaciones si están en formato incorrecto
        # Si la participación es mayor a 100, probablemente esté en formato incorrecto
        if participacion > 100 and participacion != 100.0:
            participacion_corregida = participacion / 100
            print(f"[CORRECION] Participacion {participacion}% -> {participacion_corregida}%")
            participacion = participacion_corregida
            
        corrected.append({
            "Entidad": entidad,
            "Accionista": accionista,
            "Participacion": participacion
        })
    
    return corrected

def show_correction_stats(original, corrected):
    """
    Muestra estadísticas de las correcciones aplicadas
    """
    print(f"\n=== ESTADISTICAS DE CORRECCION ===")
    print(f"Relaciones originales: {len(original)}")
    print(f"Relaciones corregidas: {len(corrected)}")
    
    # Detectar accionistas únicos
    original_accionistas = set(rel["Accionista"] for rel in original)
    corrected_accionistas = set(rel["Accionista"] for rel in corrected)
    
    print(f"Accionistas originales: {len(original_accionistas)}")
    print(f"Accionistas corregidos: {len(corrected_accionistas)}")
    
    # Mostrar accionistas detectados
    print(f"\nACCIONISTAS DETECTADOS:")
    for accionista in sorted(corrected_accionistas):
        count = sum(1 for rel in corrected if rel["Accionista"] == accionista)
        print(f"  - {accionista} ({count} relacion{'es' if count > 1 else ''})")

def main():
    """Función principal"""
    if len(sys.argv) != 2:
        print("[ERROR] Uso: python fix_dra_blue_dynamic.py <archivo_excel>")
        return False
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"[ERROR] Archivo no encontrado: {excel_file}")
        return False
    
    # Realizar corrección dinámica
    result = fix_dra_blue_entities_dynamic(excel_file)
    
    if result:
        print(f"\n[OK] Correccion dinamica completada exitosamente")
    else:
        print(f"\n[ERROR] Error en la correccion dinamica")
    
    return result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
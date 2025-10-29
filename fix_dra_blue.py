import pandas as pd
import sys

def convert_hierarchical_to_standard(df):
    """
    Convierte el formato jerárquico a formato estándar de 3 columnas
    """
    relationships = []
    current_entity = None
    last_entity_was_parent = False
    
    for i, row in df.iterrows():
        entity_name = row.iloc[0]  # Primera columna
        participation = row.iloc[1]  # Segunda columna
        
        # Skip filas vacías o con valores NaN
        if pd.isna(entity_name) or entity_name == "X":
            continue
            
        entity_name = str(entity_name).strip()
        
        # Debug para DRA BLUE
        if 'DRA BLUE' in entity_name or 'ALEXANDRA' in entity_name:
            print(f"DEBUG: Fila {i}, entity='{entity_name}', participation={participation}, current_entity={current_entity}, last_was_parent={last_entity_was_parent}")
        
        # Si la participación es NaN y hay un nombre, es una nueva entidad padre
        if pd.isna(participation) and entity_name:
            current_entity = entity_name
            last_entity_was_parent = True
            if 'DRA BLUE' in entity_name:
                print(f"DEBUG: Estableciendo current_entity = {current_entity} (participación NaN)")
        # Si acabamos de ver una entidad padre y vemos participación válida, es hijo de esa entidad
        elif last_entity_was_parent and pd.notna(participation) and current_entity:
            # Convertir de decimal (0-1) a porcentaje (0-100) si es necesario
            participation_value = float(participation)
            if participation_value <= 1.0:
                participation_value *= 100
            
            relationships.append({
                'Entidad': current_entity,
                'Accionista': entity_name,
                'Participación': participation_value
            })
            last_entity_was_parent = False
            if current_entity and 'DRA BLUE' in current_entity:
                print(f"DEBUG: Agregando relación {current_entity} -> {entity_name} ({participation_value}%)")
        # Si la participación es 1.0 y no es después de una entidad padre, es una nueva entidad raíz
        elif pd.notna(participation) and float(participation) == 1.0 and not last_entity_was_parent:
            current_entity = entity_name
            last_entity_was_parent = False
            if 'DRA BLUE' in entity_name:
                print(f"DEBUG: Estableciendo current_entity = {current_entity} (raíz con 1.0)")
        # Si hay participación válida y menor a 1.0, es una relación hijo-padre
        elif pd.notna(participation) and current_entity and float(participation) < 1.0:
            # Convertir de decimal (0-1) a porcentaje (0-100)
            participation_percent = float(participation) * 100
            relationships.append({
                'Entidad': current_entity,
                'Accionista': entity_name,
                'Participación': participation_percent
            })
            last_entity_was_parent = False
    
    return pd.DataFrame(relationships)

def fix_dra_blue_entities(excel_path):
    """
    Corrige las entidades DRA BLUE GLOW INC y DRA BLUE GOW INC unificándolas
    """
    print(f"Leyendo Excel: {excel_path}")
    df_raw = pd.read_excel(excel_path)
    
    print("=== Convirtiendo formato jerárquico a estándar ===")
    df = convert_hierarchical_to_standard(df_raw)
    
    print("=== Datos convertidos ===")
    print(df.to_string())
    
    # Buscar las relaciones problemáticas
    dra_blue_glow_rows = df[df['Accionista'] == 'DRA BLUE GLOW INC']
    dra_blue_gow_rows = df[df['Entidad'] == 'DRA BLUE GOW INC']
    
    print(f"\n=== DRA BLUE GLOW INC como accionista ===")
    print(dra_blue_glow_rows.to_string())
    
    print(f"\n=== DRA BLUE GOW INC como entidad ===") 
    print(dra_blue_gow_rows.to_string())
    
    if not dra_blue_glow_rows.empty and not dra_blue_gow_rows.empty:
        print("\n=== Aplicando corrección ===")
        
        # Para cada hijo de DRA BLUE GOW INC, crear una nueva relación con DRA BLUE GLOW INC
        new_rows = []
        for _, gow_row in dra_blue_gow_rows.iterrows():
            # Convertir participación si es necesario
            participation_value = gow_row['Participación']
            if participation_value <= 1.0:
                participation_value *= 100
                
            new_row = {
                'Entidad': 'DRA BLUE GLOW INC',
                'Accionista': gow_row['Accionista'], 
                'Participación': participation_value
            }
            new_rows.append(new_row)
            print(f"Nueva relación: {new_row}")
        
        # Remover las filas de DRA BLUE GOW INC
        df_fixed = df[df['Entidad'] != 'DRA BLUE GOW INC'].copy()
        
        # Agregar las nuevas relaciones
        for new_row in new_rows:
            df_fixed = pd.concat([df_fixed, pd.DataFrame([new_row])], ignore_index=True)
        
        print(f"\n=== Datos corregidos ===")
        print(df_fixed.to_string())
        
        # Guardar el Excel corregido
        output_path = excel_path.replace('.xlsx', '_cleaned_fixed.xlsx')
        df_fixed.to_excel(output_path, index=False)
        print(f"\nExcel corregido guardado: {output_path}")
        
        return output_path
    else:
        print("No se encontraron las entidades DRA BLUE para corregir")
        return excel_path

if __name__ == "__main__":
    excel_file = r"C:\Users\user\Documents\data.xlsx"
    fix_dra_blue_entities(excel_file)
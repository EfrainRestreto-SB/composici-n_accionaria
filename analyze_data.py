import pandas as pd

def analyze_data_xlsx():
    try:
        # Leer el archivo Excel data.xlsx
        df = pd.read_excel('data.xlsx', header=None)
        print('Análisis de data.xlsx - Filas 4 a 45:')
        print('='*80)
        
        # Mostrar las filas 4-45 (índices 3-44 en Python)
        for i in range(3, min(45, len(df))):
            row = df.iloc[i]
            # Mostrar las primeras 3 columnas que contienen datos
            col_a = str(row[0]) if pd.notna(row[0]) else ''
            col_b = str(row[1]) if pd.notna(row[1]) else ''
            col_c = str(row[2]) if pd.notna(row[2]) else ''
            
            if col_a or col_b or col_c:  # Solo mostrar filas que tienen datos
                print(f'Fila {i+1:2d}: A="{col_a}" | B="{col_b}" | C="{col_c}"')
        
        print('='*80)
        print('Total de filas en el archivo:', len(df))
        
        # Analizar la estructura de columnas
        print('\nEstructura de columnas:')
        for col_idx in range(min(5, len(df.columns))):
            non_empty = df[col_idx].dropna()
            if len(non_empty) > 0:
                print(f'Columna {col_idx} ({chr(65+col_idx)}): {len(non_empty)} valores no vacíos')
                # Mostrar algunos ejemplos
                examples = non_empty.head(3).tolist()
                print(f'  Ejemplos: {examples}')
        
    except Exception as e:
        print(f'Error al leer data.xlsx: {e}')

if __name__ == "__main__":
    analyze_data_xlsx()
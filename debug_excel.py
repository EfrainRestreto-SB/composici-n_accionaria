import pandas as pd

df = pd.read_excel(r"C:\Users\user\Documents\data.xlsx")
print('=== Filas 35-38 ===')
for i in range(35, 38):
    if i < len(df):
        row = df.iloc[i]
        entity = row.iloc[0]
        participation = row.iloc[1]
        print(f'Fila {i}: "{entity}" -> {participation} (isna={pd.isna(participation)})')

print('\n=== Buscando ALEXANDRA después de DRA BLUE GOW ===')
found_dra_gow = False
for i, row in df.iterrows():
    entity = row.iloc[0]
    if pd.notna(entity) and 'DRA BLUE GOW' in str(entity):
        found_dra_gow = True
        print(f'Encontrado DRA BLUE GOW en fila {i}')
    elif found_dra_gow and pd.notna(entity) and 'ALEXANDRA' in str(entity):
        participation = row.iloc[1]
        print(f'ALEXANDRA encontrada en fila {i}: "{entity}" -> {participation}')
        break
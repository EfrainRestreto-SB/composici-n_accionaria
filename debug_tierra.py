import pandas as pd
df = pd.read_excel('C:/Users/user/Documents/data.xlsx')
print('=== TODAS LAS ENTIDADES CON TIERRA ARCO IRIS ===')
tierra_rows = df[df.apply(lambda row: any('TIERRA ARCO IRIS' in str(cell).upper() if pd.notna(cell) else False for cell in row), axis=1)]
for idx, row in tierra_rows.iterrows():
    print(f'Fila {idx}: {row["Unnamed: 0"]} | {row["DESGLOSE DE COMPOSICION ACCIONARIA "]} | {row["Unnamed: 2"]}')

print()
print('=== TODAS LAS ENTIDADES CON PARTICIPACIONES FINALES NO NULAS ===')
final_participation_rows = df[pd.notna(df["Unnamed: 2"]) & (df["Unnamed: 2"] != 0)]
for idx, row in final_participation_rows.iterrows():
    print(f'Fila {idx}: {row["Unnamed: 0"]} | {row["DESGLOSE DE COMPOSICION ACCIONARIA "]} | {row["Unnamed: 2"]}')
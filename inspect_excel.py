import pandas as pd

# Leer el Excel
excel_file = 'Seguimiento desercion FCE 2023 - 2025 (1).xlsx'

# Obtener nombres de hojas
xls = pd.ExcelFile(excel_file)
print("Hojas en el Excel:", xls.sheet_names)

# Leer cada hoja y mostrar las primeras filas y columnas
for sheet in xls.sheet_names:
    df = pd.read_excel(excel_file, sheet_name=sheet)
    print(f"\nHoja: {sheet}")
    print("Columnas:", list(df.columns))
    print("Primeras 5 filas:")
    print(df.head())
    print("Número de filas:", len(df))
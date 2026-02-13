import pandas as pd

path = 'precios_market.xlsx'
try:
    df = pd.read_excel(path)
    print(df.head(10).to_string(index=False))
except Exception as e:
    print('ERROR leyendo', path, e)

import pandas as pd

df = pd.DataFrame({
    'nome': ['João', 'Maria'],
    'idade': [25, 30]
})

df.to_csv('arquivo.csv', index=False)

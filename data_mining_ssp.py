# Importação das bibliotecas padrões utilizadas
import pandas as pd
import numpy as np
import datetime
import chardet

# Detecta a codificação do arquivo
with open('sheets/SspCompleta.csv', 'rb') as f:
    rawdata = f.read(10000)  # Lê amostra inicial
    result = chardet.detect(rawdata)

# Efetua a leitura do arquivo, armazenando em um dataframe
df = pd.read_csv('sheets/SspCompleta.csv', encoding=result['encoding'], sep=';')

# Exibir as primeiras linhas do dataframe obtido
print(df.head())
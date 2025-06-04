import pandas as pd
import numpy as np
import datetime
import chardet

# Detecta a codificação do arquivo
with open('sheets/SafCompleta.csv', 'rb') as f:
    rawdata = f.read(10000)  # Lê amostra inicial
    result = chardet.detect(rawdata)

# Efetua a leitura do arquivo, armazenando em um dataframe
df = pd.read_csv('sheets/SafCompleta.csv', encoding=result['encoding'], sep=';')

# Filtrando as SAFs por linha
df_linha_tue = df[df['Grupo'] == 'MATERIAL RODANTE - TUE']

filtro_vlt = df['Grupo'] == 'MATERIAL RODANTE - VLT'

filtro_linha_oeste = df['Veículo'].isin(['VLT01', 'VLT02', 'VLT03', 'VLT04', 'VLT05', 'VLT06'])
df_vlt_oeste = df[filtro_vlt & filtro_linha_oeste]

filtro_linha_nordeste = df['Veículo'].isin(['VLT07', 'VLT08', 'VLT09', 'VLT10', 'VLT11', 'VLT12', 'VLT13'])
df_vlt_nordeste = df[filtro_vlt & filtro_linha_nordeste]

filtro_linha_sobral = df['Veículo'].isin(['VLTS02', 'VLTS03', 'VLTS04', 'VLTS05', 'VLTS06'])
df_vlt_sobral = df[filtro_vlt & filtro_linha_sobral]

filtro_linha_cariri = df['Veículo'].isin(['TRAM1', 'TRAM2', 'VLTC03'])
df_vlt_cariri = df[filtro_vlt & filtro_linha_cariri]

countA = np.zeros(32, dtype=int)
countB = np.zeros(32, dtype=int)
countC = np.zeros(32, dtype=int)

# Filtrando e fazendo a contagem de SAFs por dia [TUE]
df_tue = df_linha_tue.copy()
if not df_linha_tue.empty:
    df_tue["Data de Abertura"] = pd.to_datetime(df_tue["Data de Abertura"], dayfirst=True, format="mixed")
    df_tue["Apenas_Data"] = df_tue["Data de Abertura"].dt.strftime('%d/%m/%Y')

    # # Verificar o formato real das datas na coluna
    primeira_data = df_tue['Apenas_Data'].iloc[0]
    print(f"Formato da data de exemplo: {primeira_data}")

    # Extrair mês e ano de uma data válida
    try:
        partes = primeira_data.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        ano = int(partes[2])
    except (IndexError, ValueError):
        mes = datetime.now().month  # Valor padrão se falhar
        ano = datetime.now().year   # Valor padrão se falhar

    # Iterar pelos dias do mês (de 1 a 31)
    for dia in range(1, 32):
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        filtro_data = df_tue['Apenas_Data'] == data_str
        df_filtrado = df_tue[filtro_data]
    # Contar os níveis
        countA[dia] = len(df_filtrado[df_filtrado['Nível'] == 'A'])
        countB[dia] = len(df_filtrado[df_filtrado['Nível'] == 'B'])
        countC[dia] = len(df_filtrado[df_filtrado['Nível'] == 'C'])
        
    # Verifica se o DataFrame filtrado não está vazio
        if not df_filtrado.empty: 
            print(f" Dados para {data_str}: ".center(70, '='))
            print(df_filtrado[['Nível', 'Grupo', 'Veículo', 'Apenas_Data', 'Status']])
            print(f"Dia {dia}: A = {countA[dia]}, B = {countB[dia]}, C = {countC[dia]}.")
            print(70*'=', '\n')



# Filtrando e fazendo a contagem de SAFs por dia [VLT OESTE]
df_vlt = df_vlt_oeste.copy()
if not df_vlt.empty:
    df_vlt["Data de Abertura"] = pd.to_datetime(df_vlt["Data de Abertura"], dayfirst=True, format="mixed")
    df_vlt["Apenas_Data"] = df_vlt["Data de Abertura"].dt.strftime('%d/%m/%Y')

    # Verificar o formato real das datas na coluna
    primeira_data = df_vlt['Apenas_Data'].iloc[0]

    # Extrair mês e ano de uma data válida
    try:
        partes = primeira_data.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        ano = int(partes[2])
    except (IndexError, ValueError):
        mes = datetime.now().month  # Valor padrão se falhar
        ano = datetime.now().year   # Valor padrão se falhar

    # Iterar pelos dias do mês (de 1 a 31)
    for dia in range(1, 32):
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        filtro_data = df_vlt['Apenas_Data'] == data_str
        df_filtrado = df_vlt[filtro_data]
    # Contar os níveis
        countA[dia] = len(df_filtrado[df_filtrado['Nível'] == 'A'])
        countB[dia] = len(df_filtrado[df_filtrado['Nível'] == 'B'])
        countC[dia] = len(df_filtrado[df_filtrado['Nível'] == 'C'])
        
    # Verifica se o DataFrame filtrado não está vazio
        if not df_filtrado.empty: 
            print(f" Dados para {data_str}: ".center(70, '='))
            print(df_filtrado[['Nível', 'Grupo', 'Veículo', 'Apenas_Data', 'Status']])
            print(f"Dia {dia}: A = {countA[dia]}, B = {countB[dia]}, C = {countC[dia]}.")
            print(70*'=', '\n')


# Filtrando e fazendo a contagem de SAFs por dia [VLT NORDESTE]
df_vlt = df_vlt_nordeste.copy()
if not df_vlt.empty:
    df_vlt["Data de Abertura"] = pd.to_datetime(df_vlt["Data de Abertura"], dayfirst=True, format="mixed")
    df_vlt["Apenas_Data"] = df_vlt["Data de Abertura"].dt.strftime('%d/%m/%Y')

    # Verificar o formato real das datas na coluna
    primeira_data = df_vlt['Apenas_Data'].iloc[0]

    # Extrair mês e ano de uma data válida
    try:
        partes = primeira_data.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        ano = int(partes[2])
    except (IndexError, ValueError):
        mes = datetime.now().month  # Valor padrão se falhar
        ano = datetime.now().year   # Valor padrão se falhar

    # Iterar pelos dias do mês (de 1 a 31)
    for dia in range(1, 32):
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        filtro_data = df_vlt['Apenas_Data'] == data_str
        df_filtrado = df_vlt[filtro_data]
    # Contar os níveis
        countA[dia] = len(df_filtrado[df_filtrado['Nível'] == 'A'])
        countB[dia] = len(df_filtrado[df_filtrado['Nível'] == 'B'])
        countC[dia] = len(df_filtrado[df_filtrado['Nível'] == 'C'])
        
    # Verifica se o DataFrame filtrado não está vazio
        if not df_filtrado.empty: 
            print(f" Dados para {data_str}: ".center(70, '='))
            print(df_filtrado[['Nível', 'Grupo', 'Veículo', 'Apenas_Data', 'Status']])
            print(f"Dia {dia}: A = {countA[dia]}, B = {countB[dia]}, C = {countC[dia]}.")
            print(70*'=', '\n')


# Filtrando e fazendo a contagem de SAFs por dia [VLT SOBRAL]
df_vlt = df_vlt_sobral.copy()
if not df_vlt.empty:
    df_vlt["Data de Abertura"] = pd.to_datetime(df_vlt["Data de Abertura"], dayfirst=True, format="mixed")
    df_vlt["Apenas_Data"] = df_vlt["Data de Abertura"].dt.strftime('%d/%m/%Y')

    # Verificar o formato real das datas na coluna
    primeira_data = df_vlt['Apenas_Data'].iloc[0]

    # Extrair mês e ano de uma data válida
    try:
        partes = primeira_data.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        ano = int(partes[2])
    except (IndexError, ValueError):
        mes = datetime.now().month  # Valor padrão se falhar
        ano = datetime.now().year   # Valor padrão se falhar

    # Iterar pelos dias do mês (de 1 a 31)
    for dia in range(1, 32):
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        filtro_data = df_vlt['Apenas_Data'] == data_str
        df_filtrado = df_vlt[filtro_data]
    # Contar os níveis
        countA[dia] = len(df_filtrado[df_filtrado['Nível'] == 'A'])
        countB[dia] = len(df_filtrado[df_filtrado['Nível'] == 'B'])
        countC[dia] = len(df_filtrado[df_filtrado['Nível'] == 'C'])
        
    # Verifica se o DataFrame filtrado não está vazio
        if not df_filtrado.empty: 
            print(f" Dados para {data_str}: ".center(70, '='))
            print(df_filtrado[['Nível', 'Grupo', 'Veículo', 'Apenas_Data', 'Status']])
            print(f"Dia {dia}: A = {countA[dia]}, B = {countB[dia]}, C = {countC[dia]}.")
            print(70*'=', '\n')



# Filtrando e fazendo a contagem de SAFs por dia [VLT CARIRI]
df_vlt = df_vlt_cariri.copy()
if not df_vlt.empty:
    df_vlt["Data de Abertura"] = pd.to_datetime(df_vlt["Data de Abertura"], dayfirst=True, format="mixed")
    df_vlt["Apenas_Data"] = df_vlt["Data de Abertura"].dt.strftime('%d/%m/%Y')

    # Verificar o formato real das datas na coluna
    primeira_data = df_vlt['Apenas_Data'].iloc[0]

    # Extrair mês e ano de uma data válida
    try:
        partes = primeira_data.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        ano = int(partes[2])
    except (IndexError, ValueError):
        mes = datetime.now().month  # Valor padrão se falhar
        ano = datetime.now().year   # Valor padrão se falhar

    # Iterar pelos dias do mês (de 1 a 31)
    for dia in range(1, 32):
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        filtro_data = df_vlt['Apenas_Data'] == data_str
        df_filtrado = df_vlt[filtro_data]
    # Contar os níveis
        countA[dia] = len(df_filtrado[df_filtrado['Nível'] == 'A'])
        countB[dia] = len(df_filtrado[df_filtrado['Nível'] == 'B'])
        countC[dia] = len(df_filtrado[df_filtrado['Nível'] == 'C'])
        
    # Verifica se o DataFrame filtrado não está vazio
        if not df_filtrado.empty: 
            print(f" Dados para {data_str}: ".center(70, '='))
            print(df_filtrado[['Nível', 'Grupo', 'Veículo', 'Apenas_Data', 'Status']])
            print(f"Dia {dia}: A = {countA[dia]}, B = {countB[dia]}, C = {countC[dia]}.")
            print(70*'=', '\n')
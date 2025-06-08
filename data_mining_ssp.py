# Importação das bibliotecas padrões utilizadas
import pandas as pd
import chardet

# Detecta a codificação do arquivo
with open('sheets/SspCompleta.csv', 'rb') as f:
    rawdata = f.read(10000)  # Lê amostra inicial
    result = chardet.detect(rawdata)

# Efetua a leitura do arquivo, armazenando em um dataframe
df = pd.read_csv('sheets/SspCompleta.csv', encoding=result['encoding'], sep=';')

def linha_vlt(dataframe, linha): # Função para filtrar e retornar a linha de VLT desejada
    filtro_material_rodante = dataframe['Grupo de Sistemas'] == 'MATERIAL RODANTE - VLT'
    if not filtro_material_rodante.empty:    
        match linha:
            case 'oeste':
                linha_oeste = dataframe[filtro_material_rodante & dataframe['Veiculo'].isin(['VLT01', 'VLT02', 'VLT03', 'VLT04', 'VLT05', 'VLT06'])]
                return linha_oeste
            case 'nordeste':
                linha_nordeste = dataframe[filtro_material_rodante & dataframe['Veiculo'].isin(['VLT07', 'VLT08', 'VLT09', 'VLT10', 'VLT11', 'VLT12', 'VLT13'])]
                return linha_nordeste
            case 'sobral':
                linha_sobral = dataframe[filtro_material_rodante & dataframe['Veiculo'].isin(['VLTS02', 'VLTS03', 'VLTS04', 'VLTS05', 'VLTS06'])]
                return linha_sobral
            case 'cariri':
                linha_cariri = dataframe[filtro_material_rodante & dataframe['Veiculo'].isin(['TRAM1', 'TRAM2', 'VLTC03'])]
                return linha_cariri
    else:
        print("Dataframe sem dados...")

def linha_servico(dataframe, linha): # Função para filtrar e retornar somente os serviços necessários para a leitura em cada linha
    match linha:
        case 'oeste':
            linha_servico = dataframe[dataframe['Serviço'].isin(['MANUTENÇÃO PREVENTIVA DIÁRIA', 'MANUTENÇÃO PREVENTIVA SEMANAL'])]
            return linha_servico
        case 'nordeste':
            linha_servico = dataframe[dataframe['Serviço'].isin(['MANUTENÇÃO PREVENTIVA DIÁRIA', 'MANUTENÇÃO PREVENTIVA SEMANAL'])]
            return linha_servico
        case 'sobral':
            linha_servico = dataframe[dataframe['Serviço'].isin(['MANUTENÇÃO PREVENTIVA DIÁRIA', 'MANUTENÇÃO PREVENTIVA SEMANAL'])]
            return linha_servico
        case 'cariri':
            linha_servico = dataframe[dataframe['Serviço'].isin(['MANUTENÇÃO PREVENTIVA DIÁRIA', 'MANUTENÇÃO PREVENTIVA SEMANAL'])]
            return linha_servico

linha_oeste = linha_servico(linha_vlt(df, 'oeste'), 'oeste')
linha_oeste["Data Abertura"] = pd.to_datetime(linha_oeste["Data Abertura"], dayfirst=True, format="mixed")
linha_oeste["Apenas_Data"] = linha_oeste["Data Abertura"].dt.strftime('%d/%m/%Y')
print(linha_oeste.head())

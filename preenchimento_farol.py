import os
import time
from pathlib import Path
from typing import Optional
import pandas as pd
import calendar
import datetime
import pyautogui

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# Caminhos (definidos pelo usuário)
DADOS_PATH = Path(r"C:\Users\gabriel.silva\Downloads\Documentos - Farol\dados_resumo.xlsx")
TARGET_EXCEL = r"C:\Users\gabriel.silva\Documents\py\Farol Diário_Acompanhamento de Atividades 01_10_2025.xlsx"

def read_data(path: Path) -> pd.DataFrame:
    """Lê o arquivo dados_resumo.xlsx e retorna um DataFrame."""
    return pd.read_excel(path)


def format_cell(cell_template: str, row_number: int) -> str:
    """Formata uma célula que pode conter o placeholder {row}.

    Ex: 'B{row}' com row_number=2 -> 'B2'
    """
    if '{row}' in cell_template:
        return cell_template.format(row=row_number)
    return cell_template

def open_excel(filepath: str, wait_seconds: float = 4.0):
    """Abre o Excel com o arquivo especificado (Windows)."""
    os.startfile(filepath)
    time.sleep(wait_seconds)  # tempo para o Excel abrir

def go_to_cell(cell: str):
    """Abre diálogo 'Ir para' e posiciona na célula informada."""
    pyautogui.hotkey('ctrl', 'g')  # abre Go To no Excel
    time.sleep(0.12)
    pyautogui.typewrite(cell)
    pyautogui.press('enter')
    time.sleep(0.08)


def go_to_sheet_cell(sheet_name: str, cell: str):
    """Navega para uma célula em uma aba específica usando o comando Go To com referência 'Sheet'!Cell."""
    # Excel aceita referências do tipo 'Sheet Name'!A1
    ref = f"'{sheet_name}'!{cell}"
    pyautogui.hotkey('ctrl', 'g')
    time.sleep(0.12)
    pyautogui.typewrite(ref)
    pyautogui.press('enter')
    time.sleep(0.12)

def write_value(value):
    """Escreve o valor na célula selecionada."""
    # Evita escrever None/nan de forma inesperada
    if pd.isna(value):
        value = ""
    pyautogui.typewrite(str(value))
    pyautogui.press('enter')
    time.sleep(0.15)


def column_letter_to_index(col: str) -> int:
    """Converte letra(s) de coluna Excel para índice 1-based (ex: 'A'->1, 'AA'->27)."""
    col = col.upper()
    result = 0
    for ch in col:
        if 'A' <= ch <= 'Z':
            result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def index_to_column_letter(idx: int) -> str:
    """Converte índice 1-based para letra(s) de coluna Excel (ex: 1->'A', 27->'AA')."""
    letters = ''
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(rem + ord('A')) + letters
    return letters


def fill_level_from_series(series, start_col_letter: str, excel_row: int, dry_run: bool = False, days_to_fill: Optional[int] = None):
    """Preenche uma sequência horizontal a partir de start_col_letter na linha excel_row usando valores de series.

    Ex: start_col_letter='E', excel_row=54, series com 31 itens.
    """
    start_idx = column_letter_to_index(start_col_letter)
    # determinar quantos dias preencher
    max_possible = 31
    if days_to_fill is None:
        days = min(max_possible, len(series))
    else:
        days = min(max_possible, int(days_to_fill))

    for i in range(days):
        col_letter = index_to_column_letter(start_idx + i)
        cell = f"{col_letter}{excel_row}"
        # obter valor do series (i -> dia i+1)
        try:
            value = series.iloc[i]
        except Exception:
            value = ""
        # tratar zeros -> '-'
        try:
            sval = str(value).strip()
        except Exception:
            sval = ''

        if sval in ("0", "0.0", "0,0"):
            out_value = '-'
        elif sval == 'nan' or sval == '' or sval.upper() == 'NAN':
            out_value = ''
        else:
            out_value = value

        if dry_run:
            print(f"[dry-run] {cell} <- {out_value}")
            continue

        go_to_cell(cell)
        write_value(out_value)

def fill_row_from_series(series: pd.Series, mapping: dict):
    """
    Preenche uma linha na planilha do Excel usando um mapeamento.
    mapping: dict onde chave = célula no Excel (ex: 'B2'), valor = nome da coluna no DataFrame (ex: 'nome')
    """
    for cell_template, col in mapping.items():
        # cell_template pode ser 'B{row}' ou 'B2'
        cell = format_cell(cell_template, series.name if isinstance(series.name, int) else series.name)
        value = series.get(col, "")
        go_to_cell(cell)
        write_value(value)

def fill_multiple_rows(df: pd.DataFrame, mapping: dict, row_indices=None):
    """
    Preenche várias linhas.
    row_indices: lista de índices do DataFrame a serem usados (se None usa apenas o índice 0).
    """
    if row_indices is None:
        row_indices = list(df.index)

    for idx in row_indices:
        if idx not in df.index:
            print(f"Índice {idx} não encontrado no DataFrame; pulando.")
            continue
        fill_row_from_series(df.loc[idx], mapping)



def fill_sequence(df: pd.DataFrame, mapping: dict, start_row: int = 2, count: Optional[int] = None, dry_run: bool = False):
    """Preenche uma sequência de linhas na planilha usando mapeamento com placeholders.

    mapping: dict onde chave é um template de célula (ex: 'B{row}') e valor é nome da coluna no df.
    start_row: linha inicial no Excel onde {row} começa (normalmente 2)
    count: quantas linhas preencher; se None usa len(df)
    dry_run: se True, apenas imprime as ações sem chamar pyautogui
    """
    if count is None:
        count = len(df)

    for i in range(count):
        excel_row = start_row + i
        series = df.iloc[i].copy()
        # setar um nome numérico para usar em format_cell
        series.name = excel_row
        # Mostrar o que será preenchido
        pairs = []
        for cell_template, col in mapping.items():
            cell = format_cell(cell_template, excel_row)
            value = series.get(col, "")
            pairs.append((cell, value))

        if dry_run:
            print(f"[dry-run] Linha Excel: {excel_row}")
            for cell, value in pairs:
                print(f"    {cell} <- {value}")
            continue

        # Realizar preenchimento com confirmação por linha (rápida)
        print(f"Preenchendo linha Excel {excel_row} ({i+1}/{count})...")
        for cell, value in pairs:
            go_to_cell(cell)
            write_value(value)
        time.sleep(0.2)

if __name__ == "__main__":
    # Leitura dos dados
    try:
        df = read_data(DADOS_PATH)
    except Exception as e:
        print("Erro ao ler dados_resumo.xlsx:", e)
        raise

    # Ler aba específica do arquivo dados_resumo.xlsx
    try:
        df_safts = pd.read_excel(DADOS_PATH, sheet_name="SAF's Diárias")
    except Exception as e:
        print(f"Erro ao ler aba 'SAF's Diárias' de {DADOS_PATH}: {e}")
        raise

    # As colunas B,C,D correspondem a índices 1,2,3 (0-based: 1,2,3)
    # Vamos pegar até 31 linhas (dias 1..31) da planilha de resumo; assumir que cada coluna tem os valores por dia
    # Ajuste conforme a estrutura real do seu arquivo
    # -------- TUE -----------
    try:
        col_B = df_safts.iloc[:, 1].fillna("")
        col_C = df_safts.iloc[:, 2].fillna("")
        col_D = df_safts.iloc[:, 3].fillna("")
    except Exception as e:
        print(f"Erro ao extrair colunas B,C,D: {e}")
        raise

    # -------- OESTE -----------
    try:
        col_F = df_safts.iloc[:, 5].fillna("")
        col_G = df_safts.iloc[:, 6].fillna("")
        col_H = df_safts.iloc[:, 7].fillna("")
    except Exception as e:
        print(f"Erro ao extrair colunas F,G,H: {e}")
        raise

    # -------- NORDESTE -----------
    try:
        col_J = df_safts.iloc[:, 9].fillna("")
        col_K = df_safts.iloc[:, 10].fillna("")
        col_L = df_safts.iloc[:, 11].fillna("")
    except Exception as e:
        print(f"Erro ao extrair colunas J,K,L: {e}")
        raise

    # -------- SOBRAL -----------
    try:
        col_N = df_safts.iloc[:, 13].fillna("")
        col_O = df_safts.iloc[:, 14].fillna("")
        col_P = df_safts.iloc[:, 15].fillna("")
    except Exception as e:
        print(f"Erro ao extrair colunas N,O,P: {e}")
        raise

    # -------- CARIRI -----------
    try:
        col_R = df_safts.iloc[:, 17].fillna("")
        col_S = df_safts.iloc[:, 18].fillna("")
        col_T = df_safts.iloc[:, 19].fillna("")
    except Exception as e:
        print(f"Erro ao extrair colunas R,S,T: {e}")
        raise


    # Preparar séries para níveis A (B), B (C), C (D)
    # Detectar e pular linhas de cabeçalho (por exemplo, que contenham 'TUE', 'A', 'B', 'C')
    header_labels = {'TUE', 'A', 'B', 'C', 'NÍVEL', 'NIVEL', "TUE", "NORDESTE", "OESTE", "SOBRAL", "CARIRI", ""}

    # criar DataFrame com as três colunas lado a lado para inspeção de linhas iniciais
    cols_df = pd.concat([col_B, col_C, col_D, col_F, col_G, col_H, col_O, col_P, col_R, col_S, col_T], axis=1)

    # encontrar primeiro índice que não parece ser cabeçalho
    start_idx = 0
    for idx in range(len(cols_df)):
        row_vals = cols_df.iloc[idx].astype(str).str.strip().str.upper().tolist()
        # se todos os valores estão vazios ou em header_labels, considerar como linha de cabeçalho
        all_header = all((v == '' or v in header_labels) for v in row_vals)
        if not all_header:
            start_idx = idx
            break

    if start_idx != 0:
        print(f"Pulei {start_idx} linhas iniciais (prováveis cabeçalhos) antes de coletar os 31 dias.")

    # -------- TUE -----------
    series_TUE_A = col_B.iloc[start_idx:].reset_index(drop=True)
    series_TUE_B = col_C.iloc[start_idx:].reset_index(drop=True)
    series_TUE_C = col_D.iloc[start_idx:].reset_index(drop=True)

    # -------- OESTE -----------
    series_OESTE_A = col_F.iloc[start_idx:].reset_index(drop=True)
    series_OESTE_B = col_G.iloc[start_idx:].reset_index(drop=True)
    series_OESTE_C = col_H.iloc[start_idx:].reset_index(drop=True)

    # -------- NORDESTE -----------
    series_NORDESTE_A = col_J.iloc[start_idx:].reset_index(drop=True)
    series_NORDESTE_B = col_K.iloc[start_idx:].reset_index(drop=True)
    series_NORDESTE_C = col_L.iloc[start_idx:].reset_index(drop=True)

    # -------- SOBRAL -----------
    series_SOBRAL_A = col_N.iloc[start_idx:].reset_index(drop=True)
    series_SOBRAL_B = col_O.iloc[start_idx:].reset_index(drop=True)
    series_SOBRAL_C = col_P.iloc[start_idx:].reset_index(drop=True)

    # -------- CARIRI -----------
    series_CARIRI_A = col_R.iloc[start_idx:].reset_index(drop=True)
    series_CARIRI_B = col_S.iloc[start_idx:].reset_index(drop=True)
    series_CARIRI_C = col_T.iloc[start_idx:].reset_index(drop=True)


    # Abre o arquivo Excel que será preenchido (planilha do Farol)
    open_excel(TARGET_EXCEL, wait_seconds=6.0)

    print("Posicione a janela do Excel em primeiro plano. Execução começará em 5 segundos...")
    time.sleep(5)

    # Aba TUE: Nível A linha 54 (E54..AI54), Nível B linha 55, Nível C linha 56
    # Mantenha dry_run=True por padrão para inspeção; altere para False para execução real.
    dry_run = False
    
    # calcular referência de mês/ano (None -> mês atual)
    REF_MONTH = None
    REF_YEAR = None
    now = datetime.datetime.now()
    if REF_MONTH is None or REF_YEAR is None:
        ref_month = now.month
        ref_year = now.year
    else:
        ref_month = REF_MONTH
        ref_year = REF_YEAR

    if ref_month == now.month and ref_year == now.year:
        days_to_fill = now.day
    else:
        days_to_fill = calendar.monthrange(ref_year, ref_month)[1]

    print("--- DRY RUN: mostrando células que seriam preenchidas ---")
    print(f"Preenchendo até dia {days_to_fill} do mês {ref_month}/{ref_year}")

    regions = [
        {
            'name': "TUE",
            'sheet': "TUE",
            'cols': ('E',),
            'rows': (54, 55, 56),
            'series': (series_TUE_A, series_TUE_B, series_TUE_C),
        },
        {
            'name': "VLT Oeste",
            'sheet': "VLT Oeste",
            'cols': ('F',),
            'rows': (34, 35, 36),
            'series': (series_OESTE_A, series_OESTE_B, series_OESTE_C),
        },
        {
            'name': "VLT Nordeste",
            'sheet': "VLT Nordeste",
            'cols': ('F',),
            'rows': (31, 32, 33),
            'series': (series_NORDESTE_A, series_NORDESTE_B, series_NORDESTE_C),
        },
        {
            'name': "VLT Sobral",
            'sheet': "VLT Sobral",
            'cols': ('F',),
            'rows': (30, 31, 32),
            'series': (series_SOBRAL_A, series_SOBRAL_B, series_SOBRAL_C),
        },
        {
            'name': "VLT Cariri",
            'sheet': "VLT Cariri",
            'cols': ('G',),
            'rows': (28, 29, 30),
            'series': (series_CARIRI_A, series_CARIRI_B, series_CARIRI_C),
        },
    ]

    print("--- DRY RUN: mostrando células que seriam preenchidas ---")
    print(f"Preenchendo até dia {days_to_fill} do mês {ref_month}/{ref_year}")

    for region in regions:
        sheet = region['sheet']
        print(f"\nRegião: {region['name']} (aba: {sheet})")
        for level_idx, row in enumerate(region['rows']):
            series = region['series'][level_idx]
            start_col = region['cols'][0]
            print(f"  Nível {['A','B','C'][level_idx]} -> {start_col}{row} .. (dias: {days_to_fill})")
            # navegar para a primeira célula na aba específica
            if dry_run:
                # em dry-run não navegamos entre abas para evitar efeitos colaterais
                fill_level_from_series(series, start_col_letter=start_col, excel_row=row, dry_run=dry_run, days_to_fill=days_to_fill)
            else:
                go_to_sheet_cell(sheet, f"{start_col}{row}")
                fill_level_from_series(series, start_col_letter=start_col, excel_row=row, dry_run=dry_run, days_to_fill=days_to_fill)

    print("\nPreenchimento (dry-run) concluído. Se estiver tudo correto, altere dry_run para False no script para executar.")

    print("Preenchimento concluído.")
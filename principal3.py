import pandas as pd
from pandas.io.formats import excel
import numpy as np
import openpyxl
import xlsxwriter
from rapidfuzz import process, fuzz

def normalizar(s):
    """
    Normaliza uma série de strings, removendo espaços em branco, caracteres
    especiais e convertendo para minúsculas. Deve ser usada antes de buscar 
    correspondências aproximadas entre strings (rapidfuzz).
    """
    s = s.astype(str).str.lower().str.strip()
    s = s.str.replace(r'[^\w\s]', '', regex=True)
    s = s.str.replace(r'\s+', ' ', regex=True)

    return s.str.strip()


def limpeza(df, planilha = 'transito'):
    """
    Limpeza e pré-processamento do DataFrame.
    """
    if planilha == 'transito':
        # Descarta linhas e colunas com menos de 5 valores não nulos.
        df = df.dropna(thresh=6)
        df = df.dropna(axis=1, thresh=5)

        df = df[df["Mercadoria"] != "Mercadoria"]    # Exclusão de linhas com nomes da coluna

        df['Mercadoria'] = df['Mercadoria'].str.upper()   # Maiúsculas
        df['Tomador'] = df['Tomador'].str.upper()

        # Conversão para numérico
        df['Vlr NF'] = (df['Vlr NF'].str.replace('.', '').str.replace(',', '.')
                        .astype(float))
        df['Valor Frete'] = (df['Valor Frete'].str.replace('.', '').str.replace(',', '.')
                            .astype(float))

        # Classificação de mercadorias em COMBUSTIVEL e VEGETAL
        combustiveis = ['ALCOOL', 'BIODIESEL', 'ETANOL', 'GASOLINA', 'ONU1170',
                        'ONU 1170']
        padrao = '|'.join(combustiveis)
        df['Operação'] = np.where(df['Mercadoria'].str.contains(padrao, na=False),
                            'COMBUSTIVEL', 'VEGETAL')

    else:
        df = df.dropna(axis=1, thresh=5) # Colunas
        df = df.dropna()   # Linhas
        df = df.rename(columns = {'Unnamed: 2':'Empresa', 'Unnamed: 18':'Qtd'})

        df['Empresa'] = df['Empresa'].str.replace('Razão Social : ', '')

        # Conversão para numérico
        df['Receber'] = (df['Receber'].str.replace('.', '').str.replace(',', '.')
                        .astype(float))
        df['Pagar'] = (df['Pagar'].str.replace('.', '').str.replace(',', '.')
                            .astype(float))

    return df


def fuzzymatch(df, similaridade = 87):
    """
    Recebe um DataFrame (df) e retorna o mesmo DataFrame com uma nova coluna 'Empresa',
    que contém os nomes aproximados via correspondência fuzzy.
    Retorna a pontuação de similaridade entre duas strings usando o algoritmo
    de correspondência aproximada (fuzzy matching).
    """
    # Process names from most to least frequent, so the most common
    # spelling becomes the reference and rare typos get attached to it.
    df['Tomador_norm'] = normalizar(df['Tomador'])
    
    counts = df['Tomador_norm'].value_counts()
    mapping, representatives = {}, []
    for name in counts.index:
        if representatives:
            best, score, _ = process.extractOne(
                name, representatives, scorer=fuzz.token_sort_ratio
            )
            if score >= similaridade:
                mapping[name] = best
                continue
        representatives.append(name)   # no close match -> new group
        mapping[name] = name

    # --- 3) REVIEW the proposed merges (important with financial data!) --
    for variant, canonical in mapping.items():
        if variant != canonical:
            print(f'{variant!r}  ->  {canonical!r}')

    # --- 4) aggregate ----------------------------------------------------
    df['Empresa'] = df['Tomador_norm'].map(mapping).str.upper()
    print(df)

    return df


df_t = pd.read_excel('transito.xlsx', sheet_name='ReportXML', header=9)
des = pd.read_excel('descarregados_cnpj.xlsx', header = 8)

df_t = limpeza(df_t)
print(sum(df_t['Vlr NF']), '\n', sum(df_t['Valor Frete']))

#with open('saida.csv', 'w') as f:
#    df_t.to_csv(f, index=False, sep=';', decimal=',', encoding='utf-8')

tab_c = df_t[df_t['Operação'] == 'COMBUSTIVEL']
tab_v = df_t[df_t['Operação'] == 'VEGETAL']

df_c = fuzzymatch(tab_c)
df_v = fuzzymatch(tab_v)

# Planilhas finais
p_c = tab_c.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')
p_c.rename(columns={'Vlr NF': 'Em Transito'}, inplace=True)
p_v = tab_v.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')
p_v.rename(columns={'Vlr NF': 'Em Transito'}, inplace=True)

# Remove a formatação padrão do Pandas para não conflitar com o estilo do Excel
excel.ExcelFormatter.header_style = None

# Arquivo
with pd.ExcelWriter("saida.xlsx", engine="xlsxwriter") as writer:
    workbook = writer.book

    # Formatação monetária (para as colunas normais e para a linha de total)
    moeda = workbook.add_format({"num_format": "R$ #,##0.00"})
    moeda_total = workbook.add_format({"num_format": "R$ #,##0.00", "bold": True})

    # -------------------------------------------------------------
    # 1. ABA: Base
    # -------------------------------------------------------------
    df_t = df_t[
        [
            "Dt. Emissão",
            "Operação",
            "Mercadoria",
            "Tomador",
            "Valor Frete",
            "Vlr NF",
        ]
    ]

    # ATENÇÃO: Enviamos sem cabeçalho (header=False) porque o add_table vai criar o dele
    df_t.to_excel(writer, sheet_name="Base", index=False, startrow=1, header=False)
    worksheet = writer.sheets["Base"]

    max_row_base = len(df_t) + 1  # +1 por causa da linha de cabeçalho
    max_col_base = len(df_t.columns) - 1

    # Cria a tabela com cabeçalho em negrito, listras e LINHA DE TOTAL
    worksheet.add_table(
        0,
        0,
        max_row_base,
        max_col_base,
        {
            "columns": [
                {"header": "Dt. Emissão", "total_string": "Total"},
                {"header": "Operação"},
                {"header": "Mercadoria"},
                {"header": "Tomador"},
                {"header": "Valor Frete", "total_function": "sum",
                 "format": moeda},
                {"header": "Vlr NF", "total_function": "sum", "format": moeda},
            ],
            "total_row": True,  # Ativa a linha de totais no fim da tabela
            "style": "Table Style Light 9",
        },
    )

    # Configuração de colunas e larguras
    worksheet.set_column("A:B", 15)
    worksheet.set_column("C:C", 20)
    worksheet.set_column("D:D", 30)
    worksheet.set_column("E:F", 15, moeda)

    # Aplica negrito apenas nas células de total lá embaixo
    worksheet.write(max_row_base, 4, f"=SUM(E2:E{max_row_base})", moeda_total)
    worksheet.write(max_row_base, 5, f"=SUM(F2:F{max_row_base})", moeda_total)

    # -------------------------------------------------------------
    # 2. ABA: Combustível (Ajustada sem conflito de cor)
    # -------------------------------------------------------------
    p_c.to_excel(
        writer, sheet_name="Combustível", startrow=1, header=False
    )  # Sem index=False pois você usa o index aqui
    worksheet = writer.sheets["Combustível"]

    p_c_reset = p_c.reset_index()
    max_row_c = len(p_c_reset)
    max_col_c = len(p_c_reset.columns) - 1

    worksheet.add_table(
        0,
        0,
        max_row_c,
        max_col_c,
        {
            "columns": [{"header": col} for col in p_c_reset.columns],
            "style": "Table Style Light 9",
        },
    )
    worksheet.set_column("A:A", 30)
    worksheet.set_column("B:B", 15, moeda)

    # -------------------------------------------------------------
    # 3. ABA: Vegetal (Ajustada sem conflito de cor)
    # -------------------------------------------------------------
    p_v.to_excel(writer, sheet_name="Vegetal", startrow=1, header=False)
    worksheet = writer.sheets["Vegetal"]

    p_v_reset = p_v.reset_index()
    max_row_v = len(p_v_reset)
    max_col_v = len(p_v_reset.columns) - 1

    worksheet.add_table(
        0,
        0,
        max_row_v,
        max_col_v,
        {
            "columns": [{"header": col} for col in p_v_reset.columns],
            "style": "Table Style Light 9",
        },
    )
    worksheet.set_column("A:A", 30)
    worksheet.set_column("B:B", 15, moeda)

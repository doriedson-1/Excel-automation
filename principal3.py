# Criando a couna 'EM TRÂNSITO'
import pandas as pd
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


def limpeza(df):
    """
    Limpeza e pré-processamento do DataFrame.
    """
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
    combustiveis = ['ALCOOL', 'BIODIESEL', 'ETANOL', 'GASOLINA']
    padrao = '|'.join(combustiveis)
    df['Operação'] = np.where(df['Mercadoria'].str.contains(padrao, na=False),
                        'COMBUSTIVEL', 'VEGETAL')

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


df = pd.read_excel('transito.xlsx', sheet_name='ReportXML', header=9)
df = limpeza(df)
print(sum(df['Vlr NF']), '\n', sum(df['Valor Frete']))

tab_c = df[df['Operação'] == 'COMBUSTIVEL']
tab_v = df[df['Operação'] == 'VEGETAL']

df_c = fuzzymatch(tab_c)
df_v = fuzzymatch(tab_v)

#(df.head(10))
# with open ('saida.txt', 'w') as f:
#     df = df[['Dt. Emissão', 'Operação',  'Tomador',
#              'Empresa', 'Valor Frete', 'Vlr NF']]
#     f.write(df.to_string())

# Planilhas finais
p_c = tab_c.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')
p_v = tab_v.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')

# #result = df.groupby('Empresa', as_index=False)['Vlr NF'].sum()
#print(result)



 
# Arquivo
with pd.ExcelWriter('saida.xlsx') as writer:
    # Formatação monetária
    workbook = writer.book
    moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

    df = df[['Dt. Emissão', 'Operação', 'Mercadoria', 'Tomador',
                     'Empresa', 'Valor Frete', 'Vlr NF']]
    df.to_excel(writer, sheet_name='Base', index=False)
    worksheet = writer.sheets['Base']
    worksheet.set_column('A:B', 15)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('E:E', 15, moeda)
    worksheet.set_column('F:F', 15, moeda)

    p_c.to_excel(writer, sheet_name='Combustível')
    worksheet = writer.sheets['Combustível']
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 15, moeda)

    p_v.to_excel(writer, sheet_name='Vegetal')
    worksheet = writer.sheets['Vegetal']
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 15, moeda)

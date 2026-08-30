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

        df['Empresa'] = df['Empresa'].str.replace('Razão Social : ', '').str.upper()

        # Conversão para numérico
        df['Receber'] = (df['Receber'].str.replace('.', '').str.replace(',', '.')
                        .astype(float))
        df['Pagar'] = (df['Pagar'].str.replace('.', '').str.replace(',', '.')
                            .astype(float))

    return df


def fuzzymatch(df, similaridade = 87, coluna_empresa = 'Tomador'):
    """
    Recebe um DataFrame (df) e retorna o mesmo DataFrame com uma nova coluna 'Empresa',
    que contém os nomes aproximados via correspondência fuzzy.
    Retorna a pontuação de similaridade entre duas strings usando o algoritmo
    de correspondência aproximada (fuzzy matching).
    """
    # Process names from most to least frequent, so the most common
    # spelling becomes the reference and rare typos get attached to it.
    df['Tomador_norm'] = normalizar(df[coluna_empresa])
    
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
des = limpeza(des, planilha='descarregados')
print(sum(df_t['Vlr NF']), '\n', sum(df_t['Valor Frete']))

def classificar_descarregados(df):
    """
    Comparação dos tipos de operação para cada empresa, a partir daqui
    será feita a classificação das empresas da planilha de 'descarregados_cnpj'.
    Recebe o dataframe originado da planilha 'transito'.
    Retorna um dataframe ('nomes_valores') com a operação mais feita para cada
    empresa (não há padronização dos nomes).
    """

    nomes_valores = (df.pivot_table(index='Tomador', values='Vlr NF',
                                    aggfunc='sum', columns='Operação', fill_value=0)
                                    .rename_axis(index=None, columns=None)
                    )
    nomes_valores = nomes_valores.reset_index()
    nomes_valores = nomes_valores.rename(columns = {'index':'Empresa'})
    nomes_valores['Maioria'] = nomes_valores[['COMBUSTIVEL', 'VEGETAL']].idxmax(axis=1)

    return nomes_valores[['Empresa', 'Maioria']]


nomes_op = classificar_descarregados(df_t)

des_completo = (des.merge(nomes_op, on='Empresa', how='outer')
                     .drop(['Pagar', 'Qtd'], axis=1))
des_completo['Maioria'] = des_completo['Maioria'].fillna('OUTROS')
df2 = (fuzzymatch(des_completo, coluna_empresa='Empresa')
       .drop('Tomador_norm', axis=1))
#lista = df2['Empresa'].unique()
#print(len(lista))

tabela_des = pd.pivot_table(
    df2,
    index='Empresa',
    columns='Maioria',
    values='Receber',
    aggfunc='sum',
    fill_value=0
)

descarregados = pd.DataFrame({
    'Descarregados': tabela_des.sum(axis=1),
    'Operação': tabela_des.idxmax(axis=1)
}).reset_index()

#with open('saida.csv', 'w') as f:
#    df_t.to_csv(f, index=False, sep=';', decimal=',', encoding='utf-8')

des_c = descarregados[descarregados['Operação'] == 'COMBUSTIVEL']
des_v = descarregados[descarregados['Operação'] == 'VEGETAL']
des_o = descarregados[descarregados['Operação'] == 'OUTROS']

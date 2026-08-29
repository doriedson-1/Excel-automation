import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter

df = pd.read_excel('base.xlsx', sheet_name='ReportXML', header=9)

# Descarta linhas e colunas com menos de 5 valores não nulos.
df = df.dropna(thresh=6)
df = df.dropna(axis=1, thresh=5)

df = df[df["Mercadoria"] != "Mercadoria"]    # Exclusão de linhas com nomes da coluna

df['Mercadoria'] = df['Mercadoria'].str.upper()   # Maiúsculas

# Conversão para numérico
df['Vlr NF'] = (df['Vlr NF'].str.replace('.', '').str.replace(',', '.')
                .astype(float))
df['Valor Frete'] = (df['Valor Frete'].str.replace('.', '').str.replace(',', '.')
                     .astype(float))

# Classificação de mercadorias em COMBUSTIVEL e VEGETAL
combustiveis = ['ALCOOL', 'BIODIESEL', 'ETANOL', 'GASOLINA']
padrao = '|'.join(combustiveis)
df['Produto'] = np.where(df['Mercadoria'].str.contains(padrao, na=False),
                      'COMBUSTIVEL', 'VEGETAL')

tab_c = df[df['Produto'] == 'COMBUSTIVEL']
tab_v = df[df['Produto'] == 'VEGETAL']

# Planilhas finais
p_c = tab_c.pivot_table(index=['Tomador'], values = ['Vlr NF'], aggfunc='sum')
p_v = tab_v.pivot_table(index=['Tomador'], values = ['Vlr NF'], aggfunc='sum')
 
# Arquivo
with pd.ExcelWriter('saida.xlsx') as writer:
    # Formatação monetária
    workbook = writer.book
    moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

    df_limpo = df[['Dt. Emissão', 'Produto', 'Mercadoria', 'Tomador',
                     'Valor Frete', 'Vlr NF']]
    df_limpo.to_excel(writer, sheet_name='Base', index=False)
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

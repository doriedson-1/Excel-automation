import os
import re
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
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
        df['Empresa'] = df['Empresa'].str.upper()

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


def fuzzymatch_cruzado(df_origem, df_destino, col_origem='Empresa',
                       col_destino='Empresa', similaridade=87):
    """
    Compara a coluna de empresas do df_destino contra a lista de empresas oficiais do df_origem.
    
    Retorna:
    - df_destino atualizado com a coluna 'Empresa_Padronizada'
    - Um dicionário com os detalhes do match (Score e Origem) para a auditoria do cliente.
    """
    # 1. Obter lista de empresas únicas de referência (Trânsito)
    referencias_norm = df_origem[col_origem].dropna().unique()
    
    # Dicionário de mapeamento para rápida consulta
    mapa_referencias = {normalizar(pd.Series([ref]))[0]: ref for ref in referencias_norm}
    chaves_referencia = list(mapa_referencias.keys())
    
    mapping_resultados = {}
    
    # 2. Iterar sobre as empresas da planilha 'descarregados'
    empresas_destino = df_destino[col_destino].dropna().unique()
    
    for emp_raw in empresas_destino:
        emp_norm = normalizar(pd.Series([emp_raw]))[0]
        
        # Caso 1: Nome idêntico ao de referência
        if emp_norm in mapa_referencias:
            mapping_resultados[emp_raw] = {
                'Empresa_Padronizada': mapa_referencias[emp_norm],
                'Score': 100,
                'Status_Match': 'Exato'
            }
        # Caso 2: Busca por similaridade (Fuzzy Match)
        elif chaves_referencia:
            best_match, score, _ = process.extractOne(
                emp_norm, chaves_referencia, scorer=fuzz.token_sort_ratio
            )
            
            if score >= similaridade:
                mapping_resultados[emp_raw] = {
                    'Empresa_Padronizada': mapa_referencias[best_match],
                    'Score': round(score, 1),
                    'Status_Match': 'Fuzzy Match'
                }
            else:
                # Score abaixo do limite -> Empresa não encontrada na 'transito'
                mapping_resultados[emp_raw] = {
                    'Empresa_Padronizada': emp_raw,
                    'Score': round(score, 1),
                    'Status_Match': 'Sem Match (Novo)'
                }
        else:
            mapping_resultados[emp_raw] = {
                'Empresa_Padronizada': emp_raw,
                'Score': 0,
                'Status_Match': 'Sem Match (Novo)'
            }

    # 3. Mapear os resultados de volta para o DataFrame destino
    df_destino['Empresa_Padronizada'] = df_destino[col_destino].map(
        lambda x: mapping_resultados.get(x, {}).get('Empresa_Padronizada', x)
    )
    #df_destino['Match_Score'] = df_destino[col_destino].map(
    #    lambda x: mapping_resultados.get(x, {}).get('Score', 0)
    #)
    #df_destino['Match_Status'] = df_destino[col_destino].map(
    #    lambda x: mapping_resultados.get(x, {}).get('Status_Match', 'Sem Match')
    #)

    return df_destino, mapping_resultados


def gerenciar_classificacao_manual(
    df_transito, df_descarregados, arquivo_pendentes="Classificar_Pendentes.xlsx"
):
    """Atualiza a planilha 'Classificar_Pendentes.xlsx' criando um menu suspenso

    (validação de dados) na coluna Categoria com as opções: VEGETAL, COMBUSTIVEL
    e OUTROS.

    Retorna o dicionário consolidado {Empresa: Categoria}.
    """
    # 1. Categoria base vinda da 'transito.xlsx'
    mapa_transito = (
        df_transito.groupby("Empresa_Padronizada")["Operação"].last().to_dict()
    )

    # 2. Ler histórico do cliente (se o arquivo já existir)
    mapa_manual = {}
    if os.path.exists(arquivo_pendentes):
        try:
            df_pendentes_existente = pd.read_excel(arquivo_pendentes)
            for _, row in df_pendentes_existente.iterrows():
                emp = str(row["Empresa"]).strip()
                cat = str(row["Categoria"]).strip()
                if cat in ["VEGETAL", "COMBUSTIVEL", "OUTROS"]:
                    mapa_manual[emp] = cat
        except Exception as e:
            print(f"Aviso: Não foi possível ler o histórico de pendências: {e}")

    # 3. Cruzamento e consolidação
    registros_pendentes = []
    mapa_consolidado = {}

    empresas_descarregados = df_descarregados[
        "Empresa_Padronizada"
    ].unique()

    for emp_padronizada in empresas_descarregados:
        if emp_padronizada in mapa_transito:
            categoria_final = mapa_transito[emp_padronizada]
        elif emp_padronizada in mapa_manual:
            categoria_final = mapa_manual[emp_padronizada]
        else:
            categoria_final = "OUTROS"

        mapa_consolidado[emp_padronizada] = categoria_final
        registros_pendentes.append(
            {"Empresa": emp_padronizada, "Categoria": categoria_final}
        )

    # 4. Criar o DataFrame e ordenar para manter 'OUTROS' no topo
    df_pendentes = pd.DataFrame(registros_pendentes)
    df_pendentes["Ordem"] = df_pendentes["Categoria"].apply(
        lambda x: 0 if x == "OUTROS" else 1
    )
    df_pendentes = df_pendentes.sort_values(by=["Ordem", "Empresa"]).drop(
        columns=["Ordem"]
    )

    # 5. Salvar o arquivo base via Pandas
    df_pendentes.to_excel(arquivo_pendentes, index=False)

    # 6. Adicionar a Validação de Dados (Menu Suspenso) via openpyxl
    wb = openpyxl.load_workbook(arquivo_pendentes)
    ws = wb.active

    # Criar a regra de validação em lista
    dv = DataValidation(
        type="list", formula1='"VEGETAL,COMBUSTIVEL,OUTROS"', allow_blank=True
    )
    dv.error = "Escolha uma das opções da lista: VEGETAL, COMBUSTIVEL ou OUTROS"
    dv.errorTitle = "Opção Inválida"
    dv.prompt = "Selecione a categoria da empresa"
    dv.promptTitle = "Categoria"

    ws.add_data_validation(dv)

    # Aplicar a validação em toda a Coluna B (Categoria), da linha 2 até o final
    max_row = len(df_pendentes) + 1
    if max_row >= 2:
        dv.add(f"B2:B{max_row}")

    # Ajustar largura das colunas para melhor visualização
    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 20

    wb.save(arquivo_pendentes)

    return mapa_consolidado


def normalizar_avancado(texto):
    """Remove pontuações isoladas, pontuações entre letras (ex: S.A. -> SA)

    e padroniza os espaços.
    """
    if not isinstance(texto, str):
        return ""

    texto = texto.upper().strip()

    # 1. Substitui pontuações comuns (ponto, vírgula, ponto e vírgula, hífen, barra) por espaço
    texto = re.sub(r"[.,;\-/]", " ", texto)

    # 2. Remove múltiplos espaços gerados pelas substituições
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def fuzzymatch_consolidação(
    df, col_empresa="Empresa", similaridade_corte=85):
    """Aplica o agrupamento fuzzy nos dados consolidados, unificando variações
    sutis de nomes de empresas.
    """
    df_copia = df.copy()

    # 1. Cria coluna de nome limpo para comparação
    empresas_unicas = df_copia[col_empresa].dropna().unique()

    # Mapeamento do nome original para a versão limpa
    nomes_limpos = {emp: normalizar_avancado(emp) for emp in empresas_unicas}

    # 2. Identificação de grupos de nomes equivalentes
    mapeamento_canonico = {}
    representantes_limpos = []
    mapa_limpo_para_original = {}

    # Ordena por frequência (se houver repetição) ou tamanho para usar o nome mais completo como padrão
    for emp_orig in empresas_unicas:
        emp_limpa = nomes_limpos[emp_orig]

        if not emp_limpa:
            continue

        if representantes_limpos:
            # token_set_ratio ignora diferenças causadas por pontuação e ordenação de tokens
            best_match, score, _ = process.extractOne(
                emp_limpa,
                representantes_limpos,
                scorer=fuzz.token_set_ratio,
            )

            if score >= similaridade_corte:
                # Mapeia a variação para o nome canônico existente
                nome_canonico = mapa_limpo_para_original[best_match]
                mapeamento_canonico[emp_orig] = nome_canonico
                continue

        # Novo grupo identificado
        representantes_limpos.append(emp_limpa)
        mapa_limpo_para_original[emp_limpa] = emp_orig
        mapeamento_canonico[emp_orig] = emp_orig

    # 3. Atualiza a coluna de empresa no DataFrame
    df_copia[col_empresa] = df_copia[col_empresa].map(mapeamento_canonico)

    return df_copia


# Importação
df_t = pd.read_excel('transito.xlsx', sheet_name='ReportXML', header=9)
des = pd.read_excel('descarregados_cnpj.xlsx', header = 8)

# Limpeza
df_t = limpeza(df_t)
des = limpeza(des, planilha='descarregados')
print(sum(df_t['Vlr NF']), '\n', sum(df_t['Valor Frete']))

# Padroniza os nomes dentro da própria 'transito' (usando sua fuzzymatch original)
df_t = fuzzymatch(df_t)
df_t["Empresa_Padronizada"] = df_t["Tomador"]
#with open('df_transito.csv', 'w') as f:
#    df_t.to_csv(f, index=False, sep=';', decimal=',', encoding='utf-8')

des, mapping_resultados = fuzzymatch_cruzado(
    df_origem=df_t, df_destino=des, col_origem="Empresa", col_destino="Empresa")

# Obtém o dicionário {Empresa: Categoria} e atualiza o arquivo 'Classificar_Pendentes.xlsx'
mapa_categorias = gerenciar_classificacao_manual(df_t, des)

# Aplica a categoria correspondente em cada linha da 'descarregados'
des["Operação"] = des["Empresa_Padronizada"].map(mapa_categorias)

# A partir daqui, separação das empresas conforme a coluna 'Operação'
#des_v = des[des["Operação"] == "VEGETAL"]
#des_c = des[des["Operação"] == "COMBUSTIVEL"]
#des_o = des[des["Operação"] == "OUTROS"]

def gerar_pivot_categoria(df_descarregados, df_transito, categoria):
    """Gera a tabela consolidada (Descarregados x Em Trânsito) por empresa

    para a categoria informada ('VEGETAL', 'COMBUSTIVEL' ou 'OUTROS').
    """
    # 1. Filtra os dataframes pela categoria correspondente
    sub_des = df_descarregados[df_descarregados["Operação"] == categoria]
    sub_tra = df_transito[df_transito["Operação"] == categoria]

    # 2. Agrupa 'Descarregados' por Empresa Padronizada (Somando 'Receber')
    if not sub_des.empty:
        p_des = sub_des.pivot_table(
            index="Empresa_Padronizada", values="Receber", aggfunc="sum"
        ).rename(columns={"Receber": "Descarregados"})
    else:
        p_des = pd.DataFrame(columns=["Descarregados"])
        p_des.index.name = "Empresa_Padronizada"

    # 3. Agrupa 'Em Trânsito' por Empresa Padronizada (Somando 'Vlr NF')
    if not sub_tra.empty:
        p_tra = sub_tra.pivot_table(
            index="Empresa_Padronizada", values="Vlr NF", aggfunc="sum"
        ).rename(columns={"Vlr NF": "Em Trânsito"})
    else:
        p_tra = pd.DataFrame(columns=["Em Trânsito"])
        p_tra.index.name = "Empresa_Padronizada"

    # 4. Une as duas visões (Outer Join) e preenche valores ausentes com 0.0
    resumo = pd.merge(
        p_des, p_tra, on="Empresa_Padronizada", how="outer"
    ).fillna(0.0)
    resumo = resumo.reset_index().rename(
        columns={"Empresa_Padronizada": "Empresa"}
    )

    return resumo


# ==============================================================================
# CONSOLIDAÇÃO DOS DADOS
# ==============================================================================
p_c = gerar_pivot_categoria(des, df_t, "COMBUSTIVEL")
p_v = gerar_pivot_categoria(des, df_t, "VEGETAL")
p_o = gerar_pivot_categoria(des, df_t, "OUTROS")

# ==============================================================================
# GRAVAÇÃO NA PLANILHA FINAL VIA XLSXWRITER
# ==============================================================================
with pd.ExcelWriter("FINAL.xlsx", engine="xlsxwriter") as writer:
    workbook = writer.book

    # Formatações monetárias
    moeda = workbook.add_format({"num_format": "R$ #,##0.00"})
    moeda_total = workbook.add_format(
        {"num_format": "R$ #,##0.00", "bold": True}
    )

    # -------------------------------------------------------------
    # 1. ABA: Base (Manutenção do seu código original do Trânsito)
    # -------------------------------------------------------------
    df_t_base = df_t[
        [
            "Dt. Emissão",
            "Operação",
            "Mercadoria",
            "Tomador",
            "Valor Frete",
            "Vlr NF",
        ]
    ]

    df_t_base.to_excel(
        writer, sheet_name="Base", index=False, startrow=1, header=False
    )
    ws_base = writer.sheets["Base"]

    max_row_base = len(df_t_base) + 1
    max_col_base = len(df_t_base.columns) - 1

    ws_base.add_table(
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
                {
                    "header": "Valor Frete",
                    "total_function": "sum",
                    "format": moeda,
                },
                {"header": "Vlr NF", "total_function": "sum", "format": moeda},
            ],
            "total_row": True,
            "style": "Table Style Light 9",
        },
    )

    ws_base.set_column("A:B", 15)
    ws_base.set_column("C:C", 20)
    ws_base.set_column("D:D", 30)
    ws_base.set_column("E:F", 15, moeda)

    ws_base.write(max_row_base, 4, f"=SUM(E2:E{max_row_base})", moeda_total)
    ws_base.write(max_row_base, 5, f"=SUM(F2:F{max_row_base})", moeda_total)

    # -------------------------------------------------------------
    # 2. ESTRUTURAÇÃO DAS ABAS: Combustível, Vegetal e Outros
    # -------------------------------------------------------------
    abas_config = [
        ("Combustível", p_c),
        ("Vegetal", p_v),
        ("Outros", p_o),
    ]

    for nome_aba, df_pivot in abas_config:
        # Escreve os dados na aba correspondente
        df_pivot.to_excel(
            writer, sheet_name=nome_aba, index=False, startrow=1, header=False
        )
        ws = writer.sheets[nome_aba]

        max_row = len(df_pivot) + 1
        max_col = len(df_pivot.columns) - 1

        if len(df_pivot) > 0:
            ws.add_table(
                0,
                0,
                max_row,
                max_col,
                {
                    "columns": [
                        {"header": "Empresa", "total_string": "Total"},
                        {
                            "header": "Descarregados",
                            "total_function": "sum",
                            "format": moeda,
                        },
                        {
                            "header": "Em Trânsito",
                            "total_function": "sum",
                            "format": moeda,
                        },
                    ],
                    "total_row": True,
                    "style": "Table Style Light 9",
                },
            )

            # Formatações de largura e estilo
            ws.set_column("A:A", 40)
            ws.set_column("B:C", 18, moeda)

            # Fórmulas da linha de total
            ws.write(max_row, 1, f"=SUM(B2:B{max_row})", moeda_total)
            ws.write(max_row, 2, f"=SUM(C2:C{max_row})", moeda_total)
        else:
            # Caso a aba esteja vazia na execução
            ws.write(0, 0, "Nenhum registro encontrado para esta categoria.")


# tran_c = df_t[df_t['Operação'] == 'COMBUSTIVEL']
# tran_v = df_t[df_t['Operação'] == 'VEGETAL']

# df_c = fuzzymatch(tran_c)
# df_v = fuzzymatch(tran_v)

# # Planilhas finais
# p_c = tran_c.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')
# p_c.rename(columns={'Vlr NF': 'Em Transito'}, inplace=True)
# p_v = tran_v.pivot_table(index=['Empresa'], values = ['Vlr NF'], aggfunc='sum')
# p_v.rename(columns={'Vlr NF': 'Em Transito'}, inplace=True)

# # Arquivo
# with pd.ExcelWriter("saida.xlsx", engine="xlsxwriter") as writer:
#     workbook = writer.book

#     # Formatação monetária (para as colunas normais e para a linha de total)
#     moeda = workbook.add_format({"num_format": "R$ #,##0.00"})
#     moeda_total = workbook.add_format({"num_format": "R$ #,##0.00", "bold": True})

#     # -------------------------------------------------------------
#     # 1. ABA: Base
#     # -------------------------------------------------------------
#     df_t = df_t[
#         [
#             "Dt. Emissão",
#             "Operação",
#             "Mercadoria",
#             "Tomador",
#             "Valor Frete",
#             "Vlr NF",
#         ]
#     ]

#     # ATENÇÃO: Enviamos sem cabeçalho (header=False) porque o add_table vai criar o dele
#     df_t.to_excel(writer, sheet_name="Base", index=False, startrow=1, header=False)
#     worksheet = writer.sheets["Base"]

#     max_row_base = len(df_t) + 1  # +1 por causa da linha de cabeçalho
#     max_col_base = len(df_t.columns) - 1

#     # Cria a tabela com cabeçalho em negrito, listras e LINHA DE TOTAL
#     worksheet.add_table(
#         0,
#         0,
#         max_row_base,
#         max_col_base,
#         {
#             "columns": [
#                 {"header": "Dt. Emissão", "total_string": "Total"},
#                 {"header": "Operação"},
#                 {"header": "Mercadoria"},
#                 {"header": "Tomador"},
#                 {"header": "Valor Frete", "total_function": "sum",
#                  "format": moeda},
#                 {"header": "Vlr NF", "total_function": "sum", "format": moeda},
#             ],
#             "total_row": True,  # Ativa a linha de totais no fim da tabela
#             "style": "Table Style Light 9",
#         },
#     )

#     # Configuração de colunas e larguras
#     worksheet.set_column("A:B", 15)
#     worksheet.set_column("C:C", 20)
#     worksheet.set_column("D:D", 30)
#     worksheet.set_column("E:F", 15, moeda)

#     # Aplica negrito apenas nas células de total lá embaixo
#     worksheet.write(max_row_base, 4, f"=SUM(E2:E{max_row_base})", moeda_total)
#     worksheet.write(max_row_base, 5, f"=SUM(F2:F{max_row_base})", moeda_total)

#     # -------------------------------------------------------------
#     # 2. ABA: Combustível (Ajustada sem conflito de cor)
#     # -------------------------------------------------------------
#     p_c.to_excel(
#         writer, sheet_name="Combustível", startrow=1, header=False
#     )  # Sem index=False pois você usa o index aqui
#     worksheet = writer.sheets["Combustível"]

#     p_c_reset = p_c.reset_index()
#     max_row_c = len(p_c_reset)
#     max_col_c = len(p_c_reset.columns) - 1

#     worksheet.add_table(
#         0,
#         0,
#         max_row_c,
#         max_col_c,
#         {
#             "columns": [{"header": col} for col in p_c_reset.columns],
#             "style": "Table Style Light 9",
#         },
#     )
#     worksheet.set_column("A:A", 30)
#     worksheet.set_column("B:B", 15, moeda)

#     # -------------------------------------------------------------
#     # 3. ABA: Vegetal (Ajustada sem conflito de cor)
#     # -------------------------------------------------------------
#     p_v.to_excel(writer, sheet_name="Vegetal", startrow=1, header=False)
#     worksheet = writer.sheets["Vegetal"]

#     p_v_reset = p_v.reset_index()
#     max_row_v = len(p_v_reset)
#     max_col_v = len(p_v_reset.columns) - 1

#     worksheet.add_table(
#         0,
#         0,
#         max_row_v,
#         max_col_v,
#         {
#             "columns": [{"header": col} for col in p_v_reset.columns],
#             "style": "Table Style Light 9",
#         },
#     )
#     worksheet.set_column("A:A", 30)
#     worksheet.set_column("B:B", 15, moeda)

import pandas as pd
import requests
from io import BytesIO
from flask import Flask, send_file
import os

# INICIANDO O FLASK
app = Flask(__name__)
#################################################################################################################################################
# FORMULÁRIO 3 KOBO


# Rota para o download do arquivo
@app.route('/downloadformtres')
def download_form3():
    link = 'https://eu.kobotoolbox.org/api/v2/assets/aKetFYPwuhGSD43m7RtKot/export-settings/esQ8BrbzfxRJqdvJJGX5FRr/data.xlsx?format=api'

    response = requests.get(link)
    file = BytesIO(response.content)

    # lendo as tabelas
    df1 = pd.read_excel(file, sheet_name='Formulário 03 - Cadastro anu...', engine='openpyxl')
    df2 = pd.read_excel(file, sheet_name='Dados_sociais_ufp', engine='openpyxl')
    df3 = pd.read_excel(file, sheet_name='dados_de_producao', engine='openpyxl')

    # convertendo as tabelas para formato "string" para evitar erros na mesclagem dos dados
    df1 = df1.astype(str)
    df2 = df2.astype(str)
    df3 = df3.astype(str)

    # mesclando as tabelas
    df_merged1 = pd.merge(df1, df2, left_on='_index', right_on='_parent_index', how='outer')
    df_merged2 = pd.merge(df_merged1, df3, left_on='_parent_index', right_on='_parent_index', how='outer')

    # função para garantir que apenas valores distintos serão concetenados entre "," nas linhas
    def join_unique(x):
        values = set()
        result = []
        for value in x.dropna():
            if value not in values:
                result.append(value)
                values.add(value)
        return ','.join(result)

    # agrupando os dados que tem linhas a mais pelo "index" e concatenando por ","
    df_merged1 = df_merged2.groupby('_parent_index').agg(join_unique).reset_index()


    # Salve o dataframe em um arquivo Excel
    output_path = os.path.abspath('Formulario 03 - Cadastro anual das Unidades Familiares de Producao.xlsx')
    df_merged1.to_excel(output_path, index=False)

    # Envie o arquivo Excel como resposta HTTP
    return send_file(output_path, as_attachment=True)
################################################################################################################################################

# FORMULÁRIO 6 KOBO

@app.route('/downloadformseis')
def download_form6():
    link = 'https://eu.kobotoolbox.org/api/v2/assets/ah4QnXrsbDCrg7TekPzMHs/export-settings/esJr6AaHsWVbXPpgQtDpUEb/data.xlsx'

    response = requests.get(link)
    file = BytesIO(response.content)

    # lendo a tabela excel
    df1 = pd.read_excel(file, sheet_name='Formulário 06 - Projetos de ...', engine='openpyxl')
    df2 = pd.read_excel(file, sheet_name='culturas_banco_da_amazonia', engine='openpyxl')
    df3 = pd.read_excel(file, sheet_name='culturas_outros_bancos', engine='openpyxl')

    # convertendo as colunas para string
    df1 = df1.astype(str)
    df2 = df2.astype(str)
    df3 = df3.astype(str)

    # mesclando os dados
    df_merged1 = pd.merge(df1, df2, left_on='_uuid', right_on='_submission__uuid', how='left')
    df_merged2 = pd.merge(df_merged1, df3, left_on='_uuid', right_on='_submission__uuid', how='left')

     # função para garantir que apenas valores distintos serão concetenados entre "," nas linhas
    def join_unique(x):
        values = set()
        result = []
        for value in x.dropna():
            if value not in values:
                result.append(value)
                values.add(value)
        return ','.join(result)

    # agrupando os dados que tem linhas a mais pelo "index" e concatenando por ","
    df_merged1 = df_merged2.groupby('_uuid').agg(join_unique).reset_index()

    # Salva o dataframe em um arquivo Excel
    output_path = os.path.abspath('Formulário 06 - Projetos de Crédito efetivados por Agência e Produtor.xlsx')
    df_merged1.to_excel(output_path, index=False)

   # Envie o arquivo Excel como resposta HTTP
    return send_file(output_path, as_attachment=True)

############################################################################################################################################################################

# CONEXÃO MARAJÓ

@app.route('/conexaomarajo')
def download_marajo():
    link = 'https://eu.kobotoolbox.org/api/v2/assets/a7NSRkpn27UsnzQiJX9ayt/export-settings/esEn8uiAwb9EHF8y7cctoCw/data.xlsx'
    response = requests.get(link)
    file = BytesIO(response.content)

    df1 = pd.read_excel(file, sheet_name='Cadastro Socioprodutivo - Co...', engine='openpyxl', thousands='.', decimal=',')
    df2 = pd.read_excel(file, sheet_name='demais_membros_da_ufp', engine='openpyxl', thousands='.', decimal=',')
    df3 = pd.read_excel(file, sheet_name='dados_producao_ufp', engine='openpyxl', thousands='.', decimal=',')
    df4 = pd.read_excel(file, sheet_name='begin_repeat_WAhmxmpIK', engine='openpyxl', thousands='.', decimal=',')

    # Convertendo as colunas para string
    df1 = df1.astype(str)
    df2 = df2.astype(str)
    df3 = df3.astype(str)
    df4 = df4.astype(str)

    # Mesclando as tabelas
    df_merged1 = pd.merge(df1, df2, left_on='_index', right_on='_parent_index', how='outer', suffixes=('_df1', '_df2'))
    df_merged2 = pd.merge(df_merged1, df3, left_on='_parent_index', right_on='_parent_index', how='outer',suffixes=('_merged1', '_df3'))
    df_merged3 = pd.merge(df_merged2, df4, left_on='_parent_index', right_on='_parent_index', how='outer',suffixes=('_merged2', '_df4'))

    # Função para garantir que apenas valores distintos serão concatenados com ","
    def join_unique(x):
        values = set()
        result = []
        for value in x.dropna():
            if value not in values:
                result.append(value)
                values.add(value)
        return ','.join(result)

    # Agrupando os dados que têm linhas extras pelo "index" e concatenando com ","
    df_merged1 = df_merged2.groupby('_parent_index').agg(join_unique).reset_index()
    df_merged2 = df_merged3.groupby('_parent_index').agg(join_unique).reset_index()

    # Salve o dataframe em um arquivo Excel
    output_path = os.path.abspath('Conexão Marajó.xlsx')
    df_merged2.to_excel(output_path, index=False)

    # Envie o arquivo Excel como resposta HTTP
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run()








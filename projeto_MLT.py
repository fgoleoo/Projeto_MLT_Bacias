# codigo começa aqui
import pandas as pd
import xlwings as xw
from datetime import date

# variáveis globais
filename_postos = r'\postos_bacias.dat'
filename_vazoes = r'\vazoes.dat'
path = r'H:\Comercializadora\Preços\Leonardo\Projetos_Python\Projeto_MLT_Bacias'


def main():
    # dicionario de meses
    dict_month = {'jan': '1', 'fev': '2', 'mar': '3', 'abril': '4', 'maio': '5', 'jun': '6', 'jul': '7', 'ago': '8',
                  'set': '9', 'out': '10', 'nov': '11', 'dez': '12'}

    # funcoes que leem os arquivos base (vazoes.dat e postos_bacias.dat)
    df_vazoes_base, df_postos_prod = read_vazoes(path, filename_vazoes, dict_month), read_postos_prod(path,
                                                                                                      filename_postos)

    # funcao que faz o merge entre os dois dataframes iniciais
    df_vazoes = merge(df_vazoes_base, df_postos_prod)

    # funcao que faz os cálculos necessários de ENA em cada posto/bacia
    df_ena, df_ena_bacias = calcula_ena(df_vazoes, dict_month)

    # funcao que me traz a tabela final das estatisticas
    estatisticas_finais = final_stat(df_ena_bacias)

    # faz as substituições para que o excel entenda os valores corretos
    estatisticas_finais['CONCATENADO'] = estatisticas_finais['bacia'] + '_' + estatisticas_finais['index'].apply(
        lambda x: 'MLT' if x == 'mean' else x)

    # funcao que conecta ao excel e passa os parametros necessários
    app, wb, ws_estatisticas = connect_to_excel()

    # limpa a aba escolhida e cola os valores do dataframe com os resultados/altera a primeira linha dos dados para facilitar a leitura pelo excel
    ws_estatisticas.clear_contents()
    ws_estatisticas.range('A1').value = round(estatisticas_finais, 0)
    ws_estatisticas.range('A1:O1').value = ['index', 'bacia', 'medidas', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                                            '10', '11', '12']

    wb.save()
    wb.close()
    app.quit()


########################################### INICIO DAS FUNCOES #########################################################


# funcao que le o arquivo de vazoes
def read_vazoes(path, filename_vazoes, dict_month):
    # Criar a lista de headers
    list_headers = list(dict_month.keys())
    list_headers.insert(0, 'ano')
    list_headers.insert(0, 'posto')

    # ['posto', 'ano', 'jan', 'fev', 'mar', 'abril', 'maio', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']

    # faz a leitura do arquivo .dat
    df_vazoes_base = pd.read_fwf(path + filename_vazoes, header=None, widths=[3, 5, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6],
                                 names=list_headers)

    return df_vazoes_base


# funcao que le o arquivo de postos e produtibilidade
def read_postos_prod(path, filename_postos):
    # faz a leitura do arquivo .dat
    df_postos_prod = pd.read_fwf(path + filename_postos, header=None, widths=[4, 7, 25],
                                 names=['posto', 'prod', 'bacia'])

    # troca virgula por ponto
    df_postos_prod['prod'] = df_postos_prod['prod'].str.replace(',', '.')

    # altera de string para float
    new_dtype = {"prod": float}
    df_postos_prod = df_postos_prod.astype(new_dtype)

    return df_postos_prod


# funcao que cria a tabela fato final
def merge(df_vazoes_base, df_postos_prod):
    df_vazoes = df_vazoes_base.merge(df_postos_prod, how='left')

    return df_vazoes


# funcao que faz os cálculos da ena
def calcula_ena(df_vazoes, dict_month):
    df_ena = df_vazoes.copy()
    for month_name in dict_month.keys():
        df_ena.loc[:, month_name] = df_ena[month_name] * df_ena['prod']

    # Deleta as linhas que sao dos anos de A-1/A+0 e reseta o indice do dataframe
    ano = date.today().year
    indexNames = df_ena[(df_ena['ano'] == ano - 1) | (df_ena['ano'] == ano)].index
    df_ena.drop(indexNames, inplace=True)
    df_ena.reset_index(inplace=True, drop=True)

    # filto pela bacia, e pelo ano (olha os dois juntos),e ai ele soma pelo posto, pois cada posto esta realacionado a uma bacia so
    df_ena_bacias = df_ena.groupby(['bacia', 'ano']).sum().reset_index()

    return df_ena, df_ena_bacias


# funcao que tem o intuito de calcular as variaveis de posição para cada bacia e para cada mes especifico
def final_stat(df_ena_bacias):
    df_final_stat = pd.DataFrame()
    for bacia in df_ena_bacias.bacia.unique():
        df_aux = round(df_ena_bacias.loc[df_ena_bacias.bacia == bacia].describe(), 0).drop(index=['count', '50%'])
        df_aux.reset_index(inplace=True)
        df_aux['bacia'] = bacia
        df_final_stat = pd.concat([df_final_stat, df_aux], ignore_index=True)
    df_final_stat.drop(columns=['prod', 'posto', 'ano'], inplace=True)
    df_final_stat = reorder_columns(df_final_stat)
    return df_final_stat


# Função para alocar a última coluna do DF como primeira
def reorder_columns(df):
    cols = df.columns.to_list()
    cols = cols[-1:] + cols[:-1]
    df = df[cols]

    return df


# Função que define o caminho do Excel que serve como saida para os dados gerados por este código
def connect_to_excel():
    app = xw.App()
    wb = app.books.open(r'H:\Comercializadora\Preços\Leonardo\Projetos_Python\Projeto_MLT_Bacias\MLT_por_bacia.xlsx')
    ws_estatisticas = wb.sheets('Estatisticas')

    return app, wb, ws_estatisticas


# codigo termina aqui
if __name__ == '__main__':
    main()

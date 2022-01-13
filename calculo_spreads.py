# -*- coding: utf-8 -*-

#Importações de módulos
import pandas as pd
from datetime import datetime

#Diretório raiz
#diretorio = input('Insira o diretório onde está a planilha fonte:') + '/'
diretorio = 'D:/Work/Estudos títulos/Resultados/'

#Lista de funções
def ajeita_data():
    '''Função que pega a data de hoje e transforma no formato dd/mm/YY
    Parâmetro de entrada:
    Valor de retorno: str'''
    hj = datetime.today()
    ano = str(hj.year)
    mes = '{:02d}'.format(hj.month)
    dia = '{:02d}'.format(hj.day)
    data_ajustada = dia + '/' + mes + '/' + ano
    return data_ajustada

def converter_em_lista(string):
    '''Funcao para converter string em lista
    Parametro de entrada: string
    Valor de retorno: list'''
    string = str(string)
    li = list(string.split(" "))
    return li 

def transforma_data(dados, nomecoluna, formato = '%Y-%m-%d'):
    '''Funcao que formata uma coluna de um dataframe com formato data
    Parâmetro de entrada: DataFrame, str, date
    Valor de retorno: DataFrame'''
    dados[nomecoluna] = pd.to_datetime(dados[nomecoluna], format = formato).dt.date
    return dados

def dados_serie_sgs(codigo_series, data_inicial = '01/01/2017', data_final = ajeita_data()):
    '''Funcao que pega o código de n séries e coleta seus valores entre as datas definidas
    Parâmetro de entrada: int, str, str
    Valor de retorno: pandas'''
    codigo_series = converter_em_lista(codigo_series)
    for i in range(len(list(codigo_series))):
        url_sgs = ("http://api.bcb.gov.br/dados/serie/bcdata.sgs." + str(codigo_series[i]) + "/dados?formato=csv&dataInicial=" + data_inicial + "&dataFinal=" + data_final)
        dados_um_codigo = pd.read_csv(url_sgs, sep=';', dtype = 'str')
        dados_um_codigo['valor'] = dados_um_codigo['valor'].str.replace(',', '.')
        dados_um_codigo['valor'] = dados_um_codigo['valor'].astype(float)
        dados_um_codigo = pd.DataFrame(dados_um_codigo)
        dados_um_codigo = dados_um_codigo.rename(columns = {'Index': 'data', 'valor': str(codigo_series[i])})
        if i==0:
            dados_merge = dados_um_codigo
        else:
            dados_merge = dados_merge.merge(dados_um_codigo, how='outer',on='data')
    return dados_merge


def importar_xlsx(nome_arquivo):
    '''Funcao que importa o arquivo xlsx e transforma em dataframe
    Parâmetro de entrada: excel
    Valor de retorno: pandas'''
    arquivo = pd.read_excel(diretorio + nome_arquivo, header = None)
    arquivo = arquivo.drop(arquivo.index[[0,1,2,3,4,5,6,8]])
    arquivo = arquivo.rename(columns = arquivo.iloc[0,:])
    arquivo = arquivo.drop(arquivo.index[[0]])
    return arquivo

def organiza(dados, nomecoluna):
    '''Função que organiza linhas de acordo com ordem de uma coluna
    Parâmetro de entrada: pandas, str
    Valor de retorno: DataFrame'''
    dados = dados.sort_values(by = [nomecoluna])
    return dados

def filt_dados_igual(dados, nome_coluna, condicao):
    '''Funcao que filtra os dados de acordo com igualdade com determinada condição
    Parâmetro de entrada: DataFrame, str, str
    Valor de retorno: DataFrame'''
    is_condicao = dados[nome_coluna] == condicao
    dados_filtrados = dados[is_condicao]
    return dados_filtrados

def filt_dados_dif(dados, nome_coluna, condicao):
    '''Funcao que filtra os dados de acordo com diferença com determinada condição
    Parâmetro de entrada: DataFrame, str, str
    Valor de retorno: DataFrame'''
    is_condicao = dados[nome_coluna] != condicao
    dados_filtrados = dados[is_condicao]
    return dados_filtrados

def transforma_numero(dados, nomecoluna):
    '''Funcao que transforma strings em dados numéricos
    Parâmetro de entrada: DataFrame, str
    Valor de retorno: DataFrame'''
    dados[nomecoluna] = dados[nomecoluna].str.replace('.', '')
    dados[nomecoluna] = dados[nomecoluna].str.replace(',', '.')
    dados[nomecoluna] = pd.to_numeric(dados[nomecoluna])
    return dados[nomecoluna]

#Importação de série e tratamento com meta selic diária
meta_selic = dados_serie_sgs('432')
meta_selic = transforma_data(meta_selic, 'data', '%d/%m/%Y')
meta_selic = meta_selic.rename(columns = {'data':'Data do leilão', '432':'Meta SELIC'})


###IMPORTAÇÃO DO ARQUIVO DA LTN###
#Importação do arquivo excel vindo do site do tesouro nacional
arquivo_importado = input('Insira nome do arquivo excel da LTN a ser importado:') + '.xlsx'
titulo = importar_xlsx(arquivo_importado)

#Formatando datas
titulo = transforma_data(titulo, 'Data do leilão', formato = '%d/%m/%Y')
titulo = transforma_data(titulo, 'Data de liquidação', formato = '%d/%m/%Y')
titulo = transforma_data(titulo, 'Data de vencimento', formato = '%d/%m/%Y')

#Formatando como número
#titulo['Oferta'] = transforma_numero(titulo, 'Oferta')
#titulo['Taxa média'] = transforma_numero(titulo, 'Taxa média')
#titulo['Taxa de corte'] = transforma_numero(titulo, 'Taxa de corte')
#titulo['Venda'] = transforma_numero(titulo, 'Venda')
#titulo['Financeiro (R$)'] = transforma_numero(titulo, 'Financeiro (R$)')
#titulo['Venda para Bacen'] = transforma_numero(titulo, 'Venda para Bacen')
#titulo['Financeiro para Bacen (R$)'] = transforma_numero(titulo, 'Financeiro para Bacen (R$)')

#Organizando dados de acordo com data do leilão
titulo = organiza(titulo, 'Data do leilão')

#Filtros
#1)Tipo de leilão
titulo = filt_dados_igual(titulo, 'Tipo de leilão', 'Venda')
#2)Volta
titulo = filt_dados_igual(titulo, 'Volta', '1.ª volta')
#3)Venda
titulo = filt_dados_dif(titulo, 'Venda', 0)
    
#3)Data de vencimento e inserção da meta SELIC

datas_disponiveis = pd.unique(titulo['Data de vencimento']) #usado para filtrar as diferentes datas de vencimento
datas_disponiveis.sort() #Ordenar cronologicamente as datas

nome_datas_disponiveis = pd.DataFrame(pd.unique(titulo['Data de vencimento'])) #usado pra nomear as abas da planilha
nome_datas_disponiveis = organiza(nome_datas_disponiveis, 0)

tipo_titulo = 'LTN'
    
for i in range(len(datas_disponiveis)):
    data_vencimento = datas_disponiveis[i]
    aba_titulo = filt_dados_igual(titulo, 'Data de vencimento', data_vencimento)
    aba_titulo = pd.merge(meta_selic, aba_titulo, on = 'Data do leilão')
    aba_titulo['Spread'] = aba_titulo['Taxa média'] - aba_titulo['Meta SELIC']
    data_pro_nome = str(nome_datas_disponiveis.iloc[[i]])[-10:]
    data_pro_nome = data_pro_nome.replace('-','_')
    nome_titulo = tipo_titulo + '_' + str(data_pro_nome)
    nome_aba = nome_titulo[:11]
    exec(nome_titulo + " = aba_titulo")
    writer = diretorio + 'Titulos_spreads_taxas.xlsx'
    
    if i == 0:
        aba_titulo.to_excel(writer, index = False, sheet_name = nome_aba)
    else:
        with pd.ExcelWriter(writer, engine="openpyxl", mode = 'a') as writer:
            aba_titulo.to_excel(writer, index = False, sheet_name = nome_aba)        
                    
    spread = aba_titulo[['Data do leilão', 'Spread']]
    spread = spread.rename(columns = {'Spread': nome_aba})
    if i == 0:
        spread_merge = spread
    else:
        spread_merge = spread_merge.merge(spread, how='outer',on='Data do leilão')
        
    tx_media = aba_titulo[['Data do leilão', 'Taxa média']]
    tx_media = tx_media.rename(columns = {'Taxa média': nome_aba})
    if i == 0:
        tx_media_merge = tx_media
    else:
        tx_media_merge = tx_media_merge.merge(tx_media, how='outer',on='Data do leilão')


###IMPORTAÇÃO DO ARQUIVO DA NTN-F###
#Importação do arquivo excel vindo do site do tesouro nacional
arquivo_importado2 = input('Insira nome do arquivo excel da NTN-F a ser importado:') + '.xlsx'
titulo2 = importar_xlsx(arquivo_importado2)

#Formatando datas
titulo2 = transforma_data(titulo2, 'Data do leilão', formato = '%d/%m/%Y')
titulo2 = transforma_data(titulo2, 'Data de liquidação', formato = '%d/%m/%Y')
titulo2 = transforma_data(titulo2, 'Data de vencimento', formato = '%d/%m/%Y')

#Formatando como número
#titulo2['Oferta'] = transforma_numero(titulo2, 'Oferta')
#titulo2['Taxa média'] = transforma_numero(titulo2, 'Taxa média')
#titulo2['Taxa de corte'] = transforma_numero(titulo2, 'Taxa de corte')
#titulo2['Venda'] = transforma_numero(titulo2, 'Venda')
#titulo2['Financeiro (R$)'] = transforma_numero(titulo2, 'Financeiro (R$)')
#titulo2['Venda para Bacen'] = transforma_numero(titulo2, 'Venda para Bacen')
#titulo2['Financeiro para Bacen (R$)'] = transforma_numero(titulo2, 'Financeiro para Bacen (R$)')

#Organizando dados de acordo com data do leilão
titulo2 = organiza(titulo2, 'Data do leilão')

#Filtros
#1)Tipo de leilão
titulo2 = filt_dados_igual(titulo2, 'Tipo de leilão', 'Venda')
#2)Volta
titulo2 = filt_dados_igual(titulo2, 'Volta', '1.ª volta')
#3)Venda
titulo2 = filt_dados_dif(titulo2, 'Venda', 0)
    
#3)Data de vencimento e inserção da meta SELIC

datas_disponiveis2 = pd.unique(titulo2['Data de vencimento']) #usado para filtrar as diferentes datas de vencimento
datas_disponiveis2.sort() #Ordenar cronologicamente as datas

nome_datas_disponiveis2 = pd.DataFrame(pd.unique(titulo2['Data de vencimento'])) #usado pra nomear as abas da planilha
nome_datas_disponiveis2 = organiza(nome_datas_disponiveis2, 0)

tipo_titulo2 = 'NTN_F'
    
for i in range(len(datas_disponiveis2)):
    data_vencimento2 = datas_disponiveis2[i]
    aba_titulo2 = filt_dados_igual(titulo2, 'Data de vencimento', data_vencimento2)
    aba_titulo2 = pd.merge(meta_selic, aba_titulo2, on = 'Data do leilão')
    aba_titulo2['Spread'] = aba_titulo2['Taxa média'] - aba_titulo2['Meta SELIC']
    data_pro_nome2 = str(nome_datas_disponiveis2.iloc[[i]])[-10:]
    data_pro_nome2 = data_pro_nome2.replace('-','_')
    nome_titulo2 = tipo_titulo2 + '_' + str(data_pro_nome2)
    nome_aba2 = nome_titulo2[:13]
    exec(nome_titulo2 + " = aba_titulo2")
    with pd.ExcelWriter(writer, engine="openpyxl", mode = 'a') as writer:
        aba_titulo2.to_excel(writer, index = False, sheet_name = nome_aba2)
                    
    spread2 = aba_titulo2[['Data do leilão', 'Spread']]
    spread2 = spread2.rename(columns = {'Spread': nome_aba2})
    if i == 0:
        spread_merge2 = spread2
    else:
        spread_merge2 = spread_merge2.merge(spread2, how='outer',on='Data do leilão')
        
    tx_media2 = aba_titulo2[['Data do leilão', 'Taxa média']]
    tx_media2 = tx_media2.rename(columns = {'Taxa média': nome_aba2})
    if i == 0:
        tx_media_merge2 = tx_media2
    else:
        tx_media_merge2 = tx_media_merge2.merge(tx_media2, how='outer',on='Data do leilão')


###Juntando resumos e salvando no arquivo###
spread_final = pd.merge(spread_merge, spread_merge2, how='outer',on='Data do leilão')
spread_final = spread_final.sort_values(by = 'Data do leilão') #Ordenando cronologicamente datas

tx_media_final = pd.merge(tx_media_merge, tx_media_merge2, how='outer',on='Data do leilão')
tx_media_final = tx_media_final.sort_values(by = 'Data do leilão') #Ordenando cronologicamente datas


with pd.ExcelWriter(writer, engine="openpyxl", mode = 'a') as writer:  
    spread_final.to_excel(writer, index = False, sheet_name = 'Spreads')
    tx_media_final.to_excel(writer, index = False, sheet_name = 'Taxa média')





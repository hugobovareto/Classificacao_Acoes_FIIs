'''
Variáveis importantes para se ter para classificação e tomada de decisão de cada AÇÃO:
- código;
- nome;
- segmento de listagem;
- setor;
- segmento;
- tag along;
- P/VP;
- P/L;
- DY;
- margem líquida;
- ROE;
- Dívida bruta / Patrimônio líquido;
- Dívida líquida / EBTIDA;
- Dividendos por ação no ano (possível conseguir a informação pelo Dividend Yield - DY);
- Preço-teto Graham (precisa de LPA e VPA);
- Preço-teto Bazin (precisa do Dividendos por ação no ano).


Classificação a partir de histórico de indicadores (decrescente; decrescente/constante; constante; constante/ crescente; crescente):
Histórico de dividendos;
Histórico de margem líquida;
Histórico de ROE;
HIstórico Dívida líquida / Patrimônio líquido;


A função 'get_resultado' da biblioteca 'fundamentus' não tem a quantidade de ações e lucro líquido, o que seria necessário para calcular LPA e VPA para ter o preço-teto Graham.
Mas posso usar a lista de ações e trazer as informações por 'fundamentus.get_papel([lista de ações]) que tem todos os indicadores que preciso.
(https://pypi.org/project/fundamentus/)

Para FIIs que tenho que usar outra estratégia.
Os indicadores são diferentes e a biblioteca 'fundamentus' não traz esses dados.

Variáveis importantes para se ter para classificação e tomada de decisão de cada FIIs:
- código;
- nome;
- tipo (papel; tijolo; misto);
- segmento;
- P/VP;
- DY;
- Valor Patrimonial por cota;
- Dividendos por cota (no ano);
- Preço-teto para FIIs.


Para FIIs a biblioteca 'brick-by-brick' parece funcionar.
(https://github.com/brunoruas2/brick-by-brick)

Classificação a partir de histórico de indicadores (decrescente; decrescente/constante; constante; constante/ crescente; crescente):
Histórico de dividendos;

'''
# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re
import streamlit as st
import time
import plotly.express as px
import plotly.graph_objects as go
import fundamentus
import investiny
import yfinance as yf


################################ AÇÕES ####################################

# Importação dos dados de ações
df_acoes = fundamentus.get_resultado_raw()

# Lista de Ações
lista_acoes = df_acoes.index.tolist()

# Transformar o índice em uma coluna:
df_acoes = df_acoes.reset_index()


# Percorrer todos os papeis e usar a função get_papel para trazer os dados detalhados de cada ação
dados_detalhados = []

for papel in tqdm(lista_acoes):
    dados = fundamentus.get_papel(papel)
    dados_detalhados.append(dados)


# Transformar os dados detalhados em DataFrame
df_detalhado_acoes = pd.concat(dados_detalhados, ignore_index=True)


# Valores de alguns indicadores estão vindo sem decimal (Ex.: 1153, na verdade deveria ser 11,52)
colunas_dividir_100 = [
    'PL',
    'PVP',
    'PEBIT',
    'PSR',
    'PAtivos',
    'PCap_Giro',
    'PAtiv_Circ_Liq',
    'EV_EBITDA',
    'EV_EBIT',
    'LPA',
    'VPA',
    'Liquidez_Corr',
    'Div_Liq_Patrim',
    'Giro_Ativos'
]

# Converter as colunas para numérico
df_detalhado_acoes[colunas_dividir_100] = (
    df_detalhado_acoes[colunas_dividir_100]
    .apply(pd.to_numeric, errors='coerce')
)

# Dividir por 100
df_detalhado_acoes[colunas_dividir_100] = (
    df_detalhado_acoes[colunas_dividir_100] / 100
).round(2)

# Retirar "%" das colunas em percentual para evitar problemas
# Colunas que possuem valores em percentual
colunas_percentuais = [
    'Div_Yield',
    'Cres_Rec_5a',
    'Marg_Bruta',
    'Marg_EBIT',
    'Marg_Liquida',
    'EBIT_Ativo',
    'ROIC',
    'ROE'
]

# Remover o símbolo % e converter para número
for coluna in colunas_percentuais:
    df_detalhado_acoes[coluna] = (
        df_detalhado_acoes[coluna]
        .astype(str)
        .str.replace('%', '', regex=False)
    )


# Converter todas as colunas quantitativas para numérico para evitar erros e problemas
colunas_texto = [
    'Papel',
    'Tipo',
    'Empresa',
    'Setor',
    'Subsetor',
    'Data_ult_cot',
    'Ult_balanco_processado'
]

# Todas as demais são quantitativas
colunas_numericas = [
    col for col in df_detalhado_acoes.columns
    if col not in colunas_texto
]

# Converter para numérico
df_detalhado_acoes[colunas_numericas] = (
    df_detalhado_acoes[colunas_numericas]
    .apply(pd.to_numeric, errors='coerce')
)


# Criar variáveis necessárias a partir das variáveis já existentes no df_detalhado_acoes com o get_papel()
def divisao_segura(numerador, denominador):
    numerador = pd.to_numeric(numerador, errors='coerce')
    denominador = pd.to_numeric(denominador, errors='coerce')

    if hasattr(denominador, 'where'):
        denominador = denominador.where(denominador != 0)
    else:
        denominador = None if denominador == 0 else denominador

    return numerador / denominador

# Não tem a variável EBTIDA, mas dá para calcular a partir de outra

# Obter EBTIDA (a partir de 'EV_EBITDA')
# Valor_da_firma = EV, logo EV_EBITDA = Valor_da_firma / EBITDA, então EBITDA = Valor_da_firma / EV_EBITDA
df_detalhado_acoes['EBITDA'] = divisao_segura(df_detalhado_acoes['Valor_da_firma'], df_detalhado_acoes['EV_EBITDA'])

# Dívida bruta / Patrimônio líquido;
df_detalhado_acoes['Div Bruta/ Patrim Liq'] = divisao_segura(df_detalhado_acoes['Div_Bruta'], df_detalhado_acoes['Patrim_Liq'])

# Dívida líquida / EBTIDA;
df_detalhado_acoes['Div Líquida/ EBTIDA'] = divisao_segura(df_detalhado_acoes['Div_Liquida'], df_detalhado_acoes['EBITDA'])

# Dividendos/ação (ano)
df_detalhado_acoes['Dividendo_12_meses'] = (df_detalhado_acoes['Div_Yield'] * df_detalhado_acoes['Cotacao'] / 100).round(2)

# Preço-teto Graham
df_detalhado_acoes['Preco_teto_Graham'] = np.sqrt(df_detalhado_acoes['LPA'] * df_detalhado_acoes['VPA'] * 22.5) 

# Preço-teto Bazin
df_detalhado_acoes['Preco_teto_Bazin'] = df_detalhado_acoes['Dividendo_12_meses'] / 0.06


# Mudar nome de algumas colunas
df_detalhado_acoes.rename(columns={'PL': 'P/L',
                                   'PVP': 'P/VP',
                                   'PEBIT': 'P/EBIT',
                                   'PSR': 'P/SR',
                                   'PAtivos': 'P/Ativos',
                                   'PCap_Giro': 'P/Cap_Giro',
                                   'PAtiv_Circ_Liq': 'P/Ativ_Circ_Liq',
                                   'EV_EBITDA': 'EV/EBITDA',
                                   'EV_EBIT': 'EV/EBIT',
                                   'Div_Liq_Patrim': 'Div Liq/ PL'}, inplace=True)


# Reorganizar a ordem das colunas para ter as variáveis que tenho interesse no início do DataFrame
colunas_prioritarias = [
    'Papel',
    'Empresa',
    'Setor',
    'Subsetor',
    'P/VP',
    'P/L',
    'Div_Yield',
    'Marg_Liquida',
    'ROE',
    'Div Bruta/ Patrim Liq',
    'Div Líquida/ EBTIDA',
    'Dividendo_12_meses',
    'Preco_teto_Graham',
    'Preco_teto_Bazin'
]

outras_colunas = [
    coluna for coluna in df_detalhado_acoes.columns
    if coluna not in colunas_prioritarias
]

df_acoes_finais = df_detalhado_acoes[
    colunas_prioritarias + outras_colunas
]


# Exportar em Excel o df_acoes_finais
df_acoes_finais.to_excel(
    'Acoes.xlsx',
    index=False
)


################################ FIIs #####################################













from http import HTTPStatus

import pandas as pd
from fastapi import FastAPI
from fastapi.openapi.utils import get_openapi

from backend.schemas import ListaOFs, Quantidade

tags_metadata = [{'name': 'Listas'}, {'name': 'Quantidades'}, {'name': 'Estado OF'}]

app = FastAPI(openapi_tags=tags_metadata)

chaves = [
    'Chave-C',
    'Nome',
    'E-mail',
    'Matricula',
    'Service Line',
    'W# Forecast',
    'Acionamento OF',
    'OF',
    'OF Grupo/Workitem',
    'Chave-C vinculada',
    'Perfil OF',
    'Estado OF',
    'GECAP',
    'Form',
    'Forecast USTIBB',
    'Forecast R$',
    'USTIBBs OF ATUAL',
    'DELTA USTIBBs',
    'DELTA R$ Forecast',
    'DPE',
    'Gerente Equip. - BB',
    'RT - BB',
    'Férias/Afastamentos/Observações',
    'On-board',
]

estados_of = {
    'cancelada': 'Cancelada',
    'cancelada-fiscal': 'Cancelada por fiscal de contrato',
    'perfil-incluido': 'Perfil Incluido',
    'perfil-iniciada': 'Entrega Perfil Iniciada',
    'perfil-efetivada': 'Entrega Perfil Efetivada',
    'perfil-validada': 'Entrega do perfil validada',
    'validada-gerente': 'Validada pelo Gerente',
    'rof-gerado': 'Rof Gerado',
}


# Endpoint de listagem de todas as OFs
@app.get(
    '/ofs/lista-tudo',
    status_code=HTTPStatus.OK,
    summary='Lista todas as OFs do mês',
    description='teste',
    response_model=ListaOFs,
    tags=['Listas'],
)
def lista_tudo():
    lista_linhas = []
    linha = {}

    # Lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['OF'], int) or isinstance(row['OF'], float) and not isinstance(row['OF'], str):
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint de listagem de todas as OFs vinculadas com uma Chave C
@app.get(
    '/ofs/lista-vinculadas',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs vinculadas com uma Chave C',
    description='teste',
    response_model=ListaOFs,
    tags=['Listas'],
)
def lista_ofs_vinculadas():
    lista_linhas = []
    linha = {}

    # Lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'] == 'Sim' or row['Chave-C vinculada'] == 'sim':
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint de listagem de todas as OFs não vinculadas com uma Chave C
@app.get(
    '/ofs/lista-nao-vinculadas',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs não vinculadas com uma Chave C',
    description='teste',
    response_model=ListaOFs,
    tags=['Listas'],
)
def lista_ofs_nao_vinculadas():
    lista_linhas = []
    linha = {}

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'] == 'Não' or row['Chave-C vinculada'] == 'não':
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a quantidade total de OFs
@app.get(
    '/ofs/quantidade-total',
    status_code=HTTPStatus.OK,
    summary='Retorna a quantidade total de OFs no mês',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades'],
)
def quantidade_ofs_total():
    counter = 0
    resp = {}

    # Lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['OF'], float):
            counter += 1

    resp = {'quantidade': counter}
    return resp


# Endpoint que retorna a quantidade de OFs vinculadas
@app.get(
    '/ofs/quantidade-vinculadas',
    status_code=HTTPStatus.OK,
    summary='Retorna a quantidade total de OFs vinculadas com uma Chave-C',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades'],
)
def quantidade_ofs_vinculadas():
    counter = 0
    resp = {}

    # Lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'].upper() == 'SIM':
            counter += 1

    resp = {'quantidade': counter}
    return resp


# Endpoint que retorna a quantidade de OFs não vinculadas
@app.get(
    '/ofs/quantidade-nao-vinculadas',
    status_code=HTTPStatus.OK,
    summary='Retorna a quantidade total de OFs não vinculadas com uma Chave-C',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades'],
)
def quantidade_ofs_nao_vinculadas():
    counter = 0
    resp = {}

    # Lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'].UPPER() == 'NÃO':
            counter += 1

    resp = {'quantidade': counter}
    return resp


# Endpoint que retorna a lista de OFs com estado de (Perfil Incluido)
@app.get(
    '/ofs/estado/perfil-incluido',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Perfil Incluido)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_perfil_incluido():
    lista_linhas = []
    linha = {}

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['perfil-incluido'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a lista de OFs com estado de (Entrega de Perfil Iniciado)
@app.get(
    '/ofs/estado/entrega-perfil-iniciada',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Entrega de Perfil Iniciado)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_perfil_iniciada():
    lista_linhas = []
    linha = {}

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['perfil-iniciada'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a lista de OFs com estado de (Entrega de Perfil Efetivada)
@app.get(
    '/ofs/estado/entrega-perfil-efetivada',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Entrega de Perfil Efetivada)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_perfil_efetivada():
    lista_linhas = []
    linha = {}

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['perfil-efetivada'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# OpenAPI configs
def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema
    openapi_schema = get_openapi(
        title='Backend Desafio - Level Up',
        version='1.0.0',
        summary='API em python para prover dados, utilizando o FastAPI',
        description='Utiliza como base de dados as planilhas excel/csv utilizadas no fluxo de trabalho de acompanhamento das OFs - Documentação via SwaggerUI e Redoc',
        routes=app.routes,
    )
    openapi_schema['info']['x-logo'] = {'url': 'https://www.ibm.com/brand/experience-guides/developer/b1db1ae501d522a1a4b49613fe07c9f1/01_8-bar-positive.svg'}
    app.openapi_schema = openapi_schema
    return app.openapi_schema


app.openapi = custom_openapi

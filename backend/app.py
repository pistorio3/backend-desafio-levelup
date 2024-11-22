from http import HTTPStatus

import pandas as pd
from fastapi import FastAPI
from fastapi.openapi.utils import get_openapi

from backend.schemas import ListaOFs, Quantidade

tags_metadata = [{'name': 'Listas'}, {'name': 'Quantidades'}, {'name': 'Estado OF'}, {'name': 'Quantidades Estados OFs'}]

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
    'perfil-validada': 'Entrega Perfil Validada',
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
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'].upper() == 'SIM':
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
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'].upper() == 'NÃO':
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
        if isinstance(row['Chave-C vinculada'], str) and row['Chave-C vinculada'].upper() == 'NÃO':
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


# Endpoint que retorna a lista de OFs com estado de (Entrega de Perfil Validada)
@app.get(
    '/ofs/estado/entrega-perfil-validada',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Entrega de Perfil Validada)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_perfil_validada():
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
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['perfil-validada'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a lista de OFs com estado de (Validada pelo gerente)
@app.get(
    '/ofs/estado/validada-gerente',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Validada pelo gerente)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_validada_gerente():
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
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['validada-gerente'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta

# Endpoint que retorna a lista de OFs com estado de (Cancelada)
@app.get(
    '/ofs/estado/cancelada',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Cancelada)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_validada_gerente():
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
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['cancelada'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta

# Endpoint que retorna a lista de OFs com estado de (Cancelada por fiscal de contrato)
@app.get(
    '/ofs/estado/cancelada-fiscal',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Cancelada por fiscal de contrato)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_cancelada_fiscal():
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
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['cancelada-fiscal'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a lista de OFs com estado de (Rof Gerado)
@app.get(
    '/ofs/estado/rof-gerado',
    status_code=HTTPStatus.OK,
    summary='Lista de OFs com estado de (Rof Gerado)',
    description='teste',
    response_model=ListaOFs,
    tags=['Estado OF'],
)
def lista_ofs_rof_gerado():
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
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['rof-gerado'].upper():
            for i in range(len(chaves)):
                linha[chaves[i]] = row[chaves[i]]
            lista_linhas.append(linha)

    quantidade = len(lista_linhas)
    resposta = {'resposta': lista_linhas, 'quantidade': quantidade}
    return resposta


# Endpoint que retorna a quantidade de OFs com estado de (Perfil Incluido)
@app.get(
    '/ofs/estado/quantidade/perfil-incluido',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Perfil Incluido)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

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
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Entrega Perfil Iniciada)
@app.get(
    '/ofs/estado/quantidade/entrega-perfil-iniciada',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Entrega Perfil Iniciada)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

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
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Entrega Perfil Efetivada)
@app.get(
    '/ofs/estado/quantidade/entrega-perfil-efetivada',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Entrega Perfil Efetivada)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

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
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Entrega Perfil Validada)
@app.get(
    '/ofs/estado/quantidade/entrega-perfil-validada',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Entrega Perfil Validada)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['perfil-validada'].upper():
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Validada pelo Gerente)
@app.get(
    '/ofs/estado/quantidade/validada-gerente',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Validada pelo Gerente)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['validada-gerente'].upper():
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Cancelada)
@app.get(
    '/ofs/estado/quantidade/cancelada',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Cancelada)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['cancelada'].upper():
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Cancelada por fiscal de contrato)
@app.get(
    '/ofs/estado/quantidade/cancelada-fiscal-contrato',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Cancelada por fiscal de contrato)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['cancelada-fiscal'].upper():
            counter += 1

    resp = {'quantidade': counter}
    return resp

# Endpoint que retorna a quantidade de OFs com estado de (Rof Gerado)
@app.get(
    '/ofs/estado/quantidade/rof-gerado',
    status_code=HTTPStatus.OK,
    summary='Quantidade de OFs com estado de (Rof Gerado)',
    description='teste',
    response_model=Quantidade,
    tags=['Quantidades Estados OFs'],
)
def quantidade_ofs_perfil_incluido():
    counter = 0

    # lendo planilha
    df = pd.read_excel(
        'backend/sheets/2024 - Controle de OF - Amostra.xlsm',
        sheet_name='2024-10',
    )

    # Limpando Dataframe de linhas vazias
    df = df.dropna(thresh=15)
    df = df.reset_index()

    for index, row in df.iterrows():
        if isinstance(row['Estado OF'], str) and row['Estado OF'].upper() == estados_of['rof-gerado'].upper():
            counter += 1

    resp = {'quantidade': counter}
    return resp

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

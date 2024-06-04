from pandas import DataFrame
import pandas as pd
import openpyxl
import requests
import tabula
import os

pdf_path = ""

def init_convertion(file, page_number):
    tables = tabula.read_pdf(file, pages='all')
    pages = int(page_number) - 1
    
    retorno = format_cols(tables, pages)
    return retorno

def format_cols(tables, page_number):
    columns_name = ['REVENDEDORA', 'N.F', 'DESTINO']
    numero_pagina = int(page_number)
    
    # Converta a tabela especificada para um DataFrame
    df = pd.DataFrame(tables[numero_pagina])

    # Limpe os dados
    df.dropna()
    
    # # Trate a coluna "NF do pedido"
    col_nf_pedido = df.iloc[:, 0].astype(str)
    col_nf_pedido = col_nf_pedido.str.strip().str.replace("O", "").str.replace("0 ", "").str.replace(" ","").str.replace(",", "").str.replace(".0", "").str.replace('nan', 'vazio')
    df.iloc[:, 0] = col_nf_pedido
    
    # # Trate a coluna "CEP"
    col_cep = df.iloc[:, 1].astype(str)
    col_cep = col_cep.str.strip().str.replace(" ", "").str.replace(",", "").str.replace(".0", "")
    df.iloc[:, 1] = col_cep

    # Concatenando as colunas 3 e 4 por conta do nome da consultora
    terceira_coluna = df.iloc[:, 2].astype(str)
    quarta_coluna = df.iloc[:, 3].astype(str)

    df['REVENDEDORA'] = terceira_coluna + " " + quarta_coluna

    # Tratando a coluna Revendedora
    df['REVENDEDORA'] = df["REVENDEDORA"].str.strip().str.replace('\d+', '', regex=True)
    
    # Inserindo a coluna com o nome da cidade e estado
    df['DESTINO'] = df.iloc[:, 1].apply(buscar_cidade_uf)
    
    # Apagando as linhas desnecessárias
    row_values = ['NF do pedido', 'NFdopedido']
    df = df[df.iloc[:, 0].isin(row_values) == False]
    
    # Buscando a quantidade de colunas na planilha para tratamento posterior
    num_cols = len(df.columns)
    print(f'Número de colunas: {num_cols}')
    
    remove_columns(df, num_cols)
    
    # Excluindo linhas que possuem o valor nan nan
    df.drop(df[df["REVENDEDORA"] == 'nan nan'].index, inplace=True)    
    
    # Reordenando as colunas restantes para manter o padrão usado
    df = df.iloc[:, [1,0,2]]
    
    # Alterando o nome das colunas para melhor visualização
    df.columns = columns_name
    
    # Chama a função para salvar o arquivo
    save_on_xlsx(df)
    
    retorno = repair_sheet()
    return retorno

def buscar_cidade_uf(cep):
    if len(cep) == 8:
        link = f'https://viacep.com.br/ws/{cep}/json/'
        
        try:
            requisicao = requests.get(link)
            requisicao.raise_for_status()  # Lança exceção se a requisição falhar
            
            dic_requisicao = requisicao.json()
            
            uf = dic_requisicao['uf']
            cidade = dic_requisicao['localidade']
            
            return f'{cidade} - {uf}'
        except requests.exceptions.HTTPError as http_err:
            print(f'Erro HTTP: {http_err}')
        except Exception as err:
            print(f'Erro: {err}')
        else:
            return None
    else:
        print("CEP Inválido")
        return "nan"

def save_on_xlsx(df):
    xlsx_path = "data.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(df.columns.tolist())

    for i, row in df.iterrows():
        ws.append(row.tolist())
    try:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)
            
        wb.save(xlsx_path)
        return f'Arquivo salvo com sucesso: {xlsx_path}'
    except Exception as e:
        return f"Não foi possível salvar o arquivo: {e}"

def remove_columns(df: DataFrame, num_cols: int) -> None:
    values = []
    for i in range(1, num_cols - 2):
        values.append(i)
    df.drop(df.columns[values], axis=1, inplace=True)
    
    
# Retrabalhando a planilha para que os nomes das consultoras se mantenham em apenas 1 coluna
def repair_sheet():
    # Carrega um arquivo Excel para um DataFrame
    df = pd.read_excel('data.xlsx')

    # Inicializa listas vazias para armazenar informações
    nome_revendedora = []
    numero_nota = []
    destino_lista = []
    nome_revendedoras = []
    destinos_lista = []
    numero_notas = []

    # Imprime uma mensagem indicando o início do processamento
    print('NOME REVENDEDORA -- N.F')

    # Itera sobre cada linha do DataFrame
    for indice, linha in df.iterrows():

        # Extrai valores das colunas 'REVENDEDORA', 'N.F', e 'DESTINO'
        revendedora = linha['REVENDEDORA']
        nf = linha['N.F']
        destino = linha['DESTINO']
        
        # Imprime o nome da revendedora e o valor de 'N.F'
        print(f'{revendedora} -- {nf}')
        print()
        
        # Adiciona o nome da revendedora, o valor de 'N.F', e o destino às listas correspondentes
        nome_revendedora.append(revendedora)
        numero_nota.append(nf)
        destino_lista.append(destino)
        
        # Se 'N.F' for igual a 'vazio', imprime a lista de nomes de revendedoras,
        # junta os nomes em uma string separada por vírgulas, adiciona-as às listas de resultados,
        # e limpa as listas para o próximo ciclo
        if nf == 'vazio':
            print(nome_revendedora)
            nome_revendedoras.append(' '.join(nome_revendedora))  # Concatena os nomes das revendedoras
            destinos_lista.append(' '.join(map(str, destino_lista)))  # Junta os destinos em uma string
            numero_notas.append(' '.join(map(str, numero_nota)))  # Junta os números de nota em uma string
            nome_revendedora.clear()
            destino_lista.clear()
            numero_nota.clear()

    # Converte as listas de resultados em séries do pandas
    s_nome = pd.Series(nome_revendedoras)
    s_numero = pd.Series(numero_notas)
    s_destino = pd.Series(destinos_lista)

    # Adiciona as séries como novas colunas ao DataFrame
    df['REVENDEDORAS'] = s_nome
    df['N.Fs'] = s_numero
    df['DESTINO '] = s_destino

    # Pequenos ajustes nas colunas adicionadas
    df["REVENDEDORAS"] = df["REVENDEDORAS"].str.strip().str.replace(' nan', '')
    df["N.Fs"] = df["N.Fs"].str.strip().str.replace('vazio', '')
    df["DESTINO "] = df["DESTINO "].str.strip().str.replace(' nan', '')
    
    
    # Remove as primeiras três colunas do DataFrame
    # Esta linha parece estar incompleta ou mal aplicada, pois geralmente não é recomendável remover colunas importantes assim.
    # Talvez seja um erro ou uma decisão de remoção de colunas desnecessárias.
    df = df.drop(df.columns[[0,1,2]], axis=1)

    df.dropna(how='all', inplace=True)
    # Salva o DataFrame modificado em um novo arquivo Excel chamado 'test.xlsx',
    # sem incluir o índice do DataFrame no arquivo de saída
    retorno = save_on_xlsx(df)
    return retorno
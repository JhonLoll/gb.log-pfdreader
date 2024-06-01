from pandas import DataFrame
import pandas as pd
import openpyxl
import requests

def format_cols(tables):
    columns_name = ['REVENDEDORA', 'N.F', 'DESTINO']
    # Converta a tabela especificada para um DataFrame
    df = pd.DataFrame(tables[0])

    # Limpe os dados
    df.dropna()
    
    # # Trate a coluna "NF do pedido"
    col_nf_pedido = df.iloc[:, 0].astype(str)
    col_nf_pedido = col_nf_pedido.str.strip().str.replace("O", "").str.replace("0 ", "").str.replace(" ","").str.replace(",", "")
    df.iloc[:, 0] = col_nf_pedido
    
    # # Trate a coluna "CEP"
    col_cep = df.iloc[:, 1].astype(str)
    col_cep = col_cep.str.strip().str.replace(" ", "").str.replace(",", "")
    df.iloc[:, 1] = col_cep

    # Concatenando as colunas 3 e 4 por conta do nome da consultora
    terceira_coluna = df.iloc[:, 2].astype(str)
    quarta_coluna = df.iloc[:, 3].astype(str)

    df['REVENDEDORA'] = terceira_coluna + " " + quarta_coluna

    # Tratando a coluna Revendedora
    df['REVENDEDORA'] = df["REVENDEDORA"].str.strip()
    
    # Inserindo a coluna com o nome da cidade e estado
    df['DESTINO'] = df.iloc[:, 1].apply(buscar_cidade_uf)
    
    # Apagando as linhas desnecessárias
    row_values = ['NF do pedido', 'NFdopedido']
    df = df[df.iloc[:, 0].isin(row_values) == False]
    
    # Buscando a quantidade de colunas na planilha para tratamento posterior
    num_cols = len(df.columns)
    print(f'Número de colunas: {num_cols}')
    
    remove_columns(df, num_cols)
    
    
    df_clean = df.dropna(subset=['REVENDEDORA'])
    
    print(df_clean)
    
    df = df.iloc[:, [1,0,2]]
    df.columns = columns_name
    
    save_on_xlsx(df)

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
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(df.columns.tolist())

    for i, row in df.iterrows():
        ws.append(row.tolist())

    wb.save("data.xlsx")

def remove_columns(df: DataFrame, num_cols: int) -> None:
    values = []
    for i in range(1, num_cols - 2):
        values.append(i)
    df.drop(df.columns[values], axis=1, inplace=True)
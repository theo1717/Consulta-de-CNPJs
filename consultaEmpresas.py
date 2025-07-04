import pandas as pd
import requests
import time
from datetime import datetime
from sqlalchemy import create_engine
import os


PATH = "cnpjs.xlsx"
#DB_USUARIO = "root"
#DB_SENHA = ""
#DB_HOST = "localhost"
#DB_PORTA = 3306
#DB_NOME = "empresas"
#TABELA_SQL = "dados_cnpjs"
ARQUIVO_RESULTADO = "resultado_cnpjs.xlsx"

df_cnpjs = pd.read_excel(PATH)
df_cnpjs.columns = df_cnpjs.columns.map(str).str.strip().str.lower()

if 'cnpj' not in df_cnpjs.columns:
    for col in df_cnpjs.columns:
        if 'cnpj' in col.lower().strip():
            df_cnpjs = df_cnpjs.rename(columns={col: 'cnpj'})
            break
    else:
        raise KeyError("Coluna 'cnpj' não encontrada no arquivo.")

df_cnpjs = df_cnpjs[df_cnpjs['cnpj'].notna()]
print("CNPJs carregados:", len(df_cnpjs))


def formatar_data(data_str):
    try:
        return datetime.strptime(data_str, "%d/%m/%Y").strftime("%Y-%m-%d")
    except:
        return None
    

def formatar_cnpj(cnpj):
    cnpj = ''.join(filter(str.isdigit, str(cnpj)))
    if len(cnpj) == 14:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return cnpj


def consultar_cnpj(cnpj):
    cnpj = ''.join(filter(str.isdigit, str(cnpj)))
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    headers = {'User-Agent': 'Mozilla/5.0'}

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 429:
            print(f"[{cnpj}] Limite atingido. Aguardando 60s...")
            time.sleep(60)
            return consultar_cnpj(cnpj)
        elif response.status_code != 200:
            print(f"[{cnpj}] Erro {response.status_code} ao consultar.")
            return None

        dados = response.json()
        if dados.get("status") == "ERROR":
            print(f"[{cnpj}] Erro na API: {dados.get('message')}")
            return None

        return {
            "cnpj": dados.get("cnpj"),
            "razao_social": dados.get("nome"),
            "nome_fantasia": dados.get("fantasia"),
            "abertura": formatar_data(dados.get("abertura")),
            "situacao": dados.get("situacao"),
            "data_situacao": formatar_data(dados.get("data_situacao")),
            "motivo_situacao": dados.get("motivo_situacao"),
            "ultima_atualizacao": formatar_data(dados.get("ultima_atualizacao")),
            "tipo": dados.get("tipo"),
            "porte": dados.get("porte"),
            "natureza_juridica": dados.get("natureza_juridica"),
            "capital_social": dados.get("capital_social"),
            "telefone": dados.get("telefone"),
            "email": dados.get("email"),
            "logradouro": dados.get("logradouro"),
            "numero": dados.get("numero"),
            "complemento": dados.get("complemento"),
            "bairro": dados.get("bairro"),
            "municipio": dados.get("municipio"),
            "uf": dados.get("uf"),
            "cep": dados.get("cep"),
            "cnae_principal": dados.get("atividade_principal", [{}])[0].get("code"),
            "cnae_principal_desc": dados.get("atividade_principal", [{}])[0].get("text"),
            "cnaes_secundarios": "; ".join([
                f"{a.get('code')} - {a.get('text')}" for a in dados.get("atividades_secundarias", [])
            ]),
            "socios": "; ".join([
                f"{s.get('nome', '')} ({s.get('qual', '')})" for s in dados.get("qsa", [])
            ])
        }
    except Exception as e:
        print(f"[{cnpj}] Erro na requisição: {e}")
        return None


dados_coletados = []

for cnpj in df_cnpjs["cnpj"]:
    dados = consultar_cnpj(cnpj)
    if dados:
        dados_coletados.append(dados)
        print(f"[OK] CNPJ {cnpj} coletado.")
    else:
        print(f"Falha ao consultar {cnpj}")
    time.sleep(20)


df_resultado = pd.DataFrame(dados_coletados)


if os.path.exists(ARQUIVO_RESULTADO):
    df_existente = pd.read_excel(ARQUIVO_RESULTADO)
    df_resultado = pd.concat([df_existente, df_resultado], ignore_index=True)

df_resultado.drop_duplicates(subset='cnpj', inplace=True)

df_resultado.to_excel(ARQUIVO_RESULTADO, index=False)
print("Arquivo resultado_cnpjs.xlsx atualizado!")


# try:
#     engine = create_engine(f"mysql+pymysql://{DB_USUARIO}:{DB_SENHA}@{DB_HOST}:{DB_PORTA}/{DB_NOME}")

#     df_resultado.drop_duplicates(subset='cnpj', inplace=True)

#     with engine.connect() as conn:
#         cnpjs_existentes = pd.read_sql(f"SELECT cnpj FROM {TABELA_SQL}", conn)

#     def formatar_cnpj(cnpj):
#         cnpj = ''.join(filter(str.isdigit, str(cnpj)))
#         if len(cnpj) == 14:
#             return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
#         return cnpj

#     df_resultado['cnpj_formatado'] = df_resultado['cnpj'].apply(formatar_cnpj)

#     df_resultado['cnpj_limpo'] = df_resultado['cnpj_formatado'].str.replace(r'\D', '', regex=True)
#     cnpjs_existentes = cnpjs_existentes['cnpj'].astype(str).str.replace(r'\D', '', regex=True)

#     df_resultado = df_resultado[~df_resultado['cnpj_limpo'].isin(cnpjs_existentes)]

#     df_resultado['cnpj'] = df_resultado['cnpj_formatado']
#     df_resultado = df_resultado.drop(columns=['cnpj_formatado', 'cnpj_limpo'])

#     print(f"{len(df_resultado)} novos CNPJs para inserir no banco.")

#     if not df_resultado.empty:
#         df_resultado.to_sql(TABELA_SQL, con=engine, if_exists='append', index=False)
#         print(f"Dados adicionados à tabela '{TABELA_SQL}' do banco '{DB_NOME}'.")
#     else:
#         print("Nenhum novo CNPJ para inserir no banco.")
# except Exception as e:
#     print("Erro ao inserir no banco de dados:", e)

    
# db_config.py
from sqlalchemy import engine_from_config
import pandas as pd
import os

# Definir as variáveis da conexão
servername = "spsvsql39\\metas"
dbname = "HubDados"
driver = "ODBC+Driver+17+for+SQL+Server"

# Configuração do banco de dados em um dicionário
config = {
    'sqlalchemy.url': 
    f'mssql+pyodbc://@{servername}/{dbname}?trusted_connection=yes&driver={driver}'
}

def create_engine():
    """Usando engine_from_config para criar a engine a partir da configuração"""
    return engine_from_config(config, prefix='sqlalchemy.')
    print("Conexão realizada.")

def close_connection(engine):
    """Liberar os recursos e encerrar a conexão"""
    if engine:
        engine.dispose()
        print("Conexão encerrada com sucesso.")


def insert_df_new_engine(df: pd.DataFrame, nome_arquivo: str, nome_tabela: str = "tabela_dados"):
    """
    Salva o DataFrame em uma nova engine SQLite.

    :param df: DataFrame a ser salvo.
    :param nome_arquivo: Nome do arquivo .db a ser criado.
    :param nome_tabela: Nome da tabela onde os dados serão inseridos.
    """
    df['data_atualizacao'] = pd.Timestamp.today().normalize()
    servername = "spsvsql39\\metas"
    dbname = "FINANCA"
    driver = "ODBC+Driver+17+for+SQL+Server"

    # Configuração do banco de dados em um dicionário
    config = {
    'sqlalchemy.url': 
    f'mssql+pyodbc://@{servername}/{dbname}?trusted_connection=yes&driver={driver}'
    }
    caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)
    engine_nova = engine_from_config(config, prefix='sqlalchemy.')

    with engine_nova.connect() as conexao:
        

        df.to_sql(nome_tabela, con=conexao, if_exists='replace', index=False)
    
    print(f"DataFrame salvo em '{caminho_arquivo}' na tabela '{nome_tabela}'.")

# Exemplo de uso no seu main:
# df = pd.read_sql_query(sql=text(queries["query_total"]), con=connection)
# salvar_df_em_nova_engine(df, "dados_provisao.db")

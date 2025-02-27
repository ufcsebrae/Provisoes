# db_config.py
from sqlalchemy import create_engine, event,engine_from_config
from sqlalchemy.engine import Engine
import pandas as pd


# Definir as variáveis da conexão
servername = "spsvsql39\\metas"
dbname = "HubDados"
driver = "ODBC+Driver+17+for+SQL+Server"

# Configuração do banco de dados em um dicionário
config = {
    'sqlalchemy.url': 
    f'mssql+pyodbc://@{servername}/{dbname}?trusted_connection=yes&driver={driver}'
}

def create_engine_():
    """Usando engine_from_config para criar a engine a partir da configuração"""
    return engine_from_config(config, prefix='sqlalchemy.')
    print("Conexão realizada.")

def close_connection(engine):
    """Liberar os recursos e encerrar a conexão"""
    if engine:
        engine.dispose()
        print("Conexão encerrada com sucesso.")


def insert_df_new_engine(df: pd.DataFrame, nome_arquivo: str, nome_tabela: str = "Provisao"):
    # ... (mantenha as conversões de dados anteriores)

    # String de conexão ajustada
    connection_string = (
        "mssql+pyodbc://@spsvsql39\\metas/FINANCA?"
        "trusted_connection=yes&"
        "driver=ODBC+Driver+17+for+SQL+Server"
    )

    # Configurar engine e habilitar fast_executemany via evento
    engine = create_engine(connection_string)

    # Evento para configurar o cursor do pyodbc
    @event.listens_for(engine, "connect")
    def configure_fast_executemany(conn, _):
        cursor = conn.cursor()
        cursor.fast_executemany = True  # 🚀 Ativação direta no cursor

    # Inserção dos dados
    with engine.begin() as conn:
        df.to_sql(
            nome_tabela,
            con=conn,
            schema="dbo",
            if_exists='replace',
            index=False,
            chunksize=75  # 75 linhas × 28 colunas = 2100 parâmetros
        )

    print(f"✅ DataFrame salvo na tabela '{nome_tabela}'.")
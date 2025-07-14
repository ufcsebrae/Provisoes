# db_config.py
from sqlalchemy import create_engine, event,engine_from_config
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError
import pandas as pd


# Definir as variáveis da conexão
servername = "spsvsql39"
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
    df['data_atualizacao'] = pd.to_datetime('today').date()

    connection_string = (
        "mssql+pyodbc://@spsvsql39/FINANCA?"
        "trusted_connection=yes&"
        "driver=ODBC+Driver+17+for+SQL+Server&"
        "timeout=30"
    )

    try:
        print("🔗 Criando engine...")
        engine = create_engine(connection_string)

        # Aplicar fast_executemany corretamente
        @event.listens_for(engine, "before_cursor_execute")
        def receive_before_cursor_execute(conn, cursor, statement, parameters, context, executemany):
            if executemany:
                cursor.fast_executemany = True

        with engine.begin() as conn:
            print("📤 Inserindo dados no banco...")
            df.to_sql(
                nome_tabela,
                con=conn,
                schema="dbo",
                if_exists='append',
                index=False,
                chunksize=50  # Reduzido para maior estabilidade
            )

        print(f"✅ DataFrame salvo na tabela '{nome_tabela}' com sucesso!")

    except SQLAlchemyError as e:
        print("❌ Erro durante a inserção no banco:")
        print(e)

    finally:
        engine.dispose()
        print("🔌 Conexão encerrada.")
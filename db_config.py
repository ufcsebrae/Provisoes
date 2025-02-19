# db_config.py
from sqlalchemy import engine_from_config

# Definir as variáveis
servername = "spsvsql39\\metas"
dbname = "HubDados"
driver = "ODBC+Driver+17+for+SQL+Server"

# Configuração do banco de dados em um dicionário
config = {
    'sqlalchemy.url': 
    f'mssql+pymssql://@{servername}/{dbname}?trusted_connection=yes&driver={driver}'
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

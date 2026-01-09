from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError

# String de conexão com trusted_connection e timeout
connection_string = (
    "mssql+pyodbc://@spsvsql39/FINANCA?"
    "trusted_connection=yes&"
    "driver=ODBC+Driver+17+for+SQL+Server&"
    "timeout=10"
)

try:
    print("🔗 Criando engine...")
    engine = create_engine(connection_string)

    with engine.connect() as conn:
        print("✅ Conexão estabelecida com sucesso!")
        result = conn.execute(text("SELECT TOP 1 * FROM INFORMATION_SCHEMA.TABLES"))
        row = result.fetchone()
        print("📄 Resultado do SELECT:")
        print(row)
except SQLAlchemyError as e:
    print("❌ Erro ao conectar ou executar consulta:")
    print(e)
finally:
    engine.dispose()
    print("🔌 Conexão encerrada.")

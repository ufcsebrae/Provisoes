from db_config import create_engine,close_connection
from query_2 import queries
from sqlalchemy import text
import pandas as pd

#estabelecendo conexão
engine = create_engine()
try:
    with engine.connect() as connection:
        # Executar a query
        result = connection.execute(text(queries["query_total"]))
        
        # Converter resultados para DataFrame
        df = pd.DataFrame(result.fetchall(), columns=result.keys())
        
        # Exportar para Excel
        df.to_excel("PROVISAO.xlsx", index=False, sheet_name='Dados')
        print("Arquivo 'PROVISAO.xlsx' gerado com sucesso!")

except Exception as e:
    print(f"Erro durante a execução: {e}")

finally:
    # Encerrar conexão
    engine.dispose()
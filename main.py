from db_config import create_engine_,close_connection,insert_df_new_engine
import os
from time import gmtime, strftime
from query import queries
from sqlalchemy import text
from datetime import datetime
import pandas as pd
from excel_generator import ExcelGenerator
import time  # 👈 import para o timer
import cProfile  # 👈 Import para profiling
from email_sender import enviar_relatorio_email  # 👈 Novo import

# 👇 Inicia o timer global
start_time = time.time()
data = strftime('%d-%m-%Y', gmtime())
def main():
    #estabelecendo conexão
    engine = create_engine_()
    #inicia relatório
    relatorio = ExcelGenerator(os.path.join(os.getcwd(), "provisao_" + data + ".xlsx"))
    dataframes_email = {}  # 👈 Dicionário para armazenar DataFrames para o email

    try:
        with engine.connect() as connection:
            hoje = datetime.now().strftime('%Y-%m-%d')
            print(f"Data de hoje: {hoje}")

            # Consulta para obter a data máxima de atualização
            query_max_data = "SELECT MAX(data_atualizacao) FROM financa.dbo.Provisao"
            max_data = pd.read_sql_query(sql=text(query_max_data), con=connection).iloc[0, 0]

            # Verifica se a data máxima é igual à data de hoje
            if str(max_data) == hoje:
                print("Encontrado dados de hoje, deletando...")
                    # Query para deletar registros com a data de hoje     
                connection.execute(
                        text("DELETE FROM financa.dbo.orcado WHERE data_atualizacao = :hoje"),
                        {"hoje": hoje}
                )
            # 🔍 Carrega os resultados diretamente em um DataFrame
            df = pd.read_sql_query(sql=text(queries["query_total"]),con=connection)
            insert_df_new_engine(df,"provisao_","Provisao")            
            #decide incluir no relatório
            incluir_no_relatorio = True  # False para não incluir

            if incluir_no_relatorio:
                relatorio.add_sheet(df,"Dados Gerais")# 📄 Nome da Aba
                
            # 🔍 Query 2 - Movimentos Pendentes
            td = pd.pivot_table(df
            ,values=['V_PROVISIONADO','V_PAGO','MOVIMENTOS_PENDENTES_PAGAMENTO']
            ,index='RAZÃOSOCIAL'
            ,aggfunc="sum"
            ).reset_index() # Converte o índice em coluna

            # Ordena por MOVIMENTOS_PENDENTES_PAGAMENTO (decrescente)
            td = td.sort_values(by='MOVIMENTOS_PENDENTES_PAGAMENTO', ascending=False)

            #Remover completamente o índice numérico do DataFrame
            td = td.reset_index(drop=True)  # 👈 Remove o índice numérico
            dataframes_email["Pendentes"] = td # 👈 Armazena para email
                      

            #decide incluir no relatório
            incluir_no_relatorio = True  # False para não incluir
            if incluir_no_relatorio:
                relatorio.add_sheet(td,"tdPendentes")# 📄 Nome da Aba

            '''# ❓ Decisão para segunda aba
            if input("Deseja incluir a segunda query? (s/n): ").lower() == 's':
                relatorio.add_sheet(td, "tdPendentes")'''

            # Extraindo o trimestre da DATAEMISSAO
            df['TRIMESTRE_EMISSAO'] = df['DATAEMISSAO'].dt.quarter
             # 🔍 Query 3- Pivotando Unidade
            td2 = (
            df.groupby(['UNIDADE', 'ANO_PROVISAO','TRIMESTRE_EMISSAO'])
            .agg(
                IDENTIFICADOR=('MOVIMENTOS_PENDENTES_PAGAMENTO', 'count'),
                MOVIMENTOS_PENDENTES=('MOVIMENTOS_PENDENTES_PAGAMENTO', 'sum')
            )
            .reset_index()
            .rename(columns={
                'ANO_PROVISAO': 'ANO',
                'MOVIMENTOS_PENDENTES_PAGAMENTO': 'MOVIMENTOS PENDENTES',
                'TRIMESTRE_EMISSAO':'TRIMESTRE'
            })
             )
            # Remove índice numérico
            td2 = td2.reset_index(drop=True)
            dataframes_email["Unidades"] = td2  # 👈 Armazena para email

            #decide incluir no relatório
            incluir_no_relatorio = True  # False para não incluir
            if incluir_no_relatorio:
                relatorio.add_sheet(td2,"tdUnidades")# 📄 Nome da Aba


            '''# ❓ Decisão para terceita aba
            if input("Deseja incluir a segunda query? (s/n): ").lower() == 's':
                relatorio.add_sheet(td2, "tdUnidades")'''
            print(f"Arquivo {relatorio.file_name} gerado com sucesso!")

    except Exception as e: #❌ printa erro
        print(f"Erro durante a execução: {e}")

    finally:
        #3. ✅ Salva o arquivo (só se pelo menos uma aba foi adicionada)
        if relatorio.sheets_added:
            caminho_arquivo = relatorio.save()
            if caminho_arquivo:  # 👈 Verifica se o caminho foi retornado
                enviar_relatorio_email(dataframes_email, caminho_arquivo)   
        else:
            print("Nenhuma planilha foi adicionada ao relatório")
    # Encerrar conexão
        close_connection(engine)

if __name__ == "__main__":
    # Para profiling: execute com python -m cProfile seu_script.py
    # Ou descomente a linha abaixo para profiling interno:
    #cProfile.run('main()', sort='cumtime')
    
    main()
    # 👇 Mostra tempo de execução ao final
    print(f"\n🕒 Tempo total de execução: {time.time() - start_time:.2f} segundos") 
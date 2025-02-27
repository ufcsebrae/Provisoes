import matplotlib.pyplot as plt
import win32com.client
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
from io import BytesIO
from dfpandas import _pandas    



def enviar_relatorio_email(dataframes_email, caminho_excel):
    data_hoje = datetime.today().strftime("%d/%m/%Y")
    
    FORNECEDOR_MAIOR = dataframes_email['Pendentes'].iloc[0].get("RAZÃOSOCIAL")
    VALOR_MAIOR = f"R$ {dataframes_email['Pendentes'].iloc[0].get('MOVIMENTOS_PENDENTES_PAGAMENTO', 0):,.0f}".replace(",", ".")
    dfTRIMESTRE = (
    dataframes_email['Unidades']
    .groupby(['ANO', 'TRIMESTRE'])['MOVIMENTOS_PENDENTES']
    .sum()
    .reset_index()
    .sort_values(by='MOVIMENTOS_PENDENTES', ascending=False)
    )

    # Pegar a linha com o maior valor (contendo ANO, TRIMESTRE e MOVIMENTOS_PENDENTES)
    maior_registro = dfTRIMESTRE.iloc[0]

    # Extrair os valores
    MAIOR_ANO = maior_registro['ANO']
    MAIOR_TRIMESTRE = maior_registro['TRIMESTRE']
    MAIOR_VALOR = f"R$ {maior_registro.get('MOVIMENTOS_PENDENTES', 0):,.0f}".replace(",", ".")

    # Configurações do e-mail
    DESTINATARIOS = ["cesargl@sebraesp.com.br"]
    ASSUNTO = f"Controle Provisão Serviços - {data_hoje}"
    TEXTO_CORPO = f"""<p>Prezados,</p>
        <p>Segue relatório da provisão de serviços do dia <strong>{data_hoje}</strong></p>
        <p>O fornecedor com maio valor pendente é <strong>{FORNECEDOR_MAIOR}</strong>, totalizando <strong>{VALOR_MAIOR}</strong>.</p>
        <p>Na segunda tabela, temos os valores de provisão por trimestre. Podemos destacar o <strong>{MAIOR_TRIMESTRE}º trimestre de {MAIOR_ANO}</strong></p>
        <p>com um total de <strong>{MAIOR_VALOR}</strong>.</p>
        <p>Em anexo, segue o arquivo Excel com os detalhes.</p>
        <p>Atenciosamente,</p>
        """
    tabelas_html = {}

    # 🔎 Verificação do conteúdo de dataframes_email
    print("🔎 DataFrames disponíveis:", dataframes_email.keys())
    
    tabelas_html = _pandas(dataframes_email)
           
    # 🔎 Verificação final dos dados a serem renderizados
    print("📝 Tabelas HTML:", "Geradas" if tabelas_html else "Nenhuma gerada.")
  

    # Carrega o template
    env = Environment(loader=FileSystemLoader("."))
    try:
        template = env.get_template("email_template.html")
    except Exception as e:
        print(f"❌ Erro ao carregar o template: {e}")
        return


    # Renderiza o HTML
    CORPO = template.render(
        assunto=ASSUNTO,    
        tabelas=tabelas_html,
        data_envio=data_hoje,
        texto_email=TEXTO_CORPO
    )

    # Envia o e-mail
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mensagem = outlook.CreateItem(0)
        mensagem.Subject = ASSUNTO
        mensagem.HTMLBody = CORPO
        mensagem.To = "; ".join(DESTINATARIOS)

        if os.path.exists(caminho_excel):
            mensagem.Attachments.Add(os.path.abspath(caminho_excel))
            print("📎 Arquivo Excel anexado.")
        else:
            print(f"⚠️ Arquivo Excel não encontrado: {caminho_excel}")

        mensagem.Send()
        print("📧 E-mail enviado com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao enviar o e-mail: {e}")

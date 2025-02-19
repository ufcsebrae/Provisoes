
import win32com.client as win32
import os
import pandas as pd
from time import gmtime, strftime


def gerar_corpo_email(dados):
    primeira_linha_provisao1 = dados[0]['detalhe'][0].get("MÊS", "Desconhecido")
    primeira_linha_provisao2 = dados[0]['detalhe'][0].get("ANO", "Desconhecido")
    primeira_linha_fornecedores = dados[1]['detalhe'][0].get("RAZÃO SOCIAL", "Desconhecido")
    segunda_linha_fornecedores = dados[1]['detalhe'][1].get("RAZÃO SOCIAL", "Desconhecido")
    primeira_linha_pagamento = dados[2]['detalhe'][0].get("MÊS", "Desconhecido") 

    valor_total_provisao = dados[0]['detalhe'][0].get("VALOR DOCUMENTO", 0)
    valor_total_fornecedores1 = dados[1]['detalhe'][0].get("VALOR PROVISIONADO", 0)
    valor_total_pagamento1 = dados[2]['detalhe'][0].get("VALOR PAGO", 0)

    valor_total_provisao_formatado = f"R$ {valor_total_provisao:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    valor_total_fornecedores_formatado = f"R$ {valor_total_fornecedores1:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    valor_total_pagamento_formatado = f"R$ {valor_total_pagamento1:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    # Gerar a data atual formatada
    from time import gmtime, strftime
    data_atual = strftime('%d/%m/%Y', gmtime())
    corpo = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
                line-height: 1.6;
                font-size: 14px;
                color: #333;
                background-color: #f9f9f9;
            }}
            .container {{
                background: #fff;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 20px;
                max-width: 800px;
                margin: auto;
                box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);
            }}
            h2 {{
                color: #007BFF;
                text-align: center;
                font-size: 22px;
                margin-bottom: 20px;
            }}
            h3 {{
                color: #555;
                font-size: 16px;
                margin-top: 10px;
                margin-bottom: 5px;
            }}
            table {{
                width: 80%;
                border-collapse: collapse;
                margin: 15px auto;
                font-size: 14px;
                background-color: #fff;
                color: #000;
                border: 1px solid #333;
            }}
            th, td {{
                border: 1px solid #333;
                text-align: left;
                padding: 8px;
            }}
            th {{
                background-color: #f4f4f4;
                color: #333;
                font-weight: bold;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            .footer {{
                text-align: center;
                margin-top: 20px;
                font-size: 12px;
                color: #888;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>CONTROLE PROVISÃO DE SERVIÇOS</h2>
            <p>Prezados,</p>
            <p>Segue um resumo do relatório Controle Provisão de Serviços até a data <strong>{data_atual}</strong>.</p>
            <p>Provisão Mensal: O mês e o ano com maior valor de documento é <strong>{primeira_linha_provisao1}{primeira_linha_provisao2}</strong>, sendo de <strong>{valor_total_provisao_formatado}</strong> .</p>
            <p>Fornecedores: Os principais fornecedores incluem <strong>{primeira_linha_fornecedores}</strong>, que tem o <strong>maior</strong> valor provisionado (<strong>{valor_total_fornecedores_formatado}</strong>), seguido pela <strong>{segunda_linha_fornecedores}</strong> e outros. Há valores pendentes de pagamento para diversos fornecedores.</p>
            <p>Pagamentos Realizados: Os pagamentos mensais variam, sendo <strong>{primeira_linha_pagamento}</strong> o mês marcado pelo valor de <strong>{valor_total_pagamento_formatado}</strong>.</p>
            <p> </p>
            
    """

    for resumo in dados[:1]:
        provisao_total = 0
        valor_documento = 0

# Somando os valores dentro da lista "detalhe"
        for detalhe in resumo.get("detalhe", []):
            provisao_total += detalhe.get("QUANTIDADE DE MOVIMENTO", 0) 
            valor_documento += detalhe.get("VALOR DOCUMENTO", 0)

# Formatar o valor total com separador de milhar no formato brasileiro
        valor_total_formatado = f"{valor_documento:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        corpo += f"""
        <div>
            <h3>Quantidade de Movimento: {provisao_total}</h3>
            <h3>Valor Documento: R$ {valor_total_formatado}</h3>
        </div>
        """

        # Verifique se "detalhe" existe e é uma lista válida
        if not resumo.get("detalhe"):
            corpo += """
            <p style="color: red;">Nenhum detalhe encontrado para este resumo.</p>
            """
            continue

        # Tabela para Tipo de Pagamento, Quantidade e Valor Total
        corpo += """
        <table>
            <thead>
                <tr>
                    <th>ANO</th>
                    <th>MÊS</th>
                    <th>VALOR DOCUMENTO</th>
                    <th>QUANTIDADE DO MOVIMENTO</th>
                </tr>
            </thead>
            <tbody>
        """

        # Criação do dicionário para agrupar por "Tipo de Pagamento"
        provisao_resumo = {}
        for detalhe in resumo["detalhe"]:
            ano = detalhe.get("ANO", 0)
            mes = detalhe.get("MÊS", 0)
            valor_documento_total = detalhe.get("VALOR DOCUMENTO", 0)
            quantidade_de_movimento = detalhe.get("QUANTIDADE DE MOVIMENTO", 0)

            chave = (ano, mes)

            if chave not in provisao_resumo:
                provisao_resumo[chave] = {'quantidade': 0, 'valor_total': 0}

            provisao_resumo[chave]['quantidade'] += quantidade_de_movimento
            provisao_resumo[chave]['valor_total'] += valor_documento_total 

        for (ano, mes), valores in provisao_resumo.items():
            corpo += f"""
                 <tr>
                    <td>{ano}</td>
                    <td>{mes}</td>
                    <td>R$ {valores['valor_total']:.2f}</td>
                    <td>{valores['quantidade']}</td>
                </tr>
                """

        corpo += """
            </tbody>
        </table>
        <br>
        """

    corpo += """
        </div>
        <div class="footer">
            <p>Este é um e-mail automático. Por favor, não responda.</p>
        </div>
    </body>
    </html>
    """
    return corpo

 
 
def enviar_email(destinatario, copiar, assunto, corpo, caminho_anexo):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  
    mail.To = destinatario
    mail.cc = copiar
    mail.Subject = assunto
    mail.HTMLBody = corpo
    if caminho_anexo and os.path.exists(caminho_anexo):
        mail.Attachments.Add(caminho_anexo)
   
    mail.Send()
    print("E-mail enviado com sucesso!") 
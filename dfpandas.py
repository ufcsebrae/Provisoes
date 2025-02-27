import pandas as pd
import locale
from datetime import datetime

def formatar_moeda(valor):
    """Formata valores monetários no padrão brasileiro."""
    return f'R$ {valor:,.0f}'.replace(',', 'virgula_temp').replace('.', ',').replace('virgula_temp', '.')

def _pandas(dataframes):
    tabelas_html = {}
    
    for nome_aba, df in dataframes.items():
        print(f"➡️ Processando aba: {nome_aba}")

        if df.empty:
            print(f"⚠️ DataFrame '{nome_aba}' está vazio.")
            continue

        if nome_aba == 'Unidades':
            # Renomeia as colunas
            df = df.rename(columns={
                'IDENTIFICADOR': 'Qtde.Movimentos',
                'MOVIMENTOS_PENDENTES': 'Valor Pendente',
            })

            # Converte 'ANO' para inteiro e filtra anos válidos
            ano_atual = datetime.now().year
            df['ANO'] = df['ANO'].astype(int)
            df_filtered = df[df['ANO'].between(2023, ano_atual) & (df['Valor Pendente'] > 0)]

            # Ordena e pivota o DataFrame
            df_pivot = pd.pivot_table(df_filtered, index=['UNIDADE', 'ANO', 'TRIMESTRE'], aggfunc='sum').sort_values(
                by=['UNIDADE', 'ANO', 'TRIMESTRE']
            )

            # Aplica formatação
            styled_df = df_pivot.style.format({
                'Valor Pendente': formatar_moeda,
                'Qtde.Movimentos': '{:.0f}'
            })

            tabelas_html[nome_aba] = styled_df.to_html(classes="table", border=0, index=False)
            print(f"✅ Tabela gerada para aba '{nome_aba}'.")

        elif nome_aba == 'Pendentes':
            # Seleção e renomeação de colunas
            df = df.rename(columns={
                'RAZÃOSOCIAL': 'Fornecedor',
                'V_PROVISIONADO': 'Valor Provisionado',
                'V_PAGO': 'Valor Pago',
                'MOVIMENTOS_PENDENTES_PAGAMENTO': 'Valor Pendente',
            })[['Fornecedor', 'Valor Provisionado', 'Valor Pago', 'Valor Pendente']]

            # Aplica formatação
            styled_df = df.style.format({
                'Valor Provisionado': formatar_moeda,
                'Valor Pago': formatar_moeda,
                'Valor Pendente': formatar_moeda,
            })

            tabelas_html[nome_aba] = styled_df.hide(axis="index").to_html(classes="table", border=0, index=False)
            print(f"✅ Tabela gerada para aba '{nome_aba}'.")

    return tabelas_html

# Provisões

Automatize a geração de relatórios de provisão de serviços com Python, exporte em Excel e envie tudo por e-mail com facilidade.

Este projeto é uma aplicação Python para gerar relatórios de provisão de serviços, salvar os resultados em um arquivo Excel e enviar esses relatórios por e-mail.


## Estrutura do Projeto
__pycache__/
_main.py
_query.py
.gitignore
db_config.py
dfpandas.py
email_sender.py
email_template.html
enviaremail.py
excel_gerador.py
main.py
query.py
requirements.txt

## Arquivos Principais

- **main.py**: Arquivo principal que executa o fluxo completo do projeto.
- **db_config.py**: Configurações e funções para conexão com o banco de dados.
- **dfpandas.py**: Funções para manipulação e formatação de DataFrames.
- **email_sender.py**: Funções para enviar e-mails com os relatórios gerados.
- **excel_gerador.py**: Classe para gerar e salvar arquivos Excel.
- **query.py**: Contém as consultas SQL utilizadas no projeto.
- **email_template.html**: Template HTML para o corpo do e-mail.

## Requisitos

- Python 3.11
- Bibliotecas listadas em [`requirements.txt`](requirements.txt )

## Instalação

1. Clone o repositório:
    ```sh
    git clone git@github.com:ufcsebrae/Provisoes.git
    cd provisoes
    ```

2. Crie um ambiente virtual e ative-o:
    ```sh
    python -m venv .venv
    .venv\Scripts\activate  # No Windows
    source .venv/bin/activate  # No Linux/Mac
    ```

3. Instale as dependências:
    ```sh
    pip install -r requirements.txt
    ```

## Uso

1. Configure as variáveis de conexão no arquivo [`db_config.py`](db_config.py ).

2. Execute o script principal:
    ```sh
    python main.py
    ```

## Funcionalidades

- **Consulta SQL**: Executa consultas SQL no banco de dados configurado.
- **Geração de Relatórios**: Gera relatórios em formato Excel.
- **Envio de E-mails**: Envia os relatórios gerados por e-mail.

## Contribuição

1. Faça um fork do projeto.
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`).
3. Commit suas alterações (`git commit -am 'Adiciona nova feature'`).
4. Faça um push para a branch (`git push origin feature/nova-feature`).
5. Abra um Pull Request.

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.

---



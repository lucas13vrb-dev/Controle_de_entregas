# Controle de Entregas de Produtores

Aplicativo em Streamlit para lancar entregas de frutas por produtor rural, quantidade e destino.

## Arquivos principais

- `lancamento_entregas.py`: aplicativo visual.
- `data/entregas_produtores.xlsx`: base de dados dos lancamentos.
- `requirements.txt`: dependencias Python necessarias.
- `.streamlit/secrets.toml.example`: exemplo para configurar Google Sheets no Streamlit Cloud.

## Como executar

Instale as dependencias:

```powershell
pip install -r requirements.txt
```

Abra o sistema:

```powershell
streamlit run lancamento_entregas.py
```

Para usar na rede local:

```powershell
streamlit run lancamento_entregas.py --server.address 0.0.0.0 --server.port 8502
```

## Base compartilhada no Streamlit Cloud

Para varias pessoas usarem o mesmo link, configure uma planilha Google como banco de dados:

1. Crie uma planilha no Google Sheets.
2. Copie o ID da URL da planilha.
3. Crie uma Service Account no Google Cloud e baixe o JSON da chave.
4. Compartilhe a planilha com o `client_email` da Service Account como Editor.
5. No Streamlit Cloud, abra o app, clique em `Settings` > `Secrets`.
6. Cole o conteudo seguindo o modelo de `.streamlit/secrets.toml.example`.

Quando os secrets estiverem configurados, o app usa o Google Sheets automaticamente.
Sem secrets, ele usa o Excel local em `data/entregas_produtores.xlsx`.

# Controle de Entregas de Produtores

Aplicativo em Streamlit para lancar entregas de frutas por produtor rural, quantidade e destino.

## Arquivos principais

- `lancamento_entregas.py`: aplicativo visual.
- `data/entregas_produtores.xlsx`: base de dados dos lancamentos.
- `requirements.txt`: dependencias Python necessarias.

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

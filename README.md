# Gerenciador de promocoes militares

Aplicacao web local em Python/Streamlit para processar promocoes alternadas por antiguidade e merecimento, com importacao de planilha e decisoes feitas por cliques.

## Como rodar

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Se o comando `python` nao estiver disponivel no Windows, instale Python 3.11+ ou use o executavel Python ja configurado na sua maquina. Depois abra o endereco exibido pelo Streamlit no navegador, normalmente `http://localhost:8501`.

## Planilha de entrada

Formatos aceitos: `.xlsx` e `.csv`.

Colunas obrigatorias:

- `Nome`
- `Antiguidade`
- `Merecimento`
- `Agregado`

O campo `Agregado` aceita valores como `Sim`, `Nao`, `S`, `N`, `True`, `False`, `1`, `0`, `Yes`, `No`.

## Regras implementadas

- Alternancia automatica entre antiguidade e merecimento a partir do criterio inicial escolhido.
- Exclusao cruzada: todo promovido sai imediatamente das duas filas.
- Agregado e promovido, mas nao consome vaga principal.
- Antiguidade seleciona automaticamente o primeiro elegivel da fila.
- Primeira disputa por merecimento tem 2 nao agregados.
- Demais disputas por merecimento tem 3 nao agregados.
- Agregados encontrados na composicao do merecimento sao promovidos sem entrar na disputa.
- No merecimento, o usuario escolhe por clique o promovido.
- Nas disputas de 3, o usuario escolhe por clique quem desce para a proxima disputa.
- O outro remanescente vai para o final da fila de merecimento.
- Exportacao de promovidos, historico e filas em CSV/XLSX.

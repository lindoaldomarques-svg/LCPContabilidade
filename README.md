# Aplicativo para baixar extrato do Banco do Brasil (API Extratos)

Este projeto implementa um aplicativo em **FastAPI** para consultar e baixar extrato em CSV usando a API do Banco do Brasil.

> ⚠️ Observação importante: no ambiente de execução desta tarefa, o acesso ao portal `https://developers.bb.com.br/` retornou **HTTP 403**, então não foi possível validar o catálogo diretamente. O app foi estruturado para seguir o padrão oficial do BB (OAuth2 + `X-Developer-Application-Key`) e permite ajuste fino dos endpoints via `.env`.

## Funcionalidades

- `GET /health` — health check.
- `GET /extrato` — consulta extrato (JSON).
- `GET /extrato/download` — baixa extrato em CSV.

## Pré-requisitos

- Python 3.11+
- Cadastro do aplicativo no portal Developers BB
- Credenciais OAuth + Application Key do produto de extratos

## Configuração

1. Instale dependências:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Crie o arquivo de ambiente:

```bash
cp .env.example .env
```

3. Preencha os dados no `.env` com as informações do seu app no BB:

- `BB_CLIENT_ID`
- `BB_CLIENT_SECRET`
- `BB_DEVELOPER_KEY`
- `BB_TOKEN_URL`
- `BB_API_BASE_URL`
- `BB_EXTRATO_PATH`
- `BB_SCOPE`

## Execução

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Documentação Swagger:

- `http://localhost:8000/docs`

## Exemplos de uso

### Consultar extrato (JSON)

```bash
curl "http://localhost:8000/extrato?agencia=1234&conta=123456&data_inicio=2026-01-01&data_fim=2026-01-31"
```

### Baixar extrato em CSV

```bash
curl -L -o extrato.csv "http://localhost:8000/extrato/download?agencia=1234&conta=123456&data_inicio=2026-01-01&data_fim=2026-01-31"
```

## Notas sobre a API BB

- O app usa `grant_type=client_credentials` para obter token OAuth.
- O header `X-Developer-Application-Key` é enviado na consulta do extrato.
- Como os contratos podem variar por produto/versão do portal BB, os caminhos e escopo são configuráveis por variáveis de ambiente.

## Testes

```bash
pytest -q
```

# VBA - API Extratos Banco do Brasil

Este repositório contém um módulo VBA para baixar extratos do Banco do Brasil via API.

## Arquivo principal

- `vba/bb_extratos.bas`

## Como usar

1. Abra o Excel e pressione `ALT + F11`.
2. Importe o arquivo `vba/bb_extratos.bas` no seu projeto VBA.
3. Edite as constantes no topo do módulo:
   - `CLIENT_ID`
   - `CLIENT_SECRET`
   - `APP_KEY`
   - `AGENCIA`
   - `CONTA`
   - `OAUTH_SCOPE`
   - `CLIENT_CERT_LOCATION` (se seu cenário exigir mTLS)
4. Execute a macro `BaixarExtratoBB`.
5. O JSON de retorno será salvo na aba `Extrato_BB_JSON`.

## Correção aplicada para erro de token OAuth2

Se antes aparecia a mensagem **"Não foi possível obter token OAuth2"**, o módulo agora:

- Envia `gw-dev-app-key` também na URL de token.
- Faz tentativa padrão com `Authorization: Basic` e fallback com `client_id/client_secret` no corpo.
- Força TLS 1.2 via WinHTTP.
- Exibe detalhes do erro (status HTTP e resposta da API) para facilitar diagnóstico.

## Observações

- Valide no portal do BB os endpoints exatos e os parâmetros do seu convênio.
- Alguns convênios exigem certificado cliente (mTLS) também no token.
- Em ambiente produtivo, recomenda-se usar parser JSON robusto (ex.: VBA-JSON) para quebrar lançamentos em colunas.

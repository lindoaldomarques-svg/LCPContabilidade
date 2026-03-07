# VBA - API Extratos Banco do Brasil

Este repositório contém um módulo VBA para baixar extratos do Banco do Brasil via API.

## Arquivo principal

- `vba/bb_extratos.bas`

## Já configurado

O módulo já foi deixado pré-configurado com:

- `APP_KEY`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `PRECOMPUTED_BASIC`

Se precisar trocar, basta editar as constantes no topo do arquivo.

## Como usar

1. Abra o Excel e pressione `ALT + F11`.
2. Importe o arquivo `vba/bb_extratos.bas` no seu projeto VBA.
3. Ajuste os campos da conta:
   - `AGENCIA`
   - `CONTA`
4. Se necessário, ajuste:
   - `OAUTH_SCOPE`
   - `CLIENT_CERT_LOCATION` (mTLS)
5. Execute a macro `BaixarExtratoBB`.
6. O JSON de retorno será salvo na aba `Extrato_BB_JSON`.

## Correção aplicada para erro de token OAuth2

Para o erro de transporte mostrado no Excel (`Erro 5 - Argumento ou chamada de procedimento inválida`), o módulo agora:

- Tenta autenticar com `WinHttp.WinHttpRequest.5.1`.
- Se falhar no transporte, faz fallback automático para `MSXML2.ServerXMLHTTP.6.0`.
- Continua enviando `gw-dev-app-key` na URL de token.
- Exibe detalhes de erro para facilitar diagnóstico.

## Observações

- Valide no portal do BB os endpoints exatos e os parâmetros do seu convênio.
- Alguns convênios exigem certificado cliente (mTLS) também no token.
- Em ambiente produtivo, recomenda-se parser JSON robusto (ex.: VBA-JSON) para quebrar lançamentos em colunas.

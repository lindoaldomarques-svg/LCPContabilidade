# VBA - API Extratos Banco do Brasil

Este repositório contém um módulo VBA pronto para baixar extratos do Banco do Brasil via API.

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

## Observações

- Valide no portal do BB os endpoints exatos e os parâmetros do seu convênio.
- Em ambiente produtivo, recomenda-se usar um parser JSON robusto (ex.: VBA-JSON) para quebrar os lançamentos em colunas.

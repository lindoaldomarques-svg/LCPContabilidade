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
- `CERT_PEM_PATH` = `C:\Eticons\API-Malta\Prefeitura\pm_malta.pem`
- `CERT_KEY_PATH` = `C:\Eticons\API-Malta\Prefeitura\chave_privada2.key`

## Ajuste aplicado para o erro da imagem (HTTP 400 invalid_scope)

Foi ajustado para retornar corretamente o token:

- `OAUTH_SCOPE` agora vem vazio por padrão (evita pedir escopo não autorizado).
- Se a API responder `invalid_scope`, o módulo refaz automaticamente a autenticação sem scope.
- Além de WinHTTP e ServerXMLHTTP, foi incluído fallback final com `curl.exe` usando os arquivos `.pem` e `.key`.

## Como usar

1. Abra o Excel e pressione `ALT + F11`.
2. Importe o arquivo `vba/bb_extratos.bas` no seu projeto VBA.
3. Ajuste os campos da conta:
   - `AGENCIA`
   - `CONTA`
4. Execute a macro `BaixarExtratoBB`.
5. O JSON de retorno será salvo na aba `Extrato_BB_JSON`.

## Observações

- Se seu convênio exigir escopo específico, preencha `OAUTH_SCOPE` com o valor autorizado no BB Developer.
- Se o caminho do certificado/chave mudar, ajuste `CERT_PEM_PATH` e `CERT_KEY_PATH`.
- Em ambiente produtivo, recomenda-se parser JSON robusto (ex.: VBA-JSON) para quebrar lançamentos em colunas.

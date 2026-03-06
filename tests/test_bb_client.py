from app.bb_client import BBApiClient, BBApiError
from app.config import Settings



def test_statement_to_csv():
    payload = {
        "lancamentos": [
            {"data": "2026-01-01", "descricao": "PIX", "valor": -100.5},
            {"data": "2026-01-02", "descricao": "SALARIO", "valor": 2500.0},
        ]
    }

    csv_content = BBApiClient.statement_to_csv(payload)

    assert "data,descricao,valor" in csv_content
    assert "2026-01-01,PIX,-100.5" in csv_content
    assert "2026-01-02,SALARIO,2500.0" in csv_content



def test_statement_to_csv_rejects_invalid_structure():
    payload = {"lancamentos": "texto-invalido"}

    try:
        BBApiClient.statement_to_csv(payload)
    except BBApiError as exc:
        assert "Formato de extrato inesperado" in str(exc)
    else:
        raise AssertionError("Era esperado BBApiError")



def test_missing_credentials_validation():
    settings = Settings(
        bb_client_id="",
        bb_client_secret="",
        bb_developer_key="",
        bb_token_url="https://oauth.bb.com.br/oauth/token",
        bb_api_base_url="https://api.bb.com.br",
        bb_extrato_path="/extratos/v1/conta-corrente/extrato",
        bb_scope="extrato-info",
    )
    client = BBApiClient(settings)

    try:
        client._validate_credentials()
    except BBApiError as exc:
        assert "BB_CLIENT_ID" in str(exc)
        assert "BB_CLIENT_SECRET" in str(exc)
        assert "BB_DEVELOPER_KEY" in str(exc)
    else:
        raise AssertionError("Era esperado BBApiError")

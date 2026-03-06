from __future__ import annotations

import csv
import io
from typing import Any

import requests

from app.config import Settings


class BBApiError(RuntimeError):
    pass


class BBApiClient:
    def __init__(self, settings: Settings, timeout: int = 30):
        self.settings = settings
        self.timeout = timeout

    def _validate_credentials(self) -> None:
        missing = []
        if not self.settings.bb_client_id:
            missing.append("BB_CLIENT_ID")
        if not self.settings.bb_client_secret:
            missing.append("BB_CLIENT_SECRET")
        if not self.settings.bb_developer_key:
            missing.append("BB_DEVELOPER_KEY")
        if missing:
            raise BBApiError(f"Variáveis obrigatórias não configuradas: {', '.join(missing)}")

    def get_access_token(self) -> str:
        self._validate_credentials()

        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "grant_type": "client_credentials",
            "scope": self.settings.bb_scope,
        }
        response = requests.post(
            self.settings.bb_token_url,
            headers=headers,
            data=data,
            auth=(self.settings.bb_client_id, self.settings.bb_client_secret),
            timeout=self.timeout,
        )

        if response.status_code >= 400:
            raise BBApiError(
                f"Falha ao autenticar no OAuth do BB ({response.status_code}): {response.text}"
            )

        payload = response.json()
        token = payload.get("access_token")
        if not token:
            raise BBApiError("Resposta de token não contém 'access_token'.")
        return token

    def get_statement(self, agencia: str, conta: str, data_inicio: str, data_fim: str) -> dict[str, Any]:
        token = self.get_access_token()
        base = self.settings.bb_api_base_url.rstrip("/")
        path = self.settings.bb_extrato_path.lstrip("/")
        url = f"{base}/{path}"

        headers = {
            "Authorization": f"Bearer {token}",
            "X-Developer-Application-Key": self.settings.bb_developer_key,
            "Accept": "application/json",
        }
        params = {
            "agencia": agencia,
            "conta": conta,
            "dataInicio": data_inicio,
            "dataFim": data_fim,
        }

        response = requests.get(url, headers=headers, params=params, timeout=self.timeout)
        if response.status_code >= 400:
            raise BBApiError(
                f"Falha ao consultar extrato ({response.status_code}): {response.text}"
            )

        return response.json()

    @staticmethod
    def statement_to_csv(payload: dict[str, Any]) -> str:
        lancamentos = payload.get("lancamentos") or payload.get("listaLancamentos") or []
        if not isinstance(lancamentos, list):
            raise BBApiError("Formato de extrato inesperado: não existe lista de lançamentos.")

        if not lancamentos:
            return ""

        # cria cabeçalhos dinâmicos com base nas chaves do primeiro lançamento
        fieldnames = sorted({k for item in lancamentos if isinstance(item, dict) for k in item.keys()})
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()
        for row in lancamentos:
            if isinstance(row, dict):
                writer.writerow(row)
        return output.getvalue()

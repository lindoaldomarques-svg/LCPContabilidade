from __future__ import annotations

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse, PlainTextResponse

from app.bb_client import BBApiClient, BBApiError
from app.config import load_settings

app = FastAPI(title="BB Extratos Downloader", version="1.0.0")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/extrato")
def consultar_extrato(
    agencia: str = Query(..., min_length=1),
    conta: str = Query(..., min_length=1),
    data_inicio: str = Query(..., description="Formato sugerido: YYYY-MM-DD"),
    data_fim: str = Query(..., description="Formato sugerido: YYYY-MM-DD"),
):
    settings = load_settings()
    client = BBApiClient(settings)
    try:
        payload = client.get_statement(
            agencia=agencia,
            conta=conta,
            data_inicio=data_inicio,
            data_fim=data_fim,
        )
        return JSONResponse(payload)
    except BBApiError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/extrato/download")
def baixar_extrato_csv(
    agencia: str = Query(..., min_length=1),
    conta: str = Query(..., min_length=1),
    data_inicio: str = Query(..., description="Formato sugerido: YYYY-MM-DD"),
    data_fim: str = Query(..., description="Formato sugerido: YYYY-MM-DD"),
):
    settings = load_settings()
    client = BBApiClient(settings)
    try:
        payload = client.get_statement(
            agencia=agencia,
            conta=conta,
            data_inicio=data_inicio,
            data_fim=data_fim,
        )
        csv_content = client.statement_to_csv(payload)
        filename = f"extrato_{agencia}_{conta}_{data_inicio}_{data_fim}.csv"
        headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return PlainTextResponse(csv_content, headers=headers, media_type="text/csv")
    except BBApiError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

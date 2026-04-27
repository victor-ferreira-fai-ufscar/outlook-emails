"""
Rotas de mensagens do Outlook.
"""

from typing import Any

from fastapi import APIRouter, Request

from app.utils import fetch_latest_sent_email, get_local_access_token

router = APIRouter(prefix="/messages", tags=["messages"])


@router.get("/sent/latest")
def get_latest_sent_email(request: Request) -> dict[str, Any]:
    """Retorna o ultimo email enviado do usuario autenticado."""
    access_token = get_local_access_token(request)
    return fetch_latest_sent_email(access_token)

"""
Rotas de perfil do usuario.
"""

from typing import Any

from fastapi import APIRouter, Request

from app.utils import fetch_outlook_profile, get_local_access_token, save_profile_json

router = APIRouter(prefix="/profile", tags=["profile"])


@router.get("")
def get_profile(request: Request) -> dict[str, Any]:
    """Retorna dados do perfil atual do usuario autenticado."""
    access_token = get_local_access_token(request)
    return fetch_outlook_profile(access_token)


@router.get("/export")
def export_profile_json(request: Request) -> dict[str, str]:
    """Exporta e salva snapshot do perfil em arquivo JSON."""
    access_token = get_local_access_token(request)
    profile = fetch_outlook_profile(access_token)
    json_path = save_profile_json(profile)

    return {"message": "Profile exported successfully.", "json_path": json_path}

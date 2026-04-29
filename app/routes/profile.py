"""
Rotas de perfil do usuario.
"""

from typing import Any

import requests
from fastapi import APIRouter, Request, Response

from app.config import GRAPH_BASE_URL
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


@router.get("/photo")
def get_profile_photo(request: Request) -> Response:
    """Retorna foto de perfil do usuario autenticado, quando disponivel."""
    access_token = get_local_access_token(request)
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/photos/120x120/$value",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )

    if response.status_code == 404:
        return Response(status_code=404)

    if response.status_code >= 400:
        return Response(status_code=response.status_code)

    content_type = response.headers.get("Content-Type", "image/jpeg")
    return Response(content=response.content, media_type=content_type)

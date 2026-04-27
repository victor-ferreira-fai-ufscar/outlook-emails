"""
Rota de health check da aplicacao.
"""

from fastapi import APIRouter

router = APIRouter()


@router.get("/")
def root() -> dict[str, str]:
    """Status da aplicacao."""
    return {
        "message": "Outlook integration is running.",
        "next_step": "Open /auth/login to authenticate your account.",
    }

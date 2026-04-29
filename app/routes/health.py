"""
Rota de health check da aplicacao.
"""

from fastapi import APIRouter

router = APIRouter()


@router.get("/")
def root() -> dict:
    """Status da aplicacao com mini tutorial de onboarding."""
    return {
        "message": "Outlook integration is running.",
        "next_step": "Open /auth/login to authenticate your account.",
        "quickstart": {
            "title": "Como comecar em 5 passos",
            "steps": [
                "1) Abra /docs para ver todos os endpoints.",
                "2) No WhatsApp, envie 'login' para receber o link de autenticacao.",
                "3) Conclua o login em /auth/login?whatsapp=SEU_NUMERO.",
                "4) Volte no WhatsApp e envie 'status' para validar o vinculo.",
                "5) Envie 'resumo agora' para receber o resumo dos emails.",
            ],
        },
        "whatsapp_commands": [
            "ajuda",
            "login",
            "status",
            "perfil",
            "ultimo-email",
            "resumo agora",
        ],
        "useful_endpoints": {
            "auth_login": "/auth/login",
            "daily_summary": "/notifications/daily-summary",
            "manual_command": "/notifications/command",
            "settings_get": "/notifications/settings",
            "settings_put": "/notifications/settings",
            "whatsapp_webhook": "/whatsapp/webhook",
            "docs": "/docs",
        },
    }

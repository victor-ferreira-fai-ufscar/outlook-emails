"""Rotas base do bot para onboarding e comandos iniciais no Teams."""

from fastapi import APIRouter
from pydantic import BaseModel

from app.config import BOT_LOGIN_URL

router = APIRouter(prefix="/bot", tags=["bot"])


class BotActivity(BaseModel):
    """Representa atividade mínima recebida do canal de chat."""

    type: str = "message"
    text: str = ""


@router.get("/health")
def bot_health() -> dict:
    """Healthcheck do módulo de bot."""
    return {"status": "ok", "channel": "microsoft-teams"}


@router.post("/messages")
def bot_messages(activity: BotActivity) -> dict:
    """Processa comandos textuais básicos para bootstrap do bot."""
    command = (activity.text or "").strip().lower()

    if command in {"ajuda", "help", "/help"}:
        return {
            "type": "help",
            "message": "Comandos disponíveis: ajuda, login, status, logout",
            "commands": ["ajuda", "login", "status", "logout"],
        }

    if command in {"login", "entrar"}:
        return {
            "type": "auth",
            "message": "Use o link para autenticar sua conta.",
            "login_url": BOT_LOGIN_URL,
        }

    if command == "status":
        return {
            "type": "status",
            "authenticated": False,
            "message": "Usuário ainda não autenticado neste canal.",
        }

    if command in {"logout", "sair"}:
        return {
            "type": "logout",
            "message": "Sessão local removida. Use 'login' para autenticar novamente.",
        }

    return {
        "type": "unknown",
        "message": "Comando não reconhecido. Digite 'ajuda' para ver opções.",
        "commands": ["ajuda", "login", "status", "logout"],
    }

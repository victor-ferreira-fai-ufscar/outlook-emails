"""Rotas base do bot para onboarding e comandos iniciais no Teams."""

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel

from app.config import (
    BOT_ALLOWED_CHANNEL,
    BOT_BEARER_TOKEN,
    BOT_LOGIN_URL,
    BOT_REQUIRE_AUTH,
)
from app.utils import (
    fetch_latest_sent_email,
    fetch_outlook_profile,
    get_latest_local_access_token,
)

router = APIRouter(prefix="/bot", tags=["bot"])


class BotActivity(BaseModel):
    """Representa atividade mínima recebida do canal de chat."""

    type: str = "message"
    text: str = ""
    channelId: str | None = None
    from_: dict | None = None


def _extract_bearer_token(request: Request) -> str | None:
    """Extrai token Bearer do header Authorization."""
    authorization = request.headers.get("authorization", "")
    if not authorization.lower().startswith("bearer "):
        return None
    return authorization.split(" ", 1)[1].strip()


def _enforce_webhook_auth_if_required(request: Request) -> None:
    """Valida autenticação do webhook quando habilitada no ambiente."""
    if not BOT_REQUIRE_AUTH:
        return

    incoming_token = _extract_bearer_token(request)
    if not incoming_token or not BOT_BEARER_TOKEN or incoming_token != BOT_BEARER_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized bot webhook request")


def _validate_channel(activity: BotActivity) -> None:
    """Aceita apenas o canal esperado para evitar tráfego indevido."""
    if not activity.channelId:
        return

    if activity.channelId != BOT_ALLOWED_CHANNEL:
        raise HTTPException(
            status_code=400, detail="Invalid channel for this bot endpoint"
        )


@router.get("/health")
def bot_health() -> dict:
    """Healthcheck do módulo de bot."""
    return {"status": "ok", "channel": "microsoft-teams"}


@router.post("/messages")
def bot_messages(activity: BotActivity, request: Request) -> dict:
    """Processa comandos textuais básicos para bootstrap do bot."""
    _enforce_webhook_auth_if_required(request)
    _validate_channel(activity)

    if activity.type == "conversationUpdate":
        return {
            "type": "welcome",
            "message": "Olá! Eu posso te ajudar com login e configuração. Digite 'ajuda'.",
            "commands": [
                "ajuda",
                "login",
                "status",
                "perfil",
                "ultimo-email",
                "logout",
            ],
        }

    command = (activity.text or "").strip().lower()

    if command in {"ajuda", "help", "/help"}:
        return {
            "type": "help",
            "message": "Comandos disponíveis: ajuda, login, status, perfil, ultimo-email, logout",
            "commands": [
                "ajuda",
                "login",
                "status",
                "perfil",
                "ultimo-email",
                "logout",
            ],
        }

    if command in {"login", "entrar"}:
        return {
            "type": "auth",
            "message": "Use o link para autenticar sua conta.",
            "login_url": BOT_LOGIN_URL,
        }

    if command == "status":
        try:
            access_token = get_latest_local_access_token()
            profile = fetch_outlook_profile(access_token)
        except HTTPException:
            return {
                "type": "status",
                "authenticated": False,
                "message": "Usuário ainda não autenticado neste canal.",
                "login_url": BOT_LOGIN_URL,
            }

        return {
            "type": "status",
            "authenticated": True,
            "message": "Sessão local ativa e pronta para uso.",
            "user": {
                "displayName": profile.get("displayName"),
                "mail": profile.get("mail") or profile.get("userPrincipalName"),
            },
        }

    if command in {"perfil", "profile", "me"}:
        try:
            access_token = get_latest_local_access_token()
            profile = fetch_outlook_profile(access_token)
        except HTTPException:
            return {
                "type": "auth",
                "message": "Sem sessão ativa. Use o comando 'login'.",
                "login_url": BOT_LOGIN_URL,
            }

        return {
            "type": "profile",
            "profile": profile,
        }

    if command in {"ultimo-email", "último-email", "latest-email"}:
        try:
            access_token = get_latest_local_access_token()
            latest_email = fetch_latest_sent_email(access_token)
        except HTTPException:
            return {
                "type": "auth",
                "message": "Sem sessão ativa. Use o comando 'login'.",
                "login_url": BOT_LOGIN_URL,
            }

        return {
            "type": "latest-email",
            "email": latest_email,
        }

    if command in {"logout", "sair"}:
        return {
            "type": "logout",
            "message": "Sessão local removida. Use 'login' para autenticar novamente.",
        }

    return {
        "type": "unknown",
        "message": "Comando não reconhecido. Digite 'ajuda' para ver opções.",
        "commands": ["ajuda", "login", "status", "perfil", "ultimo-email", "logout"],
    }

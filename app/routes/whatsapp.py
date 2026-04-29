"""Webhook inbound de WhatsApp via Evolution API."""

from fastapi import APIRouter, HTTPException, Request
from urllib.parse import urlencode, urlsplit, urlunsplit, parse_qsl

from app.config import (
    BOT_LOGIN_URL,
    EVOLUTION_WEBHOOK_SECRET,
    SUMMARY_TOP_N,
    SUMMARY_WINDOW_HOURS,
)
from app.supabase_client import get_user_id_by_whatsapp_number, get_user_settings
from app.utils import (
    get_access_token_for_user_id,
    fetch_latest_sent_email,
    fetch_outlook_profile,
    fetch_unread_inbox_emails,
    format_daily_email_summary,
    get_latest_local_access_token,
    normalize_whatsapp_number,
    send_whatsapp_via_evolution_api,
)

router = APIRouter(prefix="/whatsapp", tags=["whatsapp"])


def _extract_text(message: dict) -> str:
    if not isinstance(message, dict):
        return ""

    if message.get("conversation"):
        return str(message["conversation"])

    extended = message.get("extendedTextMessage") or {}
    if extended.get("text"):
        return str(extended["text"])

    image = message.get("imageMessage") or {}
    if image.get("caption"):
        return str(image["caption"])

    return ""


def _extract_sender(payload: dict) -> str:
    data = payload.get("data") or {}
    key = data.get("key") or {}
    return normalize_whatsapp_number(key.get("remoteJid", ""))


def _is_from_me(payload: dict) -> bool:
    data = payload.get("data") or {}
    key = data.get("key") or {}
    return bool(key.get("fromMe", False))


def _enforce_webhook_secret_if_configured(request: Request) -> None:
    if not EVOLUTION_WEBHOOK_SECRET:
        return

    incoming = request.headers.get("x-evolution-secret") or request.headers.get(
        "apikey"
    )
    if incoming != EVOLUTION_WEBHOOK_SECRET:
        raise HTTPException(
            status_code=401, detail="Unauthorized WhatsApp webhook request"
        )


def _build_summary_for_number(number: str) -> str:
    user_id = get_user_id_by_whatsapp_number(number)
    if not user_id:
        return (
            f"Numero ainda nao vinculado. Faça login aqui: {_build_login_url(number)}"
        )

    access_token = get_access_token_for_user_id(user_id)
    user_settings = get_user_settings(user_id)
    max_items = int(user_settings.get("max_emails_in_summary", SUMMARY_TOP_N))
    include_read = bool(user_settings.get("include_read_emails", False))
    emails = fetch_unread_inbox_emails(
        access_token=access_token,
        window_hours=SUMMARY_WINDOW_HOURS,
        max_items=max_items,
        include_read=include_read,
    )
    return format_daily_email_summary(
        emails,
        top_n=min(SUMMARY_TOP_N, max_items),
        include_read=include_read,
    )


def _build_login_url(number: str) -> str:
    """Monta link de login preservando o numero remetente para vinculo."""
    parts = urlsplit(BOT_LOGIN_URL)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["whatsapp"] = number
    return urlunsplit(
        (parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment)
    )


def _command_response(command: str, sender_number: str) -> str:
    normalized = command.strip().lower()

    if normalized in {"ajuda", "help", "/help"}:
        return "Comandos disponiveis: ajuda, login, status, perfil, ultimo-email, resumo agora"

    if normalized in {"login", "entrar"}:
        return f"Use este link para autenticar sua conta: {_build_login_url(sender_number)}"

    user_id = get_user_id_by_whatsapp_number(sender_number)
    if (
        normalized
        in {
            "status",
            "perfil",
            "profile",
            "me",
            "ultimo-email",
            "último-email",
            "latest-email",
        }
        and not user_id
    ):
        return f"Numero ainda nao vinculado. Faça login aqui: {_build_login_url(sender_number)}"

    if normalized == "status":
        try:
            access_token = get_access_token_for_user_id(user_id)
            profile = fetch_outlook_profile(access_token)
        except HTTPException:
            return f"Sem sessao ativa. Use o login: {_build_login_url(sender_number)}"

        user_email = (
            profile.get("mail") or profile.get("userPrincipalName") or "sem-email"
        )
        return f"Sessao ativa. Usuario autenticado: {user_email}"

    if normalized in {"perfil", "profile", "me"}:
        try:
            access_token = get_access_token_for_user_id(user_id)
            profile = fetch_outlook_profile(access_token)
        except HTTPException:
            return f"Sem sessao ativa. Use o login: {_build_login_url(sender_number)}"

        name = profile.get("displayName") or "Usuario"
        email = profile.get("mail") or profile.get("userPrincipalName") or "sem-email"
        return f"Perfil atual: {name} <{email}>"

    if normalized in {"ultimo-email", "último-email", "latest-email"}:
        try:
            access_token = get_access_token_for_user_id(user_id)
            latest_email = fetch_latest_sent_email(access_token)
        except HTTPException:
            return f"Sem sessao ativa. Use o login: {_build_login_url(sender_number)}"

        subject = latest_email.get("subject") or "(sem assunto)"
        return f"Ultimo email enviado: {subject}"

    if normalized in {"resumo", "resumo agora", "summary now"}:
        return "__SEND_SUMMARY__"

    return "Comando nao reconhecido. Digite 'ajuda' para ver as opcoes."


@router.post("/webhook")
def whatsapp_webhook(payload: dict, request: Request) -> dict:
    """Recebe eventos da Evolution API e responde a comandos via chat."""
    _enforce_webhook_secret_if_configured(request)

    if payload.get("event") != "messages.upsert":
        return {"status": "ok", "ignored": True, "reason": "unsupported_event"}

    if _is_from_me(payload):
        return {"status": "ok", "ignored": True, "reason": "from_me"}

    data = payload.get("data") or {}
    sender_number = _extract_sender(payload)
    message = data.get("message") or {}
    text = _extract_text(message).strip()

    if not sender_number or not text:
        return {"status": "ok", "ignored": True, "reason": "missing_sender_or_text"}

    response_text = _command_response(text, sender_number)
    if response_text == "__SEND_SUMMARY__":
        response_text = _build_summary_for_number(sender_number)

    send_whatsapp_via_evolution_api(message=response_text, number=sender_number)
    return {"status": "ok", "command": text.lower(), "recipient": sender_number}

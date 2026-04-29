"""Rotas de notificacoes diarias para WhatsApp."""

from datetime import datetime, timezone
from typing import Literal

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel, Field

from app.config import (
    NOTIFICATIONS_AUTOMATION_TOKEN,
    NOTIFICATIONS_REQUIRE_AUTH,
    SUMMARY_MAX_ITEMS,
    SUMMARY_TOP_N,
    SUMMARY_WINDOW_HOURS,
)
from app.utils import (
    fetch_unread_inbox_emails,
    format_daily_email_summary,
    get_latest_local_access_token,
    get_local_access_token,
    send_whatsapp_via_callmebot,
)
from app.supabase_client import get_user_settings, save_user_settings

router = APIRouter(prefix="/notifications", tags=["notifications"])


class NotificationSettingsUpdate(BaseModel):
    """Payload para atualizar preferencias de resumo diario."""

    max_emails_in_summary: int = Field(default=10, ge=1, le=100)
    include_read_emails: bool = False
    preferred_channel: Literal["whatsapp"] = "whatsapp"
    priority_senders: list[str] = Field(default_factory=list)


class NotificationCommandRequest(BaseModel):
    """Comandos HTTP disponiveis para acao on-demand sem Teams."""

    action: Literal["send_summary_now"]


def _session_user_id(request: Request) -> str:
    """Usa session id local como chave de configuracao para a POC."""
    return request.cookies.get("local_session_id") or "anonymous"


def _extract_bearer_token(request: Request) -> str | None:
    """Extrai token bearer opcional do cabecalho Authorization."""
    authorization = request.headers.get("authorization", "")
    if not authorization.lower().startswith("bearer "):
        return None
    return authorization.split(" ", 1)[1].strip()


def _enforce_automation_auth_if_required(request: Request) -> None:
    """Protege endpoint de automacao para uso por cron externo."""
    if not NOTIFICATIONS_REQUIRE_AUTH:
        return

    if not NOTIFICATIONS_AUTOMATION_TOKEN:
        raise HTTPException(
            status_code=500,
            detail="NOTIFICATIONS_AUTOMATION_TOKEN is required when auth is enabled.",
        )

    bearer = _extract_bearer_token(request)
    header_token = request.headers.get("x-automation-token")
    incoming = bearer or header_token

    if incoming != NOTIFICATIONS_AUTOMATION_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized automation request")


def _resolve_access_token(request: Request) -> tuple[str, str]:
    """Resolve autenticacao via token de automacao ou sessao local."""
    bearer = _extract_bearer_token(request)
    header_token = request.headers.get("x-automation-token")
    incoming = bearer or header_token

    if NOTIFICATIONS_AUTOMATION_TOKEN and incoming == NOTIFICATIONS_AUTOMATION_TOKEN:
        return get_latest_local_access_token(), "automation"

    try:
        access_token = get_local_access_token(request)
        return access_token, _session_user_id(request)
    except HTTPException:
        if NOTIFICATIONS_REQUIRE_AUTH:
            raise HTTPException(status_code=401, detail="Unauthorized request")
        return get_latest_local_access_token(), "anonymous"


def _build_daily_summary(access_token: str, user_settings: dict) -> dict:
    """Coleta emails e entrega resumo no canal configurado."""
    max_items = int(user_settings.get("max_emails_in_summary", SUMMARY_MAX_ITEMS))
    include_read = bool(user_settings.get("include_read_emails", False))

    emails = fetch_unread_inbox_emails(
        access_token=access_token,
        window_hours=SUMMARY_WINDOW_HOURS,
        max_items=max_items,
        include_read=include_read,
    )
    summary_text = format_daily_email_summary(
        emails,
        top_n=min(SUMMARY_TOP_N, max_items),
        include_read=include_read,
    )
    delivery = send_whatsapp_via_callmebot(summary_text)

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "emails_processed": len(emails),
        "summary_preview": summary_text,
        "delivery": delivery,
    }


@router.post("/daily-summary")
def send_daily_summary(request: Request) -> dict:
    """Executa coleta de emails e envia resumo diario no WhatsApp."""
    access_token, user_id = _resolve_access_token(request)
    user_settings = get_user_settings(user_id)
    result = _build_daily_summary(
        access_token=access_token, user_settings=user_settings
    )

    return {"status": "ok", "user_id": user_id, **result}


@router.post("/command")
def execute_notification_command(
    request: Request, payload: NotificationCommandRequest
) -> dict:
    """Executa comando on-demand para notificacoes sem depender de Teams."""
    access_token, user_id = _resolve_access_token(request)
    user_settings = get_user_settings(user_id)

    if payload.action == "send_summary_now":
        result = _build_daily_summary(
            access_token=access_token, user_settings=user_settings
        )
        return {"status": "ok", "action": payload.action, "result": result}

    raise HTTPException(status_code=400, detail="Unsupported command")


@router.get("/settings")
def read_notification_settings(request: Request) -> dict:
    """Retorna preferencias de notificacao do usuario autenticado."""
    _resolve_access_token(request)
    user_id = _session_user_id(request)
    settings = get_user_settings(user_id)
    return {"status": "ok", "user_id": user_id, "settings": settings}


@router.put("/settings")
def update_notification_settings(
    request: Request, payload: NotificationSettingsUpdate
) -> dict:
    """Atualiza preferencias de notificacao do usuario autenticado."""
    _resolve_access_token(request)
    user_id = _session_user_id(request)
    settings = save_user_settings(user_id=user_id, settings=payload.model_dump())
    return {"status": "ok", "user_id": user_id, "settings": settings}

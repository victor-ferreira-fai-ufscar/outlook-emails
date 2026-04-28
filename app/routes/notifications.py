"""Rotas de notificacoes diarias para WhatsApp."""

from datetime import datetime, timezone

from fastapi import APIRouter, HTTPException, Request

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
    send_whatsapp_via_callmebot,
)

router = APIRouter(prefix="/notifications", tags=["notifications"])


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


@router.post("/daily-summary")
def send_daily_summary(request: Request) -> dict:
    """Executa coleta de emails e envia resumo diario no WhatsApp."""
    _enforce_automation_auth_if_required(request)

    access_token = get_latest_local_access_token()
    emails = fetch_unread_inbox_emails(
        access_token=access_token,
        window_hours=SUMMARY_WINDOW_HOURS,
        max_items=SUMMARY_MAX_ITEMS,
    )
    summary_text = format_daily_email_summary(emails, top_n=SUMMARY_TOP_N)
    delivery = send_whatsapp_via_callmebot(summary_text)

    return {
        "status": "ok",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "emails_processed": len(emails),
        "summary_preview": summary_text,
        "delivery": delivery,
    }

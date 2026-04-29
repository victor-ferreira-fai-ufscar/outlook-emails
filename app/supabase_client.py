"""Cliente Supabase com inicializacao preguiçosa."""

from datetime import datetime, timezone
from functools import lru_cache
from typing import Any

from fastapi import HTTPException

from app.config import SUPABASE_ENABLED, SUPABASE_KEY, SUPABASE_URL

DEFAULT_USER_SETTINGS = {
    "max_emails_in_summary": 10,
    "include_read_emails": False,
    "preferred_channel": "whatsapp",
    "priority_senders": [],
}


@lru_cache(maxsize=1)
def get_supabase_client():
    """Retorna cliente Supabase quando configurado no ambiente."""
    if not SUPABASE_ENABLED:
        raise HTTPException(
            status_code=500,
            detail="Supabase não configurado. Defina SUPABASE_URL e SUPABASE_KEY.",
        )

    try:
        from supabase import create_client
    except Exception as exc:  # pragma: no cover - dependência ausente em runtime
        raise HTTPException(
            status_code=500,
            detail="Biblioteca supabase não instalada. Rode: uv add supabase",
        ) from exc

    return create_client(SUPABASE_URL, SUPABASE_KEY)


def _normalize_user_settings(raw: dict[str, Any]) -> dict[str, Any]:
    """Aplica defaults e valida campos basicos de preferencias."""
    merged = {**DEFAULT_USER_SETTINGS, **(raw or {})}

    try:
        merged["max_emails_in_summary"] = int(merged["max_emails_in_summary"])
    except (TypeError, ValueError):
        merged["max_emails_in_summary"] = DEFAULT_USER_SETTINGS["max_emails_in_summary"]

    merged["max_emails_in_summary"] = max(1, min(100, merged["max_emails_in_summary"]))
    merged["include_read_emails"] = bool(merged.get("include_read_emails", False))

    if merged.get("preferred_channel") != "whatsapp":
        merged["preferred_channel"] = "whatsapp"

    senders = merged.get("priority_senders") or []
    if not isinstance(senders, list):
        senders = []
    merged["priority_senders"] = [
        str(sender).strip().lower() for sender in senders if str(sender).strip()
    ]

    return merged


def get_user_settings(user_id: str) -> dict[str, Any]:
    """Le preferencias do usuario com fallback para defaults."""
    defaults = {"user_id": user_id, **DEFAULT_USER_SETTINGS}

    if not SUPABASE_ENABLED:
        return defaults

    try:
        client = get_supabase_client()
        response = (
            client.table("user_settings")
            .select(
                "user_id,max_emails_in_summary,include_read_emails,preferred_channel,priority_senders"
            )
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )
        rows = response.data or []
        if not rows:
            return defaults

        normalized = _normalize_user_settings(rows[0])
        return {"user_id": user_id, **normalized}
    except Exception:
        return defaults


def save_user_settings(user_id: str, settings: dict[str, Any]) -> dict[str, Any]:
    """Salva preferencias do usuario; nao quebra fluxo se Supabase falhar."""
    normalized = {"user_id": user_id, **_normalize_user_settings(settings)}

    if not SUPABASE_ENABLED:
        return normalized

    try:
        client = get_supabase_client()
        client.table("user_settings").upsert(
            {
                **normalized,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            },
            on_conflict="user_id",
        ).execute()
    except Exception:
        return normalized

    return normalized

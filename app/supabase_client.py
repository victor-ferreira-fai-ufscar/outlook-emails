"""Cliente Supabase com inicializacao preguiçosa."""

from datetime import datetime, timezone
from functools import lru_cache
import json
from pathlib import Path
from typing import Any

from fastapi import HTTPException

from app.config import SUPABASE_ENABLED, SUPABASE_KEY, SUPABASE_URL

DEFAULT_USER_SETTINGS = {
    "max_emails_in_summary": 10,
    "include_read_emails": False,
    "preferred_channel": "whatsapp",
    "priority_senders": [],
}


def _normalize_whatsapp_number(number: str) -> str:
    """Normaliza numero/JID para apenas digitos."""
    if not number:
        return ""
    normalized = str(number).split("@", 1)[0]
    return "".join(ch for ch in normalized if ch.isdigit())


def _whatsapp_links_file() -> Path:
    """Arquivo local de fallback para vinculos de WhatsApp."""
    directory = Path("sessions")
    directory.mkdir(parents=True, exist_ok=True)
    return directory / "whatsapp-links.json"


def _read_local_whatsapp_links() -> dict[str, Any]:
    """Le cache local de vinculos de WhatsApp."""
    file_path = _whatsapp_links_file()
    if not file_path.exists():
        return {"numbers": {}, "users": {}}

    try:
        with file_path.open("r", encoding="utf-8") as fp:
            data = json.load(fp)
    except Exception:
        return {"numbers": {}, "users": {}}

    return {
        "numbers": data.get("numbers", {}),
        "users": data.get("users", {}),
    }


def _write_local_whatsapp_links(payload: dict[str, Any]) -> None:
    """Escreve cache local de vinculos de WhatsApp."""
    with _whatsapp_links_file().open("w", encoding="utf-8") as fp:
        json.dump(payload, fp, indent=2, ensure_ascii=False)


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


def link_whatsapp_number_to_user(user_id: str, whatsapp_number: str) -> dict[str, str]:
    """Vincula numero de WhatsApp ao usuario autenticado."""
    normalized_number = _normalize_whatsapp_number(whatsapp_number)
    record = {
        "user_id": user_id,
        "whatsapp_number": normalized_number,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    local_links = _read_local_whatsapp_links()
    local_links["numbers"][normalized_number] = user_id
    local_links["users"][user_id] = normalized_number
    _write_local_whatsapp_links(local_links)

    if SUPABASE_ENABLED:
        try:
            client = get_supabase_client()
            client.table("whatsapp_links").upsert(
                record, on_conflict="whatsapp_number"
            ).execute()
        except Exception:
            pass

    return record


def get_user_id_by_whatsapp_number(whatsapp_number: str) -> str | None:
    """Resolve usuario pelo numero de WhatsApp vinculado."""
    normalized_number = _normalize_whatsapp_number(whatsapp_number)
    local_links = _read_local_whatsapp_links()
    local_hit = local_links["numbers"].get(normalized_number)
    if local_hit:
        return local_hit

    if not SUPABASE_ENABLED:
        return None

    try:
        client = get_supabase_client()
        response = (
            client.table("whatsapp_links")
            .select("user_id")
            .eq("whatsapp_number", normalized_number)
            .limit(1)
            .execute()
        )
        rows = response.data or []
        if not rows:
            return None
        user_id = rows[0].get("user_id")
        if user_id:
            local_links["numbers"][normalized_number] = user_id
            local_links["users"][user_id] = normalized_number
            _write_local_whatsapp_links(local_links)
        return user_id
    except Exception:
        return None


def get_whatsapp_number_by_user_id(user_id: str) -> str | None:
    """Resolve numero de WhatsApp pelo usuario autenticado."""
    local_links = _read_local_whatsapp_links()
    local_hit = local_links["users"].get(user_id)
    if local_hit:
        return local_hit

    if not SUPABASE_ENABLED:
        return None

    try:
        client = get_supabase_client()
        response = (
            client.table("whatsapp_links")
            .select("whatsapp_number")
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )
        rows = response.data or []
        if not rows:
            return None
        number = rows[0].get("whatsapp_number")
        if number:
            local_links["numbers"][number] = user_id
            local_links["users"][user_id] = number
            _write_local_whatsapp_links(local_links)
        return number
    except Exception:
        return None

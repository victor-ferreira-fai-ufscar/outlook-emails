"""Cliente Supabase com inicializacao preguiçosa."""

from functools import lru_cache

from fastapi import HTTPException

from app.config import SUPABASE_ENABLED, SUPABASE_KEY, SUPABASE_URL


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

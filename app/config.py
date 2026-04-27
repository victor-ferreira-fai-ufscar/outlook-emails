"""
Configuracoes e constantes globais da aplicacao.
"""

import os
from dotenv import load_dotenv

load_dotenv()


def _as_bool(value: str, default: bool = False) -> bool:
    """Converte variável de ambiente textual para booleano."""
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["User.Read", "Mail.Read"]

MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "common")
MS_REDIRECT_URI = os.getenv("MS_REDIRECT_URI", "http://localhost:8000/auth/callback")
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY", "change-this-in-production")
BOT_LOGIN_URL = os.getenv("BOT_LOGIN_URL", "http://localhost:8000/auth/login")
BOT_REQUIRE_AUTH = _as_bool(os.getenv("BOT_REQUIRE_AUTH", "false"), default=False)
BOT_BEARER_TOKEN = os.getenv("BOT_BEARER_TOKEN", "")
BOT_ALLOWED_CHANNEL = os.getenv("BOT_ALLOWED_CHANNEL", "msteams")

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
SUPABASE_ENABLED = bool(SUPABASE_URL and SUPABASE_KEY)

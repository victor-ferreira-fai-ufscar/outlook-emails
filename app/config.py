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


def _as_int(value: str | None, default: int) -> int:
    """Converte variável de ambiente textual para inteiro com fallback."""
    if value is None or value.strip() == "":
        return default
    try:
        return int(value)
    except ValueError:
        return default


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

EVOLUTION_API_URL = os.getenv("EVOLUTION_API_URL", "")
EVOLUTION_API_KEY = os.getenv("EVOLUTION_API_KEY", "")
EVOLUTION_INSTANCE = os.getenv("EVOLUTION_INSTANCE", "")
EVOLUTION_DEFAULT_NUMBER = os.getenv("EVOLUTION_DEFAULT_NUMBER", "")
EVOLUTION_TIMEOUT_SECONDS = _as_int(os.getenv("EVOLUTION_TIMEOUT_SECONDS"), 20)
EVOLUTION_WEBHOOK_SECRET = os.getenv("EVOLUTION_WEBHOOK_SECRET", "")
WHATSAPP_ALLOW_FROM_ME = _as_bool(
    os.getenv("WHATSAPP_ALLOW_FROM_ME", "false"), default=False
)
WHATSAPP_ALLOWED_GROUP_ID = os.getenv("WHATSAPP_ALLOWED_GROUP_ID", "").split("@")[0].strip()

NOTIFICATIONS_REQUIRE_AUTH = _as_bool(
    os.getenv("NOTIFICATIONS_REQUIRE_AUTH", "true"), default=True
)
NOTIFICATIONS_AUTOMATION_TOKEN = os.getenv("NOTIFICATIONS_AUTOMATION_TOKEN", "")

SUMMARY_WINDOW_HOURS = _as_int(os.getenv("SUMMARY_WINDOW_HOURS"), 24)
SUMMARY_MAX_ITEMS = _as_int(os.getenv("SUMMARY_MAX_ITEMS"), 20)
SUMMARY_TOP_N = _as_int(os.getenv("SUMMARY_TOP_N"), 10)

SUMMARY_PRIORITY_SENDERS = [
    sender.strip().lower()
    for sender in os.getenv("SUMMARY_PRIORITY_SENDERS", "").split(",")
    if sender.strip()
]

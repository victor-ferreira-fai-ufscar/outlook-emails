"""
Configuracoes e constantes globais da aplicacao.
"""

import os
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["User.Read", "Mail.Read"]

MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "common")
MS_REDIRECT_URI = os.getenv("MS_REDIRECT_URI", "http://localhost:8000/auth/callback")
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY", "change-this-in-production")

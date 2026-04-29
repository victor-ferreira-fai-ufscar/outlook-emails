"""
Aplicacao FastAPI para integracao com Outlook via Microsoft Graph.

Estrutura:
- app/config.py: configuracoes globais
- app/utils.py: funcoes auxiliares (Graph API, sessoes, MSAL)
- app/routes/: rotas da aplicacao
  - health.py: status da app
  - auth.py: autenticacao OAuth2
  - profile.py: perfil do usuario
  - messages.py: mensagens (emails)
"""

from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from scalar_fastapi import get_scalar_api_reference

from app.routes import health, auth, profile, messages, bot, notifications, whatsapp

app = FastAPI(
    title="Outlook Profile Integration",
    description="MVP para integrar Outlook via Microsoft Graph e FastAPI",
    version="0.1.0",
)


@app.get("/scalar", include_in_schema=False)
def scalar_docs() -> HTMLResponse:
    openapi_url = app.openapi_url or "/openapi.json"
    return get_scalar_api_reference(
        openapi_url=openapi_url,
        title=f"{app.title} - Scalar",
        # Use Scalar proxy to prevent CORS issues when loading OpenAPI from browsers.
        scalar_proxy_url="https://proxy.scalar.com",
    )


# Registrar rotas
app.include_router(health.router)
app.include_router(auth.router)
app.include_router(auth.public_router)
app.include_router(profile.router)
app.include_router(messages.router)
app.include_router(bot.router)
app.include_router(notifications.router)
app.include_router(whatsapp.router)

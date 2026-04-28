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

from app.routes import health, auth, profile, messages, bot, notifications

app = FastAPI(
    title="Outlook Profile Integration",
    description="MVP para integrar Outlook via Microsoft Graph e FastAPI",
    version="0.1.0",
)


@app.get("/scalar", include_in_schema=False)
def scalar_docs() -> HTMLResponse:
    openapi_url = app.openapi_url or "/openapi.json"
    return HTMLResponse(
        f"""
<!doctype html>
<html>
  <head>
    <meta charset=\"utf-8\" />
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
    <title>{app.title} - Scalar</title>
  </head>
  <body>
    <script id=\"api-reference\" data-url=\"{openapi_url}\"></script>
    <script src=\"https://cdn.jsdelivr.net/npm/@scalar/api-reference\"></script>
  </body>
</html>
""".strip()
    )


# Registrar rotas
app.include_router(health.router)
app.include_router(auth.router)
app.include_router(auth.public_router)
app.include_router(profile.router)
app.include_router(messages.router)
app.include_router(bot.router)
app.include_router(notifications.router)

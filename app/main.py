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

from app.routes import health, auth, profile, messages

app = FastAPI(
    title="Outlook Profile Integration",
    description="MVP para integrar Outlook via Microsoft Graph e FastAPI",
    version="0.1.0",
)

# Registrar rotas
app.include_router(health.router)
app.include_router(auth.router)
app.include_router(profile.router)
app.include_router(messages.router)

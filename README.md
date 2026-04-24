# Outlook Resume Emails

MVP para integrar uma conta Outlook com FastAPI, autenticar via Microsoft Graph e gerar um JSON com os dados de perfil.

## Stack

- Python + FastAPI
- `uv` para gerenciar dependencias
- Microsoft Graph API (OAuth2)

## Pre-requisitos

- Conta Microsoft 365
- Aplicativo registrado no Azure Entra ID (Azure AD)
- Redirect URI do app: `http://localhost:8000/auth/callback`

Permissoes minimas no Graph:

- `User.Read`

## Configuracao

1. Copie o arquivo de exemplo:

```bash
cp .env.example .env
```

1. Preencha as variaveis no `.env`:

- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_TENANT_ID` (pode ser `common` para testes)
- `MS_REDIRECT_URI`
- `SESSION_SECRET_KEY`

## Executar (recomendado: Docker Compose)

```bash
docker compose up --build
```

Aplicacao em: `http://localhost:8000`

## Executar local com uv

```bash
uv sync
uv run uvicorn app.main:app --reload
```

## Fluxo de uso

1. Acesse `http://localhost:8000/auth/login`
2. Faça login na conta Outlook
3. A API retorna:
   - `profile` com dados de usuario
   - `json_path` com o arquivo salvo em `data/`

Endpoints disponiveis:

- `GET /` status da aplicacao
- `GET /auth/login` inicia OAuth2
- `GET /auth/callback` processa callback do OAuth2
- `GET /profile` consulta perfil atual
- `GET /profile/export` exporta novo JSON do perfil

## Estrutura

```text
app/
  main.py
data/
docs/
pyproject.toml
docker-compose.yml
```

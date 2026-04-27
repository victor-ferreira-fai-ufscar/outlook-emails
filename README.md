# Outlook Resume Emails

MVP para integrar uma conta Outlook com FastAPI, autenticar via Microsoft Graph e gerar um JSON com os dados de perfil.

No estado atual, o fluxo OAuth e token de acesso ficam persistidos localmente em arquivos na pasta `sessions/`.

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
- `Mail.Read`

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
- `BOT_LOGIN_URL`
- `BOT_REQUIRE_AUTH` (`true/false`)
- `BOT_BEARER_TOKEN` (token Bearer do webhook do bot)
- `BOT_ALLOWED_CHANNEL` (padrao: `msteams`)
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
- `GET /messages/sent/latest` retorna o ultimo e-mail enviado
- `GET /bot/health` status do modulo de bot (Teams)
- `POST /bot/messages` webhook inicial para comandos (`ajuda`, `login`, `status`, `logout`)

## Bot no Teams (fase inicial)

Esta base ja inclui um modulo inicial para bot em FastAPI, preparado para evoluir para o fluxo Teams sem criar outro backend.

Comandos iniciais suportados no webhook:

- `ajuda`
- `login`
- `status`
- `logout`

Eventos aceitos no webhook:

- `message`
- `conversationUpdate` (retorna mensagem de boas-vindas)

Seguranca do webhook:

- Se `BOT_REQUIRE_AUTH=true`, o endpoint exige `Authorization: Bearer <BOT_BEARER_TOKEN>`.
- Se `channelId` vier no payload e for diferente de `BOT_ALLOWED_CHANNEL`, a requisicao e rejeitada.

Proximo passo recomendado: conectar esse endpoint ao Bot Framework e persistir estado de conversa em banco (Supabase Postgres) ou cache dedicado.

## Deploy 24/7 (Render)

Para deploy no Render, publique este backend FastAPI e configure as variaveis do `.env.example` no painel de ambiente.

Este repositorio ja inclui blueprint em `render.yaml` para facilitar o provisionamento.

Comando de start sugerido:

```bash
uv run uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

Passos sugeridos no Render:

1. Criar novo Web Service a partir do repositorio.
2. Selecionar o arquivo `render.yaml`.
3. Definir segredos (`MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_REDIRECT_URI`, `BOT_LOGIN_URL`, `BOT_BEARER_TOKEN`, `SESSION_SECRET_KEY`).
4. Validar healthcheck em `/bot/health`.

## Persistencia local temporaria

- `sessions/flow-<state>.json`: dados do auth flow durante login
- `sessions/session-<uuid>.json`: token de acesso e metadados da sessao autenticada
- `data/outlook-profile-*.json`: snapshots do perfil exportado

## Estrutura

```text
app/
  main.py
data/
docs/
pyproject.toml
docker-compose.yml
```

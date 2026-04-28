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
- `SUPABASE_URL` (opcional)
- `SUPABASE_KEY` (opcional)
- `SESSION_SECRET_KEY`

### Redirect URI por ambiente

Use valores diferentes para local e producao:

- **Local**
  - `MS_REDIRECT_URI=http://localhost:8000/auth/callback`
- **Render (producao)**
  - `MS_REDIRECT_URI=https://outlook-emails.onrender.com/auth/callback`

Observacao: o backend agora aceita tambem `https://outlook-emails.onrender.com/callback` para compatibilidade com registros existentes no Azure.

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
- `GET /docs` documentacao Swagger UI
- `GET /scalar` documentacao Scalar API Reference
- `GET /openapi.json` schema OpenAPI consumido por Swagger/Scalar
- `GET /auth/login` inicia OAuth2
- `GET /auth/callback` processa callback do OAuth2
- `GET /callback` alias de callback para compatibilidade em producao
- `GET /profile` consulta perfil atual
- `GET /profile/export` exporta novo JSON do perfil
- `GET /messages/sent/latest` retorna o ultimo e-mail enviado
- `POST /notifications/daily-summary` gera resumo diario de emails nao lidos e envia no WhatsApp (CallMeBot)
- `GET /bot/health` status do modulo de bot (Teams)
- `POST /bot/messages` webhook inicial para comandos (`ajuda`, `login`, `status`, `logout`)

## Resumo diario no WhatsApp (MVP)

Fluxo implementado:

1. Le e-mails nao lidos das ultimas 24h no Outlook (Graph API)
2. Classifica prioridade por regras simples (urgente/media/baixa)
3. Gera resumo textual
4. Envia no WhatsApp via CallMeBot

Variaveis necessarias no `.env`:

- `CALLMEBOT_PHONE`
- `CALLMEBOT_API_KEY`
- `NOTIFICATIONS_AUTOMATION_TOKEN`

Opcional:

- `SUMMARY_PRIORITY_SENDERS` (lista CSV de remetentes com prioridade alta)

Exemplo de disparo manual (apos autenticar em `/auth/login`):

```bash
curl -X POST http://localhost:8000/notifications/daily-summary \
  -H "Authorization: Bearer ${NOTIFICATIONS_AUTOMATION_TOKEN}"
```

## Agendamento diario com GitHub Actions

Workflow adicionado em `.github/workflows/daily-summary.yml` com:

- `schedule` (cron diario)
- `workflow_dispatch` (execucao manual)

Secrets esperados no GitHub:

- `DAILY_SUMMARY_URL` (ex.: `https://seu-backend.onrender.com/notifications/daily-summary`)
- `NOTIFICATIONS_AUTOMATION_TOKEN`

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

## Supabase (opcional, recomendado para producao)

Quando `SUPABASE_URL` e `SUPABASE_KEY` estao definidos, a aplicacao continua gravando localmente e tambem sincroniza dados para Supabase.

Comportamento atual:

- `sessions/` continua sendo escrita localmente.
- Tabela `sessions` no Supabase recebe upsert por `file_name`.
- Tabela `profiles` no Supabase recebe snapshots exportados.
- Se o arquivo local de sessao nao existir, o backend tenta ler da tabela `sessions`.

Estrutura minima esperada no Supabase:

- tabela `sessions`: `file_name` (text, unique), `payload` (jsonb), `updated_at` (timestamptz)
- tabela `profiles`: `id` (uuid default), `user_id` (text), `path` (text), `payload` (jsonb), `created_at` (timestamptz)

Arquivo pronto com SQL:

- `docs/supabase.sql`

## Estrutura

```text
app/
  main.py
data/
docs/
pyproject.toml
docker-compose.yml
```

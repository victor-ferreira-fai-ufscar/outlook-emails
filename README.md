# Outlook Resume Emails

MVP para integrar uma conta Outlook com FastAPI, autenticar via Microsoft Graph e gerar um JSON com os dados de perfil.

No estado atual, o fluxo OAuth e token de acesso ficam persistidos localmente em arquivos na pasta `sessions/`.

## Inicio rapido para desenvolvimento (principal)

Se o objetivo e desenvolver localmente com tranquilidade, este e o fluxo recomendado.

1- Copie e configure o ambiente:

```bash
cp .env.example .env
```

2- Suba os containers de desenvolvimento:

```bash
docker compose -f docker-compose.dev.yml up -d --build
```

3- Verifique se os servicos subiram:

```bash
docker compose -f docker-compose.dev.yml ps
```

4- Acesse:

- API: `http://localhost:8000`
- Evolution API: `http://localhost:8080`
- Evolution Manager: `http://localhost:3000`

5- Quando terminar:

```bash
docker compose -f docker-compose.dev.yml down --remove-orphans
```

6- Para reiniciar so a API (apos mudar variaveis no `.env`, por exemplo):

```bash
docker compose -f docker-compose.dev.yml stop api
docker compose -f docker-compose.dev.yml up -d --build api
```

Observacao: `docker-compose.yml` continua como alias de desenvolvimento para compatibilidade, mas para evitar duvidas prefira usar explicitamente `docker-compose.dev.yml`.

## Stack

- Python + FastAPI
- `uv` para gerenciar dependencias
- Microsoft Graph API (OAuth2)
- Jinja2 para templates de mensagens (WhatsApp)

## IA Local com Ollama

Este projeto utiliza o Ollama para rodar modelos de linguagem localmente, garantindo privacidade e permitindo a sumarização de e-mails. O modelo padrão utilizado é o `gemma3:1b`.

### Instalação do Ollama

1. Acesse o site oficial do Ollama: [https://ollama.com/download](https://ollama.com/download)
2. Baixe e instale a versão correspondente ao seu sistema operacional (Windows, macOS ou Linux).
3. Após a instalação, verifique se o Ollama está em execução abrindo um terminal e digitando:

```bash
ollama --version
```

### Baixando o modelo gemma3:1b

Com o Ollama instalado, você precisa baixar o modelo utilizado pelo projeto. No terminal, execute o comando abaixo:

```bash
ollama run gemma3:1b
```

Este comando fará o download do modelo e o inicializará no seu ambiente local. Após o download concluir e o prompt do modelo aparecer, você pode sair digitando `/bye` ou pressionando `Ctrl+d`. O Ollama manterá o modelo disponível para uso da API.

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

## Executar

### Docker Compose (RECOMENDADO)

A forma recomendada de executar e via Docker, garantindo isolamento, reproducibilidade e sem dependencias locais.

Este repositorio agora possui 2 perfis separados:

- Desenvolvimento: `docker-compose.dev.yml`
- Producao: `docker-compose.prod.yml`

Tambem mantemos `docker-compose.yml` como alias de desenvolvimento para compatibilidade.

### Rodar em desenvolvimento

Use este modo no dia a dia. Ele sobe API com hot reload e monta codigo local via volume.

```bash
docker compose -f docker-compose.dev.yml up -d --build
docker compose -f docker-compose.dev.yml ps
```

Para parar:

```bash
docker compose -f docker-compose.dev.yml down --remove-orphans
```

### Rodar em producao

Use este modo para ambiente estavel (sem reload, com volumes nomeados para dados/sessoes).

```bash
docker compose -f docker-compose.prod.yml up -d --build
docker compose -f docker-compose.prod.yml ps
```

Para parar:

```bash
docker compose -f docker-compose.prod.yml down --remove-orphans
```

Para reiniciar so a API em producao:

```bash
docker compose -f docker-compose.prod.yml stop api
docker compose -f docker-compose.prod.yml up -d --build api
```

Aplicacao em: `http://localhost:8000`

Servicos adicionais desta stack:

- Backend FastAPI: `http://localhost:8000`
- Evolution API: `http://localhost:8080`
- Evolution Manager (somente dev): `http://localhost:3000`

Beneficios:

- Ambiente isolado e reproduzivel
- Sem necessidade de instalar dependencias localmente
- Funciona igual em qualquer maquina (local, CI/CD, producao)
- Fluxo separado para dev e producao

### Local com uv (Alternativa)

Para desenvolvimento sem container, você pode rodar localmente:

```bash
uv sync
uv run uvicorn app.main:app --reload
```

Requer Python 3.12+ e `uv` instalados localmente.

## Fluxo de uso

1. Acesse `http://localhost:8000/auth/login`
2. Faça login na conta Outlook
3. A API retorna:
   - `profile` com dados de usuario
   - `json_path` com o arquivo salvo em `data/`

Endpoints disponiveis:

- `GET /` status da aplicacao
- `GET /docs` documentacao Swagger UI
- `GET /scalar` documentacao Scalar API Reference (via `scalar-fastapi`, com proxy para evitar problemas de CORS)
- `GET /openapi.json` schema OpenAPI consumido por Swagger/Scalar
- `GET /auth/login` inicia OAuth2
- `GET /auth/callback` processa callback do OAuth2
- `GET /callback` alias de callback para compatibilidade em producao
- `GET /profile` consulta perfil atual
- `GET /profile/export` exporta novo JSON do perfil
- `GET /messages/sent/latest` retorna o ultimo e-mail enviado
- `POST /notifications/daily-summary` gera resumo diario de emails e envia no WhatsApp via Evolution API
- `POST /notifications/command` executa comando on-demand (`send_summary_now`)
- `GET /notifications/settings` consulta preferencias de notificacao
- `PUT /notifications/settings` atualiza preferencias de notificacao
- `POST /whatsapp/webhook` recebe comandos inbound do WhatsApp via Evolution API
- `GET /bot/health` status do modulo de bot (Teams)
- `POST /bot/messages` webhook inicial para comandos (`ajuda`, `login`, `status`, `logout`)

## Resumo diario no WhatsApp (MVP)

Fluxo implementado:

1. Le e-mails nao lidos das ultimas 24h no Outlook (Graph API)
2. Classifica prioridade por regras simples (urgente/media/baixa)
3. Gera resumo textual
4. Envia no WhatsApp via Evolution API
5. **Agendamento Automático**: O sistema possui um scheduler interno (`APScheduler`) que envia resumos diários no horário configurado por cada usuário (padrão 08:00).

Variaveis necessarias no `.env`:

- `EVOLUTION_API_URL`
- `EVOLUTION_API_KEY`
- `EVOLUTION_INSTANCE`
- `EVOLUTION_DEFAULT_NUMBER`
- `NOTIFICATIONS_AUTOMATION_TOKEN`

Opcional:

- `SUMMARY_PRIORITY_SENDERS` (lista CSV de remetentes com prioridade alta)
- `WHATSAPP_ALLOWED_GROUP_ID` (ID do grupo permitido, opcional)
- `WHATSAPP_ALLOWED_NUMBERS` (lista CSV de números individuais permitidos, opcional)

Exemplo de disparo manual (apos autenticar em `/auth/login`):

```bash
curl -X POST http://localhost:8000/notifications/daily-summary \
  -H "Authorization: Bearer ${NOTIFICATIONS_AUTOMATION_TOKEN}"
```

## Rodando a Evolution API localmente

Esta stack ja inclui Evolution API, Redis, Postgres e Manager web (somente dev) nos arquivos de compose.

### 1. Preparar o `.env`

Defina pelo menos:

- `EVOLUTION_API_KEY`
- `EVOLUTION_INSTANCE`
- `EVOLUTION_DEFAULT_NUMBER`
- `BOT_LOGIN_URL`

Para ambiente local, estes valores funcionam bem:

```env
BOT_LOGIN_URL=http://localhost:8000/auth/login
EVOLUTION_API_URL=http://localhost:8080
EVOLUTION_API_KEY=troque-por-uma-chave-forte
EVOLUTION_INSTANCE=outlook-emails
EVOLUTION_DEFAULT_NUMBER=5511999999999
```

### 2. Subir os containers

Ambiente de desenvolvimento:

```bash
docker compose -f docker-compose.dev.yml up -d --build api evolution-postgres evolution-redis evolution-api evolution-manager
```

Ambiente de producao:

```bash
docker compose -f docker-compose.prod.yml up -d --build api evolution-postgres evolution-redis evolution-api
```

### 3. Abrir o painel da Evolution

Abra:

- `http://localhost:3000`

Crie ou localize a instância com o mesmo nome configurado em `EVOLUTION_INSTANCE`.
Neste projeto, o valor esperado e `outlook-emails`.

### 4. Conectar o WhatsApp

No Evolution Manager:

1. Crie a instância `outlook-emails`, se ainda nao existir.
2. Gere o QR Code.
3. Escaneie com o WhatsApp que sera usado pelo bot.

### 5. Validar a API da Evolution

Teste o endpoint de envio direto:

```bash
curl -X POST http://localhost:8080/message/sendText/outlook-emails \
  -H "Content-Type: application/json" \
  -H "apikey: ${EVOLUTION_API_KEY}" \
  -d '{"number": "5511999999999", "text": "teste evolution"}'
```

Se isso responder `201`, a Evolution esta pronta.

### 6. Webhook inbound ja configurado

Os arquivos de compose desta stack ja sobem a Evolution com webhook global apontando para:

- `http://api:8000/whatsapp/webhook`

Ou seja: quando a instância estiver conectada, mensagens recebidas no WhatsApp ja chegam automaticamente ao backend.

## Vinculo entre WhatsApp e login Outlook

Agora o fluxo ficou assim:

1. O usuario manda `login` no WhatsApp.
2. O bot responde com um link como `/auth/login?whatsapp=5511...`.
3. O usuario autentica no Outlook.
4. No callback OAuth, o backend vincula esse numero do WhatsApp ao `user_id` do Outlook.
5. A partir desse momento, comandos como `status`, `perfil`, `ultimo-email` e `resumo` usam a sessao do usuario autenticado correto.

Comandos suportados no chat (requerem prefixo `!` por padrão):

- `!ajuda`: Lista todos os comandos
- `!login`: Gera link de autenticação vinculado ao seu WhatsApp
- `!status`: Verifica se sua conta Microsoft está conectada
- `!perfil`: Exibe nome e e-mail da conta vinculada
- `!ultimo-email`: Mostra o assunto do último e-mail enviado
- `!resumo`: Gera um resumo imediato das últimas 24h e envia no chat
- `!agendar HH:MM`: Define o horário do resumo automático diário
- `!config`: Exibe suas configurações atuais (agendamento, limite, etc)

Exemplo de comando HTTP on-demand:

```bash
curl -X POST http://localhost:8000/notifications/command \
  -H "Content-Type: application/json" \
  --cookie "local_session_id=<seu-session-id>" \
  -d '{"action": "send_summary_now"}'
```

## Comandos pelo WhatsApp com Evolution API

Fluxo implementado:

1. A Evolution API entrega eventos `messages.upsert` para `POST /whatsapp/webhook`
2. O backend extrai o texto recebido e identifica o comando
3. O backend responde para o mesmo numero via `POST /message/sendText/{instance}`

Comandos suportados no chat (exigem prefixo):

- `!ajuda`
- `!login`
- `!status`
- `!perfil`
- `!ultimo-email`
- `!resumo`
- `!agendar HH:MM`
- `!config`

Para habilitar o inbound, configure na Evolution API o webhook do evento `messages.upsert` apontando para a URL deste backend.

## Agendamento diário

O sistema possui um **scheduler interno** (`APScheduler`) que gerencia o envio automático dos resumos. Por padrão, os resumos são enviados às 08:00, mas cada usuário pode personalizar seu horário via comando `!agendar`.

### Alternativa: GitHub Actions

Caso prefira um gatilho externo, o workflow em `.github/workflows/daily-summary.yml` continua disponível com:

Workflow adicionado em `.github/workflows/daily-summary.yml` com:

- `schedule` (cron diario)
- `workflow_dispatch` (execucao manual)

Secrets esperados no GitHub:

- `DAILY_SUMMARY_URL` (ex.: `https://seu-backend.onrender.com/notifications/daily-summary`)
- `NOTIFICATIONS_AUTOMATION_TOKEN`

## Integração com Microsoft Teams (Adaptive Cards)

O fluxo de interação pelo Microsoft Teams e WhatsApp permite a configuração ágil do assistente, suportando inclusive interfaces ricas (Adaptive Cards no Teams). 

O fluxo implementado prevê:
1. Autenticação unificada vinculando sua conta de chat ao Outlook.
2. Configuração rápida via painéis/comandos interativos.
3. Agendamentos dinâmicos (Resumo diário, X dias atrás, semanal).
4. Execução sob demanda (`!resumo hoje`, `!resumo ontem`).

Consulte o detalhamento do fluxo em `docs/email-summarizer-flow.md`.

A segurança base do webhook do Bot inclui:
- `BOT_REQUIRE_AUTH=true` exige cabeçalho `Authorization: Bearer <BOT_BEARER_TOKEN>`.
- Validação estrita do canal configurado em `BOT_ALLOWED_CHANNEL`.

## Segurança e Privacidade (Privacy-First)

Este projeto implementa um modelo rigoroso de privacidade (Zero Retention Data), garantindo que dados sensíveis não fiquem expostos ou persistidos de forma vulnerável:

- **Criptografia em Repouso (AES-256)**: Os tokens de acesso e de atualização do Microsoft Graph são fortemente criptografados no banco de dados. Apenas o sistema de backend possui a chave (`SESSION_SECRET_KEY`) para operá-los em tempo de execução.
- **Processamento Efêmero na RAM**: Todo o conteúdo dos e-mails baixados e os respectivos resumos da IA existem **apenas na memória volátil**. O sistema atua como um "quadro branco", apagando todo o contexto imediatamente após enviar a notificação no chat. Nada de dados crus no BD.

Saiba mais como protegemos seus dados lendo nossa [Explicação de Segurança e Privacidade](docs/backend-security-flow.md).

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
  scheduler.py
  worker.py
  supabase_client.py
  utils.py
  routes/
  templates/
data/
docs/
pyproject.toml
docker-compose.dev.yml
docker-compose.prod.yml
```

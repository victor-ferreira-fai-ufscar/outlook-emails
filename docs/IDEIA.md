# Ideia

## Tech Stack

- Back-end:
  - <https://github.com/astral-sh/uv>
  - <https://fastapi.tiangolo.com/>

## Minha escrita/rascunho

```markdown
- [ ] Apresentar Proposta de Resumo de E-mail para Roziane (Gestora projetos)

Roziane - Emails

- Outlook
- Acessa o conjunto dos e-mails e faça um relatório/resumo do email
- Ela recebe muitos e-mails é bom resumir
- Verificar se o Copilot
- Enviar notificações no Teams resumido dos e-mails do dia anterior
- Categorizar por ordem de importância (urgente, médio, baixo)

estava pensando em usar a API do Outlook para buscar esses emails e resumir e depois enviar em notificação pelo Teams
```

## Output

Essa é uma excelente iniciativa para demonstrar o valor da integração de IA na FAI•UFSCar. Automatizar esse fluxo não só economiza o tempo da Roziane, como também cria um "dashboard" passivo de prioridades que ajuda na gestão de projetos.

Para implementar essa solução de forma robusta e escalável, a arquitetura mais eficiente envolve o uso da **Microsoft Graph API**, que centraliza o acesso tanto ao Outlook quanto ao Teams.

### Arquitetura Sugerida

1. **Backend (FastAPI):** Como você já domina Python e FastAPI, utilize-o para criar os endpoints que serão acionados por um *cron job* (ou GitHub Actions/vps) diariamente.
2. **Integração Microsoft Graph:** * Você precisará registrar um aplicativo no Azure AD para obter as credenciais de OAuth2.
    - A API permite buscar mensagens do dia anterior usando filtros `$filter` (ex: `receivedDateTime ge 2026-04-23T00:00:00Z`).
3. **Processamento com LLM:**
    - **Local (Ollama):** Se houver preocupação com privacidade de dados sensíveis da fundação, rodar um modelo como `Llama 3` ou `Mistral` via Ollama localmente é o ideal.
    - **Prompt Engineering:** O prompt deve solicitar explicitamente o resumo e a classificação. Ex: *"Analise o seguinte e-mail e retorne: 1. Resumo (2 frases), 2. Prioridade (Urgente/Média/Baixa) e 3. Ação necessária"*.
4. **Notificação no Teams:**
    - O jeito mais simples é via **Incoming Webhooks** no canal de projetos da Roziane. Você envia um JSON formatado como um "Adaptive Card", que permite uma visualização limpa com botões e cores para as prioridades.

### Copilot vs. Solução Customizada

O Microsoft 365 Copilot já possui o recurso "Summarize this email", mas ele é reativo (ela precisa clicar). Sua solução é **proativa**: ela acorda com um relatório consolidado no Teams sem precisar abrir o Outlook. Isso resolve o problema da sobrecarga cognitiva de olhar uma caixa de entrada lotada.

Abaixo, projetei um visualizador do fluxo de dados para ajudar a estruturar a lógica do seu script de integração.

```json?chameleon
{"component":"LlmGeneratedComponent","props":{"height":"700px","prompt":"Crie um fluxograma interativo usando D3.js que represente a arquitetura de automação de e-mails para o Teams. O fluxo deve ter 4 nós principais: 1. 'Microsoft Graph (Outlook)' -> 2. 'Processador FastAPI' -> 3. 'LLM (Ollama/API)' -> 4. 'Microsoft Teams (Notificação)'. \n\nDetalhes funcionais:\n- O usuário deve poder clicar em cada nó para ver os detalhes técnicos de implementação (ex: 'Filtros OData', 'OAuth2', 'Prompt de Categorização', 'Adaptive Cards').\n- Use animações para mostrar o fluxo de dados saindo do Outlook, sendo processado pela IA e chegando no Teams.\n- O layout deve ser vertical e responsivo.\n- Linguagem: Português.\n- Estilize de forma profissional e técnica, sem cores específicas, focando na clareza da lógica de engenharia de software.","id":"im_aa8d8cdcb551f429"}}
```

### Próximos Passos Técnicos

- **Categorização:** No seu script, você pode definir que e-mails de remetentes específicos (diretoria, parceiros estratégicos) recebam automaticamente um peso maior na prioridade.
- **Segurança:** Como você trabalha na fundação, lembre-se de armazenar as chaves do Azure no Supabase Vault ou em variáveis de ambiente (`.env`), nunca diretamente no código.
- **MVP:** Comece enviando um resumo simples em texto puro para o Teams e depois evolua para os *Adaptive Cards* formatados.

Essa solução tem grandes chances de ser um dos seus principais entregues como aspirante a engenheiro de IA na FAI. Se precisar de ajuda com a estrutura do JSON do Adaptive Card ou com a query do Graph API, é só dizer.

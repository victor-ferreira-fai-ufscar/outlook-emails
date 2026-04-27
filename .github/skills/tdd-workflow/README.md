# Skill de Desenvolvimento Orientado a Testes (TDD)

Esta skill guia você através de Test-Driven Development (TDD) para implementar novas features no projeto de integração FastAPI com Outlook.

## 📁 Arquivos da Skill

| Arquivo                        | Propósito                                                  | Tempo  |
| ------------------------------ | ---------------------------------------------------------- | ------ |
| **SKILL.md**                   | Guia completo com ciclo Vermelho-Verde-Refactor e exemplos | 15 min |
| **CHEATSHEET.md**              | Referência rápida para consulta diária                     | 3 min  |
| **EXAMPLE.md**                 | Exemplo real passo-a-passo                                 | 10 min |
| **INTEGRATION.md**             | Guia específico do projeto                                 | 10 min |
| **test-template.py**           | Template para testes - copie e adapte                      | -      |
| **implementation-template.py** | Template para implementação                                | -      |

## 🚀 Escolha seu caminho

### ⚡ Rápido (5 min)
1. Leia [CHEATSHEET.md](./CHEATSHEET.md)
2. Use `/tdd-workflow [descrição]` no Copilot Chat
3. Siga a orientação através de Vermelho → Verde → Refactor

### 📖 Aprender (15 min)
1. Leia [CHEATSHEET.md](./CHEATSHEET.md)
2. Siga [EXAMPLE.md](./EXAMPLE.md)
3. Use templates para suas features

### 🎓 Profundo (30 min)
1. Leia [README.md](./README.md) (este arquivo)
2. Leia [SKILL.md](./SKILL.md)
3. Siga [EXAMPLE.md](./EXAMPLE.md)
4. Consulte [INTEGRATION.md](./INTEGRATION.md)

## 🚀 Início Rápido

### 1. Use o comando `/tdd-workflow` no Copilot Chat

Digite `/tdd-workflow` seguido da descrição da feature:

```
/tdd-workflow Quero adicionar um endpoint /data/export que exporta dados do usuário
```

O Copilot vai guiar você através do workflow TDD.

### 2. Siga o Ciclo Vermelho-Verde-Refactor

**🔴 VERMELHO:** Escreva testes que falham
```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_sua_feature.py
# Edite com seus testes
pytest tests/test_sua_feature.py -v  # Deve falhar
```

**🟢 VERDE:** Implemente código mínimo
```bash
cp .github/skills/tdd-workflow/implementation-template.py app/routes/sua_feature.py
# Edite para passar nos testes
pytest tests/test_sua_feature.py -v  # Deve passar
```

**🔵 REFACTOR:** Limpe o código
```bash
# Melhore qualidade, nomes, comentários
pytest tests/ -v  # Certifique-se que nada quebrou
```

## 📊 Conceitos Principais

### O Ciclo TDD

1. **🔴 VERMELHO**: Escreva testes que descrevem a feature → testes falham
2. **🟢 VERDE**: Escreva código mínimo para passar nos testes → testes passam
3. **🔵 REFACTOR**: Melhore qualidade do código → testes ainda passam

Isso garante:
- ✅ Código funciona como esperado (testes provam)
- ✅ Sem over-engineering (escreva só o necessário)
- ✅ Código limpo e mantível (refatorado ao final)

### Estrutura de Teste

```python
# Setup (arranjo)
session_id = authenticated_session

# Action (ação)
response = client.get(
    "/endpoint",
    cookies={"local_session_id": session_id}
)

# Verification (verificação)
assert response.status_code == 200
```

## 📋 Comandos Úteis

```bash
# Suite completa
pytest tests/ -v

# Com cobertura
pytest --cov=app tests/

# Arquivo específico
pytest tests/test_sua_feature.py -v

# Parar no primeiro erro
pytest -x
```

## 🎓 Recursos

- [Exemplo Real Completo](./EXAMPLE.md)
- [Guia do Projeto](./INTEGRATION.md)
- [Referência Rápida](./CHEATSHEET.md)
- [Guia Completo](./SKILL.md)

## ✨ Próximos Passos

Escolha uma:
1. **Teste agora**: `/tdd-workflow [sua feature]` no Copilot Chat
2. **Aprenda vendo**: Siga [EXAMPLE.md](./EXAMPLE.md)
3. **Copie template**: `cp test-template.py tests/test_xxx.py`
4. **Leia guia**: Abra [SKILL.md](./SKILL.md)

---

**Lembre-se: Testes escritos primeiro = código que você pode confiar! 🚀**

## Key Concepts

### The Red-Green-Refactor Cycle

1. **RED** 🔴: Write tests that describe the feature → tests fail
2. **GREEN** 🟢: Write minimal code to make tests pass → tests pass
3. **REFACTOR** 🔵: Improve code quality → tests still pass

This ensures:
- Code works as expected (tests prove it)
- No over-engineering (write only what's needed)
- Clean, maintainable code (refactored at the end)

### Test Structure

```python
# Setup (arrange)
session_id = authenticated_session

# Action (act)
response = client.get(
    "/endpoint",
    cookies={"local_session_id": session_id}
)

# Verification (assert)
assert response.status_code == 200
```

### Common Test Patterns

**Testing authenticated endpoints:**
```python
def test_endpoint_requires_auth(self):
    response = client.get("/protected-endpoint")
    assert response.status_code == 401

def test_endpoint_with_valid_auth(self, authenticated_session):
    response = client.get(
        "/protected-endpoint",
        cookies={"local_session_id": authenticated_session}
    )
    assert response.status_code == 200
```

**Mocking external APIs:**
```python
from unittest.mock import patch, MagicMock

def test_with_mocked_api(self):
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"data": "value"}
    
    with patch("app.utils.requests.get", return_value=mock_resp):
        response = client.get("/endpoint")
    
    assert response.status_code == 200
```

## Best Practices

✅ **DO:**
- Write descriptive test names: `test_endpoint_returns_401_without_auth`
- Test one behavior per test function
- Mock external dependencies (APIs, databases)
- Test both happy path and error cases
- Run full test suite before committing

❌ **DON'T:**
- Skip tests for "quick fixes"
- Mix multiple assertions without clear intent
- Test implementation details instead of behavior
- Write tests after code (that's not TDD!)
- Commit code without running `pytest`

## Project-Specific Notes

### For this FastAPI Outlook Integration:

**Authentication:**
- Use `authenticated_session` fixture for tests requiring auth
- Use `mock.patch("app.utils.requests.get")` for Graph API calls
- Session files are auto-created in `sessions/` directory

**File Structure:**
```
app/
├── main.py          # Router registration only
├── config.py        # Configuration constants
├── utils.py         # Helper functions, Graph API calls
└── routes/
    ├── auth.py      # OAuth endpoints
    ├── profile.py   # Profile endpoints
    └── messages.py  # Email endpoints
```

**Running Tests:**
```bash
# Full test suite
pytest tests/ -v

# Specific test file
pytest tests/test_auth.py -v

# Stop at first failure
pytest -x

# With coverage report
pytest --cov=app tests/
```

## Resources

- [Complete TDD Guide](./SKILL.md)
- [Test Template](./test-template.py)
- [Implementation Template](./implementation-template.py)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)

## Need Help?

1. Read the complete [SKILL.md](./SKILL.md) guide
2. Copy and adapt the templates
3. Follow the Red-Green-Refactor cycle
4. Run tests frequently
5. Ask Copilot: `/tdd-workflow [your feature description]`

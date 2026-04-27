# TDD Workflow - Quick Reference Card

## 🚀 Quick Start (2 minutos)

```bash
# 1. Copie os templates
cp .github/skills/tdd-workflow/test-template.py tests/test_seu_recurso.py
cp .github/skills/tdd-workflow/implementation-template.py app/routes/seu_recurso.py

# 2. RED - Escrever testes (devem falhar)
pytest tests/test_seu_recurso.py -v
# Esperado: ❌ FAILED

# 3. GREEN - Implementar (testes devem passar)
pytest tests/test_seu_recurso.py -v
# Esperado: ✅ PASSED

# 4. REFACTOR - Melhorar (testes continuam passando)
pytest tests/ -v
# Esperado: ✅ PASSED
```

---

## 📋 Checklist TDD

### RED Phase 🔴
- [ ] Criar arquivo de teste: `tests/test_[feature].py`
- [ ] Escrever testes que descrevem o comportamento
- [ ] Incluir: happy path + error cases
- [ ] Mockar dependências externas
- [ ] Rodar: `pytest tests/test_[feature].py -v`
- [ ] Confirmar: ❌ TODOS FALHAM
- [ ] Commit: `git commit -m "test: add tests for [feature]"`

### GREEN Phase 🟢
- [ ] Criar arquivo de implementação: `app/routes/[feature].py`
- [ ] Implementar código MÍNIMO para passar testes
- [ ] Registrar router em `app/main.py`
- [ ] Rodar: `pytest tests/test_[feature].py -v`
- [ ] Confirmar: ✅ TODOS PASSAM
- [ ] Commit: `git commit -m "feat: implement [feature]"`

### REFACTOR Phase 🔵
- [ ] Melhorar nomes de variáveis/funções
- [ ] Adicionar/melhorar docstrings
- [ ] Extrair lógica para helper functions
- [ ] Rodar: `pytest tests/ -v`
- [ ] Confirmar: ✅ TODOS AINDA PASSAM
- [ ] Commit: `git commit -m "refactor: improve [feature] code quality"`

---

## 🧪 Padrões Comuns de Teste

### Teste: Autenticação Requerida
```python
def test_endpoint_requires_auth():
    response = client.get("/protected")
    assert response.status_code == 401
```

### Teste: Com Sessão Válida
```python
def test_endpoint_with_auth(authenticated_session):
    response = client.get(
        "/protected",
        cookies={"local_session_id": authenticated_session}
    )
    assert response.status_code == 200
```

### Teste: Mockando API Externa
```python
from unittest.mock import patch, MagicMock

def test_with_mocked_api(authenticated_session):
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"data": "value"}
    
    with patch("app.utils.requests.get", return_value=mock_resp):
        response = client.get("/endpoint", cookies=...)
    
    assert response.status_code == 200
```

### Teste: Erro da API Externa
```python
def test_api_error(authenticated_session):
    with patch("app.utils.requests.get") as mock:
        mock.return_value.status_code = 403
        mock.return_value.text = "Forbidden"
        
        response = client.get("/endpoint", cookies=...)
    
    assert response.status_code == 403
```

---

## 🛠️ Padrões Comuns de Implementação

### Endpoint Simples
```python
from fastapi import APIRouter, Request
from app.utils import get_local_access_token

router = APIRouter(prefix="/data", tags=["data"])

@router.get("/endpoint")
def get_data(request: Request):
    """Descrição clara do que o endpoint faz."""
    access_token = get_local_access_token(request)  # 401 se não auth
    return {"data": "value"}
```

### Endpoint com Helper
```python
@router.get("/endpoint")
def get_data(request: Request):
    access_token = get_local_access_token(request)
    raw_data = _fetch_external_data(access_token)
    return _process_data(raw_data)

def _fetch_external_data(access_token: str):
    # Lógica de fetch
    pass

def _process_data(data):
    # Lógica de processamento
    pass
```

### Registrar Router em main.py
```python
from app.routes import seu_arquivo

app.include_router(seu_arquivo.router)
```

---

## 📊 Comandos Pytest Úteis

```bash
# Rodar tudo
pytest

# Rodar com verbose
pytest -v

# Rodar arquivo específico
pytest tests/test_seu_arquivo.py

# Parar no primeiro erro
pytest -x

# Rodar apenas um teste
pytest tests/test_seu_arquivo.py::test_funcao_especifica

# Com coverage
pytest --cov=app tests/

# Modo watch (rerun on changes)
pytest-watch tests/

# Com output mais bonito
pytest -v --tb=short
```

---

## ✅ Boas Práticas

### Nomes de Testes
```python
# ❌ Ruim
def test_func():
    ...

# ✅ Bom
def test_endpoint_returns_401_without_auth():
    ...
```

### Uma Coisa por Teste
```python
# ❌ Ruim - testa 3 coisas
def test_endpoint():
    assert status == 200
    assert "field" in data
    assert data["field"] == "value"

# ✅ Bom - testa 1 coisa
def test_endpoint_returns_200():
    assert response.status_code == 200

def test_endpoint_response_has_field():
    assert "field" in response.json()

def test_endpoint_field_has_correct_value():
    assert response.json()["field"] == "value"
```

### Mockar Tudo que é Externo
```python
# ❌ Ruim - chamada real
with patch("app.utils.requests.get") as mock:
    mock.return_value = actual_api_call()  # ❌ Chamada real!

# ✅ Bom - mock completo
mock_resp = MagicMock()
mock_resp.status_code = 200
mock_resp.json.return_value = {"data": "mocked"}
with patch("app.utils.requests.get", return_value=mock_resp):
    ...
```

---

## 🔗 Links Úteis

- [Skill Completa](./SKILL.md)
- [Exemplo Passo-a-Passo](./EXAMPLE.md)
- [README com Detalhes](./README.md)
- [Test Template](./test-template.py)
- [Implementation Template](./implementation-template.py)

---

## 💡 Exemplo Rápido (5 minutos)

**Sua feature:** Adicionar POST `/save` que salva dados

### RED - Escrever teste
```python
def test_save_endpoint_saves_data(authenticated_session):
    response = client.post(
        "/save",
        json={"name": "test"},
        cookies={"local_session_id": authenticated_session}
    )
    assert response.status_code == 200
    assert response.json()["saved"] == True
```

### GREEN - Implementar
```python
@router.post("/save")
def save_data(request: Request, data: dict):
    access_token = get_local_access_token(request)
    return {"saved": True}
```

### REFACTOR - Melhorar
```python
@router.post("/save")
def save_user_data(request: Request, data: dict):
    """Salva dados do usuário autenticado."""
    access_token = get_local_access_token(request)
    _persist_data(data)
    return {"saved": True}

def _persist_data(data: dict):
    """Helper para persistência."""
    # implementação
    pass
```

---

## ❓ FAQ Rápido

**P: E se meu teste passar sem implementar nada?**
R: Significa seu teste é fraco. Melhore o teste para forçar implementação.

**P: Quantos testes preciso escrever?**
R: Pelo menos: 1 happy path + 1 error case. Mais é melhor!

**P: Posso fazer commit apenas na phase RED?**
R: Sim! Commits pequenos são bons. Ideal: test → impl → refactor.

**P: E se refactoring quebrar um teste?**
R: Volte ao código anterior. Você quebrou algo. Refactor mais cuidadosamente.

**P: Como mockar coisas complexas?**
R: Use `MagicMock()` para simular comportamentos. Leia unittest.mock docs.

---

## 🎯 Resumo: Red-Green-Refactor

| Fase     | Cor | O que fazer        | Testes |
| -------- | --- | ------------------ | ------ |
| RED      | 🔴   | Escrever testes    | ❌ FAIL |
| GREEN    | 🟢   | Implementar código | ✅ PASS |
| REFACTOR | 🔵   | Melhorar qualidade | ✅ PASS |

**Lembre-se:** Testes dirigi o design!

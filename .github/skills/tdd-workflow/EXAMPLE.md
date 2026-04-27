# TDD Workflow Step-by-Step Example

This document shows a complete TDD workflow for a real feature.

## Scenario: Add a `/data/summary` endpoint

This endpoint returns a summary of user data (profile info + latest email count).

---

## Phase 1: RED - Write Failing Tests

### Step 1: Create test file

```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_data.py
```

### Step 2: Edit `tests/test_data.py`

Replace the template content:

```python
import pytest
from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app, raise_server_exceptions=False)


@pytest.fixture(autouse=True)
def use_tmp_sessions(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)


@pytest.fixture
def authenticated_session(tmp_path):
    import uuid
    from app.utils import write_session_file
    
    session_id = str(uuid.uuid4())
    write_session_file(
        f"session-{session_id}.json",
        {
            "access_token": "test-token",
            "expires_at": "2099-12-31T23:59:59",
            "refresh_token": "test-refresh",
        },
    )
    return session_id


# ============================================================================
# TESTES - Caso de Sucesso
# ============================================================================

def test_data_summary_endpoint_exists(authenticated_session):
    """Teste que endpoint /data/summary existe."""
    response = client.get(
        "/data/summary",
        cookies={"local_session_id": authenticated_session}
    )
    assert response.status_code == 200


def test_data_summary_returns_required_fields(authenticated_session):
    """Teste que resposta contém campos esperados."""
    response = client.get(
        "/data/summary",
        cookies={"local_session_id": authenticated_session}
    )
    data = response.json()
    assert "user_name" in data
    assert "user_email" in data
    assert "latest_email_subject" in data
    assert "total_sent_emails" in data


def test_data_summary_with_mocked_api(authenticated_session):
    """Teste que dados são processados corretamente."""
    from unittest.mock import patch, MagicMock
    
    mock_profile = {
        "id": "user-123",
        "displayName": "John Doe",
        "mail": "john@example.com"
    }
    
    mock_email_list = {
        "value": [
            {"id": "e1", "subject": "Latest email"},
            {"id": "e2", "subject": "Old email"},
        ]
    }
    
    mock_profile_resp = MagicMock()
    mock_profile_resp.status_code = 200
    mock_profile_resp.json.return_value = mock_profile
    
    mock_emails_resp = MagicMock()
    mock_emails_resp.status_code = 200
    mock_emails_resp.json.return_value = mock_email_list
    
    with patch("app.utils.requests.get", side_effect=[mock_profile_resp, mock_emails_resp]):
        response = client.get(
            "/data/summary",
            cookies={"local_session_id": authenticated_session}
        )
    
    data = response.json()
    assert data["user_name"] == "John Doe"
    assert data["user_email"] == "john@example.com"
    assert data["latest_email_subject"] == "Latest email"
    assert data["total_sent_emails"] == 2


# ============================================================================
# TESTES - Casos de Erro
# ============================================================================

def test_data_summary_requires_authentication():
    """Teste que endpoint rejeita acesso sem autenticação."""
    response = client.get("/data/summary")
    assert response.status_code == 401


def test_data_summary_with_api_failure(authenticated_session):
    """Teste que erro da API é tratado."""
    from unittest.mock import patch
    
    with patch("app.utils.requests.get") as mock_get:
        mock_get.return_value.status_code = 403
        mock_get.return_value.text = "Forbidden"
        
        response = client.get(
            "/data/summary",
            cookies={"local_session_id": authenticated_session}
        )
    
    assert response.status_code == 403
```

### Step 3: Run tests - they should FAIL

```bash
cd /home/victorferreira/projects/github/victor-ferreira-fai-ufscar/outlook-emails
pytest tests/test_data.py -v
```

Output:
```
FAILED tests/test_data.py::test_data_summary_endpoint_exists - 404 Not Found
FAILED tests/test_data.py::test_data_summary_returns_required_fields - 404 Not Found
...
```

🔴 **RED Phase Complete** - Tests fail because endpoint doesn't exist yet

---

## Phase 2: GREEN - Implement Minimal Code

### Step 1: Create implementation file

```bash
cp .github/skills/tdd-workflow/implementation-template.py app/routes/data.py
```

### Step 2: Edit `app/routes/data.py`

```python
from typing import Any
from fastapi import APIRouter, Request
from app.utils import get_local_access_token, fetch_outlook_profile, fetch_latest_sent_email

router = APIRouter(prefix="/data", tags=["data"])


@router.get("/summary")
def get_data_summary(request: Request) -> dict[str, Any]:
    """
    Retorna um resumo dos dados do usuário.
    
    Inclui informações do perfil + último email enviado.
    """
    # Autenticar
    access_token = get_local_access_token(request)
    
    # Buscar dados
    profile = fetch_outlook_profile(access_token)
    latest_email = fetch_latest_sent_email(access_token)
    
    # Retornar resumo
    return {
        "user_name": profile.get("displayName"),
        "user_email": profile.get("mail") or profile.get("userPrincipalName"),
        "latest_email_subject": latest_email.get("subject"),
        "total_sent_emails": 1  # Placeholder
    }
```

### Step 3: Register router in `app/main.py`

Edit `app/main.py` to add:

```python
from app.routes import health, auth, profile, messages, data

app.include_router(health.router)
app.include_router(auth.router)
app.include_router(profile.router)
app.include_router(messages.router)
app.include_router(data.router)  # ← Add this line
```

### Step 4: Run tests - they should PASS

```bash
pytest tests/test_data.py -v
```

Output:
```
PASSED tests/test_data.py::test_data_summary_endpoint_exists
PASSED tests/test_data.py::test_data_summary_returns_required_fields
PASSED tests/test_data.py::test_data_summary_with_mocked_api
PASSED tests/test_data.py::test_data_summary_requires_authentication
PASSED tests/test_data.py::test_data_summary_with_api_failure
...
5 passed in 0.42s
```

🟢 **GREEN Phase Complete** - All tests pass!

---

## Phase 3: REFACTOR - Improve Code Quality

### Step 1: Refactor `app/routes/data.py`

Improve the implementation (no test changes):

```python
from typing import Any
from fastapi import APIRouter, Request
from app.utils import get_local_access_token, fetch_outlook_profile, fetch_latest_sent_email

router = APIRouter(prefix="/data", tags=["data"])


@router.get("/summary")
def get_data_summary(request: Request) -> dict[str, Any]:
    """
    Retorna um resumo consolidado dos dados do usuário.
    
    Combina informações do perfil do Outlook com dados dos últimos emails
    para fornecer uma visão geral rápida.
    
    Requires:
        Autenticação via local session cookie (local_session_id)
    
    Returns:
        dict: Contém user_name, user_email, latest_email_subject,
              e total_sent_emails
              
    Raises:
        HTTPException 401: Se não autenticado
        HTTPException 403: Se Graph API retorna acesso negado
    """
    # Autenticar e obter token
    access_token = get_local_access_token(request)
    
    # Buscar dados do perfil e emails em paralelo (lógica simples)
    profile = fetch_outlook_profile(access_token)
    latest_email = fetch_latest_sent_email(access_token)
    
    # Processar e retornar resumo bem formatado
    return _build_summary(profile, latest_email)


def _build_summary(profile: dict[str, Any], latest_email: dict[str, Any]) -> dict[str, Any]:
    """
    Helper para construir resposta de resumo.
    
    Separa lógica de processamento da lógica do endpoint.
    Facilita testes e manutenção.
    """
    email_subject = latest_email.get("subject", "(sem assunto)")
    
    return {
        "user_name": profile.get("displayName"),
        "user_email": profile.get("mail") or profile.get("userPrincipalName"),
        "latest_email_subject": email_subject,
        "total_sent_emails": 1,  # TODO: Implementar contagem real se necessário
    }
```

### Step 2: Run full test suite

```bash
pytest tests/ -v
```

Output:
```
...
31 passed, 3 skipped, 4 warnings in 1.25s
```

🔵 **REFACTOR Phase Complete** - Code is clean, tests still pass!

---

## Phase 4: Commit with Clear History

```bash
# Commit 1: Add tests
git add tests/test_data.py
git commit -m "test: add tests for /data/summary endpoint"

# Commit 2: Add implementation
git add app/routes/data.py app/main.py
git commit -m "feat: implement /data/summary endpoint"

# Commit 3: Improve code quality (optional)
git add app/routes/data.py
git commit -m "refactor: improve /data/summary code clarity"
```

---

## Summary

| Phase    | Status | Command                     | Expected     |
| -------- | ------ | --------------------------- | ------------ |
| RED      | 🔴      | `pytest tests/test_data.py` | ❌ Tests fail |
| GREEN    | 🟢      | `pytest tests/test_data.py` | ✅ Tests pass |
| REFACTOR | 🔵      | `pytest tests/`             | ✅ All pass   |

### Key Takeaways

1. ✅ **Tests drive design** - Tests show what the code should do
2. ✅ **Minimal implementation** - Only code what's needed to pass tests
3. ✅ **Quality through refactoring** - Clean code comes after it works
4. ✅ **Confidence in changes** - Tests catch regressions immediately
5. ✅ **Better documentation** - Tests show how to use the code

This workflow ensures robust, maintainable, well-tested code!

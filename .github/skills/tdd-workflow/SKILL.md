---
name: tdd-workflow
description: "Ciclo Test-Driven Development (TDD) para implementar novas features. Use quando começar qualquer feature, endpoint ou função. Siga o ciclo Vermelho-Verde-Refactor: escreva testes que falham primeiro, implemente código mínimo para passar nos testes, depois refatore para qualidade."
argument-hint: "[descrição da feature]"
user-invocable: true
---

# Desenvolvimento Orientado a Testes (TDD)

Este skill guia você através da implementação de novas features usando Test-Driven Development. Em vez de escrever código primeiro e testes depois, TDD escreve os testes primeiro para definir o comportamento esperado, depois implementa o código que satisfaz esses testes.

## Por Que TDD?

- ✅ **Requisitos Claros**: Testes documentam o que o código deve fazer
- ✅ **Maior Confiança**: Código testado desde o primeiro dia
- ✅ **Menos Bugs**: Defeitos encontrados cedo, não em produção
- ✅ **Melhor Design**: Escrever testes primeiro leva a melhor arquitetura
- ✅ **Refatoração Segura**: Testes detectam regressões imediatamente

## O Ciclo TDD: Vermelho → Verde → Refactor

### 1️⃣ RED: Write Failing Tests

Write test cases that describe the desired behavior. These tests **should fail** because the feature doesn't exist yet.

**Checklist:**
- [ ] Create test file (e.g., `tests/test_new_feature.py`)
- [ ] Write descriptive test function names (e.g., `test_endpoint_returns_200_on_success`)
- [ ] Mock external dependencies (APIs, databases)
- [ ] Include both happy path and error cases
- [ ] Run tests to confirm they fail: `pytest tests/test_new_feature.py`

**Example test structure for FastAPI endpoint:**
```python
import pytest
from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)

def test_new_endpoint_returns_success():
    """Test that /new-endpoint returns expected data."""
    response = client.get("/new-endpoint")
    assert response.status_code == 200
    assert "expected_field" in response.json()

def test_new_endpoint_requires_auth():
    """Test that /new-endpoint rejects unauthenticated requests."""
    response = client.get("/new-endpoint")
    assert response.status_code == 401
```

### 2️⃣ GREEN: Implement Minimal Code

Write the **simplest code possible** to make the tests pass. Don't over-engineer or add extra features.

**Checklist:**
- [ ] Create the new function/endpoint/module
- [ ] Write only code needed to pass tests
- [ ] Keep implementation focused and minimal
- [ ] Run tests again: `pytest tests/test_new_feature.py -v`
- [ ] Verify all tests pass (green ✅)

**Example implementation:**
```python
from fastapi import APIRouter, HTTPException
from app.utils import get_local_access_token

router = APIRouter(prefix="/new", tags=["features"])

@router.get("/endpoint")
def new_endpoint(request):
    """Simple endpoint implementation."""
    access_token = get_local_access_token(request)
    return {
        "expected_field": "value",
        "status": "success"
    }
```

### 3️⃣ REFACTOR: Clean Up Code

Now that tests pass, improve code quality without changing functionality.

**Checklist:**
- [ ] Remove duplicate code
- [ ] Add helpful comments/docstrings
- [ ] Improve variable/function names
- [ ] Extract complex logic to helper functions
- [ ] Run tests again to ensure nothing broke: `pytest tests/ -v`

**Refactored example:**
```python
@router.get("/endpoint")
def get_new_feature_data(request):
    """
    Retrieve processed data for new feature.
    
    Requires authentication via local session cookie.
    Returns feature data with status indicator.
    """
    access_token = get_local_access_token(request)
    feature_data = _fetch_and_process_data(access_token)
    return feature_data

def _fetch_and_process_data(access_token: str) -> dict:
    """Helper function for data processing."""
    # Implementation details here
    pass
```

## Running Tests

### Run all tests:
```bash
pytest
```

### Run specific test file:
```bash
pytest tests/test_new_feature.py
```

### Run with verbose output:
```bash
pytest -v
```

### Run and stop at first failure:
```bash
pytest -x
```

### Run with coverage report:
```bash
pytest --cov=app tests/
```

## Best Practices

### 1. Write Clear Test Names
Bad: `test_func()`, `test_data()`
Good: `test_endpoint_returns_401_without_auth()`, `test_token_refresh_updates_session_file()`

### 2. Test One Thing Per Test
Each test should verify a single behavior. Use multiple test functions for multiple scenarios.

```python
# ❌ Bad: tests multiple behaviors in one function
def test_endpoint():
    response = client.get("/endpoint", cookies={"session_id": "123"})
    assert response.status_code == 200
    assert "data" in response.json()
    assert response.json()["data"]["field"] == "value"

# ✅ Good: separate tests for separate concerns
def test_endpoint_requires_auth():
    response = client.get("/endpoint")
    assert response.status_code == 401

def test_endpoint_returns_200_with_valid_session():
    response = client.get("/endpoint", cookies={"session_id": "123"})
    assert response.status_code == 200

def test_endpoint_response_contains_expected_fields():
    response = client.get("/endpoint", cookies={"session_id": "123"})
    assert "data" in response.json()
```

### 3. Mock External Dependencies
Use `unittest.mock` to avoid actual API/database calls in tests:

```python
from unittest.mock import patch, MagicMock

def test_with_mocked_api():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"result": "mocked"}
    
    with patch("app.utils.requests.get", return_value=mock_response):
        response = client.get("/endpoint", cookies={"session_id": "123"})
    
    assert response.status_code == 200
```

### 4. Use Fixtures for Setup/Teardown
Keep test code DRY with reusable fixtures:

```python
import pytest

@pytest.fixture
def authenticated_client(tmp_path, monkeypatch):
    """Fixture that provides a test client with valid session."""
    monkeypatch.chdir(tmp_path)
    session_id = "test-session-123"
    # Create session file
    from app.utils import write_session_file
    write_session_file(f"session-{session_id}.json", {
        "access_token": "test-token",
        "expires_at": "2099-12-31T23:59:59",
        "refresh_token": "refresh"
    })
    return session_id

def test_with_fixture(authenticated_client):
    response = client.get(
        "/endpoint", 
        cookies={"local_session_id": authenticated_client}
    )
    assert response.status_code == 200
```

### 5. Test Error Cases
Don't only test happy path - test errors:

```python
def test_endpoint_with_expired_token():
    """Test handling of expired authentication tokens."""
    response = client.get("/endpoint", cookies={"local_session_id": "expired"})
    assert response.status_code == 401

def test_endpoint_with_missing_dependency():
    """Test graceful failure when external API is unavailable."""
    with patch("app.utils.requests.get", side_effect=TimeoutError):
        response = client.get("/endpoint", cookies={"session_id": "123"})
    assert response.status_code >= 500
```

## TDD Workflow Checklist

When starting a new feature:

- [ ] **RED Phase**
  - [ ] Create test file in `tests/`
  - [ ] Write failing tests (happy path + error cases)
  - [ ] Run tests, confirm they fail
  - [ ] Commit: "test: add tests for [feature]"

- [ ] **GREEN Phase**
  - [ ] Implement minimal code to pass tests
  - [ ] Run tests, confirm they all pass
  - [ ] Commit: "feat: implement [feature]"

- [ ] **REFACTOR Phase**
  - [ ] Improve code quality (no test changes)
  - [ ] Run tests, confirm they still pass
  - [ ] Commit: "refactor: improve [feature] code quality"

- [ ] **Integration**
  - [ ] Run full test suite: `pytest tests/ -v`
  - [ ] Check coverage if desired
  - [ ] Create PR with all commits

## Example: Complete TDD Session

### Scenario: Add a new `/health/detailed` endpoint

**RED - Write Tests:**
```python
def test_detailed_health_endpoint_exists():
    response = client.get("/health/detailed")
    assert response.status_code == 200

def test_detailed_health_returns_required_fields():
    response = client.get("/health/detailed")
    data = response.json()
    assert "status" in data
    assert "version" in data
    assert "timestamp" in data
```

Run: `pytest tests/test_health.py` → ❌ FAILS (endpoint doesn't exist)

**GREEN - Implement:**
```python
from datetime import datetime
import app

@app.get("/health/detailed")
def detailed_health():
    return {
        "status": "ok",
        "version": "0.1.0",
        "timestamp": datetime.now().isoformat()
    }
```

Run: `pytest tests/test_health.py` → ✅ PASSES (all tests pass)

**REFACTOR - Improve:**
```python
from datetime import datetime, timezone

@app.get("/health/detailed")
def get_detailed_health_status():
    """Return detailed application health status."""
    return {
        "status": "operational",
        "version": "0.1.0",
        "timestamp": datetime.now(timezone.utc).isoformat()
    }
```

Run: `pytest tests/ -v` → ✅ ALL PASS (no regressions)

## Resources

- [GitHub Copilot TDD Guide](https://github.blog/ai-and-ml/github-copilot/github-for-beginners-test-driven-development-tdd-with-github-copilot/)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing Guide](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock Documentation](https://docs.python.org/3/library/unittest.mock.html)

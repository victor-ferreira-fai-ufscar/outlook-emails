"""
Template para testes TDD.
Copie este arquivo, adapte os nomes e escreva seus testes.
"""

import pytest
from fastapi.testclient import TestClient
from datetime import datetime, timedelta, timezone
from unittest.mock import patch, MagicMock

from app.main import app

client = TestClient(app, raise_server_exceptions=False)


@pytest.fixture(autouse=True)
def use_tmp_sessions(tmp_path, monkeypatch):
    """Isola sessões em diretório temporário."""
    monkeypatch.chdir(tmp_path)


@pytest.fixture
def authenticated_session(tmp_path):
    """Cria uma sessão autenticada para testes."""
    import uuid
    from app.utils import write_session_file

    session_id = str(uuid.uuid4())
    expires_at = (datetime.now(timezone.utc) + timedelta(hours=1)).isoformat()
    write_session_file(f"session-{session_id}.json", {
        "access_token": "test-token",
        "expires_at": expires_at,
        "refresh_token": "test-refresh",
    })
    return session_id


# TESTES - CASO DE SUCESSO
def test_new_endpoint_returns_200(authenticated_session):
    """Endpoint retorna 200 com sessão válida."""
    response = client.get("/new-endpoint", cookies={"local_session_id": authenticated_session})
    assert response.status_code == 200


def test_new_endpoint_returns_expected_fields(authenticated_session):
    """Resposta contém campos esperados."""
    response = client.get("/new-endpoint", cookies={"local_session_id": authenticated_session})
    assert "field" in response.json()


# TESTES - ERRO
def test_new_endpoint_requires_auth():
    """Endpoint rejeita requisições sem autenticação."""
    response = client.get("/new-endpoint")
    assert response.status_code == 401


# TESTES - EDGE CASES
def test_new_endpoint_with_empty_response():
    """Endpoint trata respostas vazias."""
    pass

        "/new-endpoint", cookies={"local_session_id": authenticated_session}
    )
    data = response.json()
    assert "field_one" in data
    assert "field_two" in data
    assert data["status"] == "success"


def test_new_endpoint_returns_correct_data(authenticated_session):
    """
    RED + GREEN: Teste que endpoint retorna dados corretos.

    Comportamento esperado: Dados retornados correspondem ao esperado.
    """
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"id": "123", "name": "Test"}

    with patch("app.utils.requests.get", return_value=mock_response):
        response = client.get(
            "/new-endpoint", cookies={"local_session_id": authenticated_session}
        )

    data = response.json()
    assert data["id"] == "123"
    assert data["name"] == "Test"


# ============================================================================
# TESTES - CASOS DE ERRO (Error Cases)
# ============================================================================


def test_new_endpoint_requires_authentication():
    """
    RED + GREEN: Teste que endpoint protegido rejeita acesso sem auth.

    Comportamento esperado: GET sem session_id retorna 401.
    """
    response = client.get("/new-endpoint")
    assert response.status_code == 401


def test_new_endpoint_rejects_invalid_session_id():
    """
    RED + GREEN: Teste que session_id inválida é rejeitada.

    Comportamento esperado: Session inexistente retorna 401.
    """
    response = client.get(
        "/new-endpoint", cookies={"local_session_id": "nonexistent-session"}
    )
    assert response.status_code == 401


def test_new_endpoint_handles_expired_token(tmp_path, monkeypatch):
    """
    RED + GREEN: Teste que token expirado tenta refresh automaticamente.

    Comportamento esperado: Token expirado sem refresh_token retorna 401.
    """
    monkeypatch.chdir(tmp_path)
    import uuid
    from app.utils import write_session_file

    session_id = str(uuid.uuid4())
    past = (datetime.now(timezone.utc) - timedelta(seconds=10)).isoformat()

    write_session_file(
        f"session-{session_id}.json",
        {
            "access_token": "old-token",
            "expires_at": past,
            "refresh_token": None,  # Sem refresh token
        },
    )

    response = client.get("/new-endpoint", cookies={"local_session_id": session_id})
    assert response.status_code == 401


def test_new_endpoint_with_api_error(authenticated_session):
    """
    RED + GREEN: Teste comportamento quando API externa falha.

    Comportamento esperado: Erro da API é propagado corretamente.
    """
    with patch("app.utils.requests.get") as mock_get:
        mock_get.return_value.status_code = 403
        mock_get.return_value.text = "Forbidden"

        response = client.get(
            "/new-endpoint", cookies={"local_session_id": authenticated_session}
        )

    assert response.status_code == 403


# ============================================================================
# TESTES - EDGE CASES
# ============================================================================


def test_new_endpoint_with_empty_response(authenticated_session):
    """
    RED + GREEN: Teste com resposta vazia da API.

    Comportamento esperado: Endpoint lidando graciosamente com response vazia.
    """
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {}

    with patch("app.utils.requests.get", return_value=mock_response):
        response = client.get(
            "/new-endpoint", cookies={"local_session_id": authenticated_session}
        )

    assert response.status_code == 200


def test_new_endpoint_with_special_characters(authenticated_session):
    """
    RED + GREEN: Teste com dados contendo caracteres especiais/unicode.

    Comportamento esperado: Endpoint processa unicode corretamente.
    """
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "name": "José Silva",
        "email": "josé@exemplo.com.br",
        "description": "✅ Teste com emoji",
    }

    with patch("app.utils.requests.get", return_value=mock_response):
        response = client.get(
            "/new-endpoint", cookies={"local_session_id": authenticated_session}
        )

    data = response.json()
    assert "José" in data["name"]
    assert "✅" in data["description"]


# ============================================================================
# PASSOS PARA USAR ESTE TEMPLATE
# ============================================================================

"""
1. Copie este arquivo: cp .github/skills/tdd-workflow/test-template.py tests/test_seu_recurso.py

2. Adapte os nomes:
   - test_new_endpoint → test_seu_endpoint
   - /new-endpoint → /seu-endpoint
   - authenticated_session → seu_fixture se precisar
   
3. Escreva testes (RED):
   - Execute: pytest tests/test_seu_recurso.py -v
   - Veja tudo falhar (vermelho)
   
4. Implemente a feature (GREEN):
   - Crie o novo endpoint em app/routes/seu_arquivo.py
   - Execute: pytest tests/test_seu_recurso.py -v
   - Veja tudo passar (verde)
   
5. Refatore se precisar (REFACTOR):
   - Melhore nomes, código, documentação
   - Execute: pytest tests/ -v
   - Garanta que nada quebrou
   
6. Commit com histórico claro:
   - "test: add tests for new endpoint"
   - "feat: implement new endpoint"
   - "refactor: improve new endpoint code quality"
"""

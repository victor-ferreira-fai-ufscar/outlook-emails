"""
Testes de integracao dos endpoints FastAPI usando TestClient.
Nenhuma chamada real a APIs externas ou ao Azure AD é feita - tudo mockado.
"""
import json
import uuid
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from app.main import app, _write_session_file

client = TestClient(app, raise_server_exceptions=False)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=True)
def use_tmp_sessions(tmp_path, monkeypatch):
    """Redireciona leitura/escrita de sessions e data para tmp_path."""
    monkeypatch.chdir(tmp_path)


def _create_session(
    tmp_path,
    access_token: str = "valid-token",
    expires_offset_seconds: int = 3600,
    refresh_token: str | None = "my-refresh",
) -> str:
    session_id = str(uuid.uuid4())
    expires_at = (
        datetime.now(timezone.utc) + timedelta(seconds=expires_offset_seconds)
    ).isoformat()
    _write_session_file(
        f"session-{session_id}.json",
        {
            "access_token": access_token,
            "expires_at": expires_at,
            "refresh_token": refresh_token,
        },
    )
    return session_id


# ---------------------------------------------------------------------------
# GET /
# ---------------------------------------------------------------------------


def test_root_returns_200():
    response = client.get("/")
    assert response.status_code == 200
    body = response.json()
    assert "message" in body
    assert "next_step" in body


# ---------------------------------------------------------------------------
# GET /auth/login
# ---------------------------------------------------------------------------


def test_auth_login_redirects_to_microsoft():
    # TestClient uses 'testclient' as host, which differs from 'localhost:8000'
    # in redirect_uri. The host-normalization guard redirects to localhost:8000
    # first. We verify the 302 and the Location points to the configured host.
    fake_flow = {
        "state": "abc123",
        "auth_uri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=abc123",
        "redirect_uri": "http://localhost:8000/auth/callback",
        "scope": ["User.Read", "Mail.Read"],
        "code_verifier": "verifier",
        "nonce": "nonce",
        "claims_challenge": None,
    }

    mock_msal = MagicMock()
    mock_msal.initiate_auth_code_flow.return_value = fake_flow

    with patch("app.main._build_msal_app", return_value=mock_msal):
        response = client.get("/auth/login", follow_redirects=False)

    assert response.status_code == 302
    # Host normalization kicks in: redirects to configured host first
    assert "localhost:8000" in response.headers["location"]


def test_auth_login_saves_flow_file(tmp_path):
    # When request host matches configured host, the full MSAL flow runs and
    # the flow file is created. We patch redirect_uri to match testclient host.
    fake_flow = {
        "state": "flow-state-xyz",
        "auth_uri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=flow-state-xyz",
        "redirect_uri": "http://testserver/auth/callback",
        "scope": ["User.Read", "Mail.Read"],
        "code_verifier": "verifier",
        "nonce": "nonce",
        "claims_challenge": None,
    }

    mock_msal = MagicMock()
    mock_msal.initiate_auth_code_flow.return_value = fake_flow

    import app.main as main_module

    with (
        patch("app.main._build_msal_app", return_value=mock_msal),
        patch.object(main_module, "redirect_uri", "http://testserver/auth/callback"),
    ):
        client.get("/auth/login", follow_redirects=False)

    flow_file = tmp_path / "sessions" / "flow-flow-state-xyz.json"
    assert flow_file.exists()
    content = json.loads(flow_file.read_text())
    assert content["auth_flow"]["state"] == "flow-state-xyz"


# ---------------------------------------------------------------------------
# POST /auth/callback
# ---------------------------------------------------------------------------


def test_auth_callback_post_success(tmp_path):
    state = "test-state"
    _write_session_file(
        f"flow-{state}.json",
        {
            "created_at": datetime.now(timezone.utc).isoformat(),
            "auth_flow": {
                "state": state,
                "redirect_uri": "http://localhost:8000/auth/callback",
                "scope": ["User.Read", "Mail.Read"],
                "code_verifier": "verifier",
                "nonce": "nonce",
                "claims_challenge": None,
            },
        },
    )

    fake_token = {
        "access_token": "valid-token",
        "expires_in": 3600,
        "refresh_token": "refresh-token",
        "token_type": "Bearer",
    }
    fake_profile = {
        "id": "user-123",
        "displayName": "Victor",
        "mail": "victor@example.com",
        "userPrincipalName": "victor@example.com",
        "givenName": "Victor",
        "surname": "Ferreira",
        "jobTitle": None,
        "department": None,
        "officeLocation": None,
        "mobilePhone": None,
        "businessPhones": [],
        "preferredLanguage": "pt-BR",
    }
    fake_email = {"id": "e1", "subject": "Test Email", "bodyPreview": "hello"}

    mock_msal = MagicMock()
    mock_msal.acquire_token_by_auth_code_flow.return_value = fake_token

    mock_profile_resp = MagicMock()
    mock_profile_resp.status_code = 200
    mock_profile_resp.json.return_value = fake_profile

    mock_email_resp = MagicMock()
    mock_email_resp.status_code = 200
    mock_email_resp.json.return_value = {"value": [fake_email]}

    with (
        patch("app.main._build_msal_app", return_value=mock_msal),
        patch("app.main.requests.get", side_effect=[mock_profile_resp, mock_email_resp]),
    ):
        response = client.post(
            "/auth/callback",
            content=f"code=test-code&state={state}",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )

    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "Victor" in response.text
    assert "Test Email" in response.text
    assert "local_session_id" in response.cookies


def test_auth_callback_missing_code_or_state():
    response = client.post(
        "/auth/callback",
        content="code=only-code",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    assert response.status_code == 400


def test_auth_callback_flow_not_found():
    response = client.post(
        "/auth/callback",
        content="code=abc&state=nonexistent-state",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    assert response.status_code == 400
    assert "not found" in response.json()["detail"].lower()


def test_auth_callback_msal_token_failure(tmp_path):
    state = "bad-msal-state"
    _write_session_file(
        f"flow-{state}.json",
        {
            "created_at": datetime.now(timezone.utc).isoformat(),
            "auth_flow": {
                "state": state,
                "redirect_uri": "http://localhost:8000/auth/callback",
                "scope": ["User.Read"],
                "code_verifier": "verifier",
                "nonce": "nonce",
                "claims_challenge": None,
            },
        },
    )

    mock_msal = MagicMock()
    mock_msal.acquire_token_by_auth_code_flow.return_value = {
        "error": "invalid_grant",
        "error_description": "AADSTS70000: Something went wrong.",
    }

    with patch("app.main._build_msal_app", return_value=mock_msal):
        response = client.post(
            "/auth/callback",
            content=f"code=bad-code&state={state}",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# GET /profile
# ---------------------------------------------------------------------------


def test_profile_returns_user_data(tmp_path):
    session_id = _create_session(tmp_path)
    fake_profile = {"id": "u1", "displayName": "Victor", "mail": "v@x.com"}

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = fake_profile

    with patch("app.main.requests.get", return_value=mock_resp):
        response = client.get(
            "/profile", cookies={"local_session_id": session_id}
        )

    assert response.status_code == 200
    assert response.json()["displayName"] == "Victor"


def test_profile_without_session_returns_401():
    response = client.get("/profile")
    assert response.status_code == 401


# ---------------------------------------------------------------------------
# GET /profile/export
# ---------------------------------------------------------------------------


def test_profile_export_creates_json_file(tmp_path):
    session_id = _create_session(tmp_path)
    fake_profile = {
        "id": "user-export",
        "displayName": "Victor Export",
        "mail": "v@x.com",
    }

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = fake_profile

    with patch("app.main.requests.get", return_value=mock_resp):
        response = client.get(
            "/profile/export", cookies={"local_session_id": session_id}
        )

    assert response.status_code == 200
    body = response.json()
    assert "json_path" in body
    saved = (tmp_path / body["json_path"]).resolve()
    assert saved.exists()
    content = json.loads(saved.read_text())
    assert content["id"] == "user-export"


def test_profile_export_without_session_returns_401():
    response = client.get("/profile/export")
    assert response.status_code == 401


# ---------------------------------------------------------------------------
# GET /messages/sent/latest
# ---------------------------------------------------------------------------


def test_messages_sent_latest_returns_email(tmp_path):
    session_id = _create_session(tmp_path)
    fake_email = {"id": "e1", "subject": "Hello World", "bodyPreview": "Hi"}

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"value": [fake_email]}

    with patch("app.main.requests.get", return_value=mock_resp):
        response = client.get(
            "/messages/sent/latest", cookies={"local_session_id": session_id}
        )

    assert response.status_code == 200
    assert response.json()["subject"] == "Hello World"


def test_messages_sent_latest_empty_folder(tmp_path):
    session_id = _create_session(tmp_path)

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"value": []}

    with patch("app.main.requests.get", return_value=mock_resp):
        response = client.get(
            "/messages/sent/latest", cookies={"local_session_id": session_id}
        )

    assert response.status_code == 200
    assert response.json() == {"message": "No sent emails found."}


def test_messages_sent_latest_without_session_returns_401():
    response = client.get("/messages/sent/latest")
    assert response.status_code == 401

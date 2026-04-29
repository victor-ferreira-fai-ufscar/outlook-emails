"""
Testes de integracao dos endpoints FastAPI usando TestClient.
Nenhuma chamada real a APIs externas ou ao Azure AD é feita - tudo mockado.
"""

import json
import uuid
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch

import pytest
from fastapi.responses import HTMLResponse
from fastapi.testclient import TestClient

from app.main import app
import app.routes.bot as bot_routes
from app.utils import write_session_file

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
    write_session_file(
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
    assert "text/html" in response.headers["content-type"]
    assert "Onboarding" in response.text


def test_root_includes_quickstart_tutorial_and_links():
    response = client.get("/")
    assert response.status_code == 200
    assert "Como comecar" in response.text
    assert "/auth/login" in response.text
    assert "/docs" in response.text
    assert "wa.me" in response.text


def test_swagger_docs_endpoint_is_available():
    response = client.get("/docs")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "Swagger UI" in response.text


def test_scalar_docs_endpoint_is_available_and_uses_openapi_json():
    response = client.get("/scalar")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "/openapi.json" in response.text


def test_scalar_docs_uses_scalar_fastapi_helper_with_expected_openapi_url():
    fake_html = HTMLResponse("<html>scalar</html>")

    with patch("app.main.get_scalar_api_reference", return_value=fake_html) as mock_ref:
        response = client.get("/scalar")

    assert response.status_code == 200
    assert response.text == "<html>scalar</html>"
    mock_ref.assert_called_once_with(
        openapi_url="/openapi.json",
        title="Outlook Profile Integration - Scalar",
        scalar_proxy_url="https://proxy.scalar.com",
    )


def test_scalar_docs_uses_fallback_openapi_url_when_app_openapi_url_is_none():
    fake_html = HTMLResponse("<html>scalar-fallback</html>")

    with (
        patch("app.main.app.openapi_url", None),
        patch("app.main.get_scalar_api_reference", return_value=fake_html) as mock_ref,
    ):
        response = client.get("/scalar")

    assert response.status_code == 200
    assert response.text == "<html>scalar-fallback</html>"
    mock_ref.assert_called_once_with(
        openapi_url="/openapi.json",
        title="Outlook Profile Integration - Scalar",
        scalar_proxy_url="https://proxy.scalar.com",
    )


def test_openapi_json_endpoint_is_available():
    response = client.get("/openapi.json")
    assert response.status_code == 200
    assert response.headers["content-type"].startswith("application/json")
    assert response.json()["openapi"].startswith("3.")


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

    with patch("app.utils.build_msal_app", return_value=mock_msal):
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

    with (
        patch("app.utils.build_msal_app", return_value=mock_msal),
        patch("app.routes.auth.MS_REDIRECT_URI", "http://testserver/auth/callback"),
    ):
        client.get("/auth/login", follow_redirects=False)

    flow_file = tmp_path / "sessions" / "flow-flow-state-xyz.json"
    assert flow_file.exists()
    content = json.loads(flow_file.read_text())
    assert content["auth_flow"]["state"] == "flow-state-xyz"


def test_auth_login_persists_whatsapp_number_in_flow(tmp_path):
    fake_flow = {
        "state": "flow-whatsapp-xyz",
        "auth_uri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=flow-whatsapp-xyz",
        "redirect_uri": "http://testserver/auth/callback",
        "scope": ["User.Read", "Mail.Read"],
        "code_verifier": "verifier",
        "nonce": "nonce",
        "claims_challenge": None,
    }

    mock_msal = MagicMock()
    mock_msal.initiate_auth_code_flow.return_value = fake_flow

    with (
        patch("app.utils.build_msal_app", return_value=mock_msal),
        patch("app.routes.auth.MS_REDIRECT_URI", "http://testserver/auth/callback"),
    ):
        client.get(
            "/auth/login?whatsapp=5511999999999",
            follow_redirects=False,
        )

    flow_file = tmp_path / "sessions" / "flow-flow-whatsapp-xyz.json"
    assert flow_file.exists()
    content = json.loads(flow_file.read_text())
    assert content["whatsapp_number"] == "5511999999999"


# ---------------------------------------------------------------------------
# POST /auth/callback
# ---------------------------------------------------------------------------


def test_auth_callback_post_success(tmp_path):
    state = "test-state"
    write_session_file(
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
        patch("app.utils.build_msal_app", return_value=mock_msal),
        patch(
            "app.utils.requests.get", side_effect=[mock_profile_resp, mock_email_resp]
        ),
    ):
        response = client.post(
            "/auth/callback",
            content=f"code=test-code&state={state}",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            follow_redirects=False,
        )

    assert response.status_code == 302
    assert response.headers["location"] == "/?welcome=1"
    assert "local_session_id" in response.cookies


def test_auth_callback_links_whatsapp_number_to_authenticated_user(tmp_path):
    state = "test-state-whatsapp"
    write_session_file(
        f"flow-{state}.json",
        {
            "created_at": datetime.now(timezone.utc).isoformat(),
            "whatsapp_number": "5511999999999",
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
    }

    mock_msal = MagicMock()
    mock_msal.acquire_token_by_auth_code_flow.return_value = fake_token

    mock_profile_resp = MagicMock()
    mock_profile_resp.status_code = 200
    mock_profile_resp.json.return_value = fake_profile

    mock_email_resp = MagicMock()
    mock_email_resp.status_code = 200
    mock_email_resp.json.return_value = {"value": []}

    with (
        patch("app.utils.build_msal_app", return_value=mock_msal),
        patch(
            "app.utils.requests.get", side_effect=[mock_profile_resp, mock_email_resp]
        ),
    ):
        response = client.post(
            "/auth/callback",
            content=f"code=test-code&state={state}",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            follow_redirects=False,
        )

    assert response.status_code == 302
    from app.supabase_client import get_user_id_by_whatsapp_number

    assert get_user_id_by_whatsapp_number("5511999999999") == "user-123"


def test_auth_callback_missing_code_or_state():
    response = client.post(
        "/auth/callback",
        content="code=only-code",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    assert response.status_code == 400


def test_auth_callback_alias_missing_code_or_state():
    response = client.post(
        "/callback",
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
    write_session_file(
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

    with patch("app.utils.build_msal_app", return_value=mock_msal):
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

    with patch("app.utils.requests.get", return_value=mock_resp):
        response = client.get("/profile", cookies={"local_session_id": session_id})

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

    with patch("app.utils.requests.get", return_value=mock_resp):
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

    with patch("app.utils.requests.get", return_value=mock_resp):
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

    with patch("app.utils.requests.get", return_value=mock_resp):
        response = client.get(
            "/messages/sent/latest", cookies={"local_session_id": session_id}
        )

    assert response.status_code == 200
    assert response.json() == {"message": "No sent emails found."}


def test_messages_sent_latest_without_session_returns_401():
    response = client.get("/messages/sent/latest")
    assert response.status_code == 401


# ---------------------------------------------------------------------------
# BOT / Teams webhook bootstrap
# ---------------------------------------------------------------------------


def test_bot_health_returns_200_and_channel():
    response = client.get("/bot/health")
    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "ok"
    assert body["channel"] == "microsoft-teams"


def test_bot_messages_help_command_returns_available_commands():
    response = client.post("/bot/messages", json={"type": "message", "text": "ajuda"})
    assert response.status_code == 200
    body = response.json()
    assert "login" in body["commands"]
    assert "logout" in body["commands"]


def test_bot_messages_login_command_returns_login_url():
    response = client.post("/bot/messages", json={"type": "message", "text": "login"})
    assert response.status_code == 200
    body = response.json()
    assert body["type"] == "auth"
    assert body["login_url"]


def test_bot_messages_status_command_returns_not_authenticated_by_default():
    response = client.post("/bot/messages", json={"type": "message", "text": "status"})
    assert response.status_code == 200
    body = response.json()
    assert body["authenticated"] is False


def test_bot_messages_status_with_existing_session_returns_authenticated(tmp_path):
    _create_session(tmp_path)
    fake_profile = {"displayName": "Victor", "mail": "victor@example.com"}

    with patch("app.routes.bot.fetch_outlook_profile", return_value=fake_profile):
        response = client.post(
            "/bot/messages", json={"type": "message", "text": "status"}
        )

    assert response.status_code == 200
    body = response.json()
    assert body["authenticated"] is True
    assert body["user"]["displayName"] == "Victor"


def test_bot_messages_profile_command_uses_existing_session(tmp_path):
    _create_session(tmp_path)
    fake_profile = {"displayName": "Victor", "mail": "victor@example.com"}

    with patch("app.routes.bot.fetch_outlook_profile", return_value=fake_profile):
        response = client.post(
            "/bot/messages", json={"type": "message", "text": "perfil"}
        )

    assert response.status_code == 200
    body = response.json()
    assert body["type"] == "profile"
    assert body["profile"]["mail"] == "victor@example.com"


def test_bot_messages_latest_email_command_uses_existing_session(tmp_path):
    _create_session(tmp_path)
    fake_email = {"subject": "Hello World", "bodyPreview": "preview"}

    with patch("app.routes.bot.fetch_latest_sent_email", return_value=fake_email):
        response = client.post(
            "/bot/messages", json={"type": "message", "text": "ultimo-email"}
        )

    assert response.status_code == 200
    body = response.json()
    assert body["type"] == "latest-email"
    assert body["email"]["subject"] == "Hello World"


def test_bot_messages_unknown_command_returns_hint():
    response = client.post(
        "/bot/messages", json={"type": "message", "text": "comando-invalido"}
    )
    assert response.status_code == 200
    body = response.json()
    assert "ajuda" in body["message"].lower()


def test_bot_conversation_update_returns_welcome_message():
    response = client.post(
        "/bot/messages",
        json={
            "type": "conversationUpdate",
            "channelId": "msteams",
            "from": {"id": "user-123", "name": "Victor"},
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["type"] == "welcome"
    assert "ajuda" in body["message"].lower()


def test_bot_messages_requires_auth_when_enabled(monkeypatch):
    monkeypatch.setattr(bot_routes, "BOT_REQUIRE_AUTH", True)
    monkeypatch.setattr(bot_routes, "BOT_BEARER_TOKEN", "token-test")

    response = client.post("/bot/messages", json={"type": "message", "text": "ajuda"})
    assert response.status_code == 401


def test_bot_messages_accepts_valid_bearer_token_when_auth_enabled(monkeypatch):
    monkeypatch.setattr(bot_routes, "BOT_REQUIRE_AUTH", True)
    monkeypatch.setattr(bot_routes, "BOT_BEARER_TOKEN", "token-test")

    response = client.post(
        "/bot/messages",
        headers={"Authorization": "Bearer token-test"},
        json={"type": "message", "text": "ajuda"},
    )
    assert response.status_code == 200


def test_bot_messages_rejects_invalid_channel(monkeypatch):
    monkeypatch.setattr(bot_routes, "BOT_ALLOWED_CHANNEL", "msteams")

    response = client.post(
        "/bot/messages",
        json={"type": "message", "text": "ajuda", "channelId": "slack"},
    )
    assert response.status_code == 400


# ---------------------------------------------------------------------------
# WhatsApp / Evolution webhook
# ---------------------------------------------------------------------------


def test_whatsapp_webhook_help_command_replies_back(monkeypatch):
    send_mock = MagicMock(return_value={"ok": True, "status_code": 201})
    monkeypatch.setattr(
        "app.routes.whatsapp.send_whatsapp_via_evolution_api", send_mock
    )

    response = client.post(
        "/whatsapp/webhook",
        json={
            "event": "messages.upsert",
            "data": {
                "key": {
                    "remoteJid": "5511999999999@s.whatsapp.net",
                    "fromMe": False,
                },
                "pushName": "Victor",
                "message": {"conversation": "ajuda"},
            },
        },
    )

    assert response.status_code == 200
    assert response.json()["status"] == "ok"
    send_mock.assert_called_once()
    assert send_mock.call_args.kwargs["number"] == "5511999999999"


def test_whatsapp_webhook_login_command_includes_sender_number(monkeypatch):
    send_mock = MagicMock(return_value={"ok": True, "status_code": 201})
    monkeypatch.setattr(
        "app.routes.whatsapp.send_whatsapp_via_evolution_api", send_mock
    )

    response = client.post(
        "/whatsapp/webhook",
        json={
            "event": "messages.upsert",
            "data": {
                "key": {
                    "remoteJid": "5511999999999@s.whatsapp.net",
                    "fromMe": False,
                },
                "message": {"conversation": "login"},
            },
        },
    )

    assert response.status_code == 200
    outbound = send_mock.call_args.kwargs["message"]
    assert "whatsapp=5511999999999" in outbound


def test_whatsapp_webhook_summary_command_sends_summary(monkeypatch):
    monkeypatch.setattr(
        "app.routes.whatsapp.get_access_token_for_user_id", lambda user_id: "token"
    )
    monkeypatch.setattr(
        "app.routes.whatsapp.get_user_id_by_whatsapp_number",
        lambda number: "user-123",
    )
    monkeypatch.setattr(
        "app.routes.whatsapp.get_user_settings",
        lambda user_id: {
            "user_id": user_id,
            "max_emails_in_summary": 10,
            "include_read_emails": False,
            "preferred_channel": "whatsapp",
            "priority_senders": [],
        },
    )
    monkeypatch.setattr(
        "app.routes.whatsapp.fetch_unread_inbox_emails",
        lambda access_token, window_hours=24, max_items=10, include_read=False: [
            {
                "subject": "Urgente",
                "priority": "urgente",
                "sender_name": "Diretoria",
                "received_at": "2026-04-29T10:00:00Z",
            }
        ],
    )
    monkeypatch.setattr(
        "app.routes.whatsapp.format_daily_email_summary",
        lambda emails, top_n=10, include_read=False: "Resumo enviado",
    )
    send_mock = MagicMock(return_value={"ok": True, "status_code": 201})
    monkeypatch.setattr(
        "app.routes.whatsapp.send_whatsapp_via_evolution_api", send_mock
    )

    response = client.post(
        "/whatsapp/webhook",
        json={
            "event": "messages.upsert",
            "data": {
                "key": {
                    "remoteJid": "5511999999999@s.whatsapp.net",
                    "fromMe": False,
                },
                "message": {"extendedTextMessage": {"text": "resumo agora"}},
            },
        },
    )

    assert response.status_code == 200
    assert response.json()["status"] == "ok"
    assert send_mock.call_args.kwargs["message"] == "Resumo enviado"


def test_whatsapp_webhook_ignores_messages_sent_by_self(monkeypatch):
    send_mock = MagicMock()
    monkeypatch.setattr(
        "app.routes.whatsapp.send_whatsapp_via_evolution_api", send_mock
    )

    response = client.post(
        "/whatsapp/webhook",
        json={
            "event": "messages.upsert",
            "data": {
                "key": {
                    "remoteJid": "5511999999999@s.whatsapp.net",
                    "fromMe": True,
                },
                "message": {"conversation": "ajuda"},
            },
        },
    )

    assert response.status_code == 200
    assert response.json()["ignored"] is True
    send_mock.assert_not_called()

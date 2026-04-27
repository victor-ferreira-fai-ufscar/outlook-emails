"""
Testes unitarios das funcoes auxiliares de app.utils.
Nenhuma chamada real a APIs externas é feita aqui - tudo é mockado.
"""

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import app.config as config
import app.utils as utils
from app.utils import (
    authority_url,
    fetch_latest_sent_email,
    fetch_outlook_profile,
    get_latest_local_access_token,
    get_local_access_token,
    read_session_file,
    save_profile_json,
    sessions_dir,
    write_session_file,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

FAKE_PROFILE = {
    "id": "user-123",
    "displayName": "Victor Ferreira",
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

FAKE_EMAIL = {
    "id": "email-abc",
    "subject": "Assunto do email de teste",
    "from": {"emailAddress": {"address": "victor@example.com", "name": "Victor"}},
    "toRecipients": [],
    "bodyPreview": "Corpo do email de teste",
    "sentDateTime": "2026-04-24T10:00:00Z",
    "importance": "normal",
    "isRead": True,
    "webLink": "https://outlook.com/mail/email-abc",
}


# ---------------------------------------------------------------------------
# _authority_url
# ---------------------------------------------------------------------------


def test_authority_url_common():
    # MS_TENANT_ID is "common" by default in config
    assert authority_url() == "https://login.microsoftonline.com/common"


# ---------------------------------------------------------------------------
# _sessions_dir / _write_session_file / _read_session_file
# ---------------------------------------------------------------------------


def test_sessions_dir_creates_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    sessions = sessions_dir()
    assert sessions.is_dir()
    assert sessions.name == "sessions"


def test_write_and_read_session_file(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    payload = {"key": "value", "number": 42}
    write_session_file("test-file.json", payload)

    result = read_session_file("test-file.json")
    assert result == payload


def test_read_session_file_missing_returns_none(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    result = read_session_file("nonexistent.json")
    assert result is None


def test_write_session_file_stores_unicode(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    payload = {"nome": "Victor Ferreira", "emoji": "✅"}
    write_session_file("unicode.json", payload)
    raw = (tmp_path / "sessions" / "unicode.json").read_text(encoding="utf-8")
    assert "Victor Ferreira" in raw
    assert "✅" in raw


# ---------------------------------------------------------------------------
# _save_profile_json
# ---------------------------------------------------------------------------


def test_save_profile_json_creates_file(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    path = save_profile_json(FAKE_PROFILE)
    saved = Path(path)
    assert saved.exists()
    content = json.loads(saved.read_text(encoding="utf-8"))
    assert content["id"] == "user-123"
    assert content["displayName"] == "Victor Ferreira"


def test_save_profile_json_filename_contains_user_id(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    path = save_profile_json(FAKE_PROFILE)
    assert "user-123" in path


def test_save_profile_json_unknown_user_fallback(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    path = save_profile_json({})
    assert "unknown-user" in path


# ---------------------------------------------------------------------------
# _fetch_outlook_profile
# ---------------------------------------------------------------------------


def test_fetch_outlook_profile_success():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = FAKE_PROFILE

    with patch("app.utils.requests.get", return_value=mock_response):
        result = fetch_outlook_profile("fake-token")

    assert result["displayName"] == "Victor Ferreira"
    assert result["mail"] == "victor@example.com"


def test_fetch_outlook_profile_raises_on_error():
    mock_response = MagicMock()
    mock_response.status_code = 401
    mock_response.text = "Unauthorized"

    with patch("app.utils.requests.get", return_value=mock_response):
        from fastapi import HTTPException

        with pytest.raises(HTTPException) as exc_info:
            fetch_outlook_profile("bad-token")

    assert exc_info.value.status_code == 401


def test_fetch_outlook_profile_sends_auth_header():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = FAKE_PROFILE

    with patch("app.utils.requests.get", return_value=mock_response) as mock_get:
        fetch_outlook_profile("my-access-token")

    call_kwargs = mock_get.call_args
    assert call_kwargs.kwargs["headers"]["Authorization"] == "Bearer my-access-token"


# ---------------------------------------------------------------------------
# _fetch_latest_sent_email
# ---------------------------------------------------------------------------


def test_fetch_latest_sent_email_returns_first_message():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": [FAKE_EMAIL]}

    with patch("app.utils.requests.get", return_value=mock_response):
        result = fetch_latest_sent_email("fake-token")

    assert result["id"] == "email-abc"
    assert result["subject"] == "Assunto do email de teste"


def test_fetch_latest_sent_email_empty_inbox():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": []}

    with patch("app.utils.requests.get", return_value=mock_response):
        result = fetch_latest_sent_email("fake-token")

    assert result == {"message": "No sent emails found."}


def test_fetch_latest_sent_email_raises_on_error():
    mock_response = MagicMock()
    mock_response.status_code = 403
    mock_response.text = "Forbidden"

    with patch("app.utils.requests.get", return_value=mock_response):
        from fastapi import HTTPException

        with pytest.raises(HTTPException) as exc_info:
            fetch_latest_sent_email("fake-token")

    assert exc_info.value.status_code == 403


# ---------------------------------------------------------------------------
# _get_local_access_token - token valido
# ---------------------------------------------------------------------------


def test_get_local_access_token_returns_token(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    future = (datetime.now(timezone.utc) + timedelta(hours=1)).isoformat()
    write_session_file(
        "session-test-id.json",
        {"access_token": "valid-token", "expires_at": future},
    )

    mock_request = MagicMock()
    mock_request.cookies = {"local_session_id": "test-id"}

    result = get_local_access_token(mock_request)
    assert result == "valid-token"


def test_get_local_access_token_missing_cookie_raises(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    from fastapi import HTTPException

    mock_request = MagicMock()
    mock_request.cookies = {}

    with pytest.raises(HTTPException) as exc_info:
        get_local_access_token(mock_request)
    assert exc_info.value.status_code == 401


def test_get_local_access_token_missing_file_raises(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    from fastapi import HTTPException

    mock_request = MagicMock()
    mock_request.cookies = {"local_session_id": "ghost-id"}

    with pytest.raises(HTTPException) as exc_info:
        get_local_access_token(mock_request)
    assert exc_info.value.status_code == 401


# ---------------------------------------------------------------------------
# _get_local_access_token - refresh automatico
# ---------------------------------------------------------------------------


def test_get_local_access_token_refreshes_expired_token(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    past = (datetime.now(timezone.utc) - timedelta(seconds=10)).isoformat()
    write_session_file(
        "session-refresh-id.json",
        {
            "access_token": "old-token",
            "expires_at": past,
            "refresh_token": "my-refresh-token",
        },
    )

    mock_msal = MagicMock()
    mock_msal.acquire_token_by_refresh_token.return_value = {
        "access_token": "new-token",
        "expires_in": 3600,
        "refresh_token": "new-refresh-token",
    }

    mock_request = MagicMock()
    mock_request.cookies = {"local_session_id": "refresh-id"}

    with patch("app.utils.build_msal_app", return_value=mock_msal):
        result = get_local_access_token(mock_request)

    assert result == "new-token"

    updated = read_session_file("session-refresh-id.json")
    assert updated["access_token"] == "new-token"
    assert updated["refresh_token"] == "new-refresh-token"


def test_get_local_access_token_expired_no_refresh_token_raises(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    from fastapi import HTTPException

    past = (datetime.now(timezone.utc) - timedelta(seconds=10)).isoformat()
    write_session_file(
        "session-no-refresh.json",
        {"access_token": "old-token", "expires_at": past, "refresh_token": None},
    )

    mock_request = MagicMock()
    mock_request.cookies = {"local_session_id": "no-refresh"}

    with pytest.raises(HTTPException) as exc_info:
        get_local_access_token(mock_request)
    assert exc_info.value.status_code == 401


# ---------------------------------------------------------------------------
# Supabase storage integration (optional backend)
# ---------------------------------------------------------------------------


def test_write_session_file_syncs_to_supabase_when_enabled(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    payload = {"access_token": "token", "expires_at": "2099-01-01T00:00:00+00:00"}

    mock_table = MagicMock()
    mock_upsert = MagicMock()
    mock_upsert.execute.return_value = MagicMock()
    mock_table.upsert.return_value = mock_upsert
    mock_client = MagicMock()
    mock_client.table.return_value = mock_table

    monkeypatch.setattr(utils, "SUPABASE_ENABLED", True)
    monkeypatch.setattr(utils, "get_supabase_client", lambda: mock_client)

    write_session_file("session-abc.json", payload)

    mock_client.table.assert_called_with("sessions")
    assert mock_table.upsert.called


def test_read_session_file_falls_back_to_supabase_when_local_missing(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)

    mock_execute = MagicMock()
    mock_execute.data = [{"payload": {"access_token": "from-supabase"}}]

    mock_limit = MagicMock()
    mock_limit.execute.return_value = mock_execute

    mock_eq = MagicMock()
    mock_eq.limit.return_value = mock_limit

    mock_select = MagicMock()
    mock_select.eq.return_value = mock_eq

    mock_table = MagicMock()
    mock_table.select.return_value = mock_select

    mock_client = MagicMock()
    mock_client.table.return_value = mock_table

    monkeypatch.setattr(utils, "SUPABASE_ENABLED", True)
    monkeypatch.setattr(utils, "get_supabase_client", lambda: mock_client)

    result = read_session_file("session-missing.json")

    assert result == {"access_token": "from-supabase"}
    mock_client.table.assert_called_with("sessions")

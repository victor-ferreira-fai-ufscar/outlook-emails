"""Testes para fluxo de resumo diário e envio no WhatsApp."""

from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.utils import (
    fetch_unread_inbox_emails,
    format_daily_email_summary,
    send_whatsapp_via_evolution_api,
)

client = TestClient(app, raise_server_exceptions=False)


@pytest.fixture(autouse=True)
def use_tmp_dirs(tmp_path, monkeypatch):
    """Isola escrita local em diretório temporário."""
    monkeypatch.chdir(tmp_path)


FAKE_GRAPH_EMAIL = {
    "id": "m1",
    "subject": "Urgente: revisar contrato",
    "from": {"emailAddress": {"name": "Diretoria", "address": "diretoria@empresa.com"}},
    "receivedDateTime": "2026-04-28T12:00:00Z",
    "bodyPreview": "Precisamos da sua aprovacao hoje.",
    "importance": "high",
    "isRead": False,
}


def test_fetch_unread_inbox_emails_returns_normalized_list():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": [FAKE_GRAPH_EMAIL]}

    with patch("app.utils.requests.get", return_value=mock_response):
        result = fetch_unread_inbox_emails("fake-token", window_hours=24, max_items=5)

    assert len(result) == 1
    assert result[0]["subject"] == "Urgente: revisar contrato"
    assert result[0]["sender_email"] == "diretoria@empresa.com"
    assert result[0]["priority"] == "urgente"


def test_fetch_unread_inbox_emails_returns_empty_list_when_no_messages():
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": []}

    with patch("app.utils.requests.get", return_value=mock_response):
        result = fetch_unread_inbox_emails("fake-token")

    assert result == []


def test_fetch_unread_inbox_emails_raises_http_exception_on_graph_error():
    mock_response = MagicMock()
    mock_response.status_code = 403
    mock_response.text = "Forbidden"

    with patch("app.utils.requests.get", return_value=mock_response):
        from fastapi import HTTPException

        with pytest.raises(HTTPException) as exc_info:
            fetch_unread_inbox_emails("fake-token")

    assert exc_info.value.status_code == 403


def test_format_daily_email_summary_includes_counts_and_items():
    summary = format_daily_email_summary(
        [
            {
                "subject": "Urgente: revisar contrato",
                "sender_name": "Diretoria",
                "sender_email": "diretoria@empresa.com",
                "received_at": "2026-04-28T12:00:00Z",
                "preview": "Precisamos da sua aprovacao hoje.",
                "priority": "urgente",
            },
            {
                "subject": "Status semanal",
                "sender_name": "Equipe",
                "sender_email": "equipe@empresa.com",
                "received_at": "2026-04-28T09:00:00Z",
                "preview": "Segue consolidado semanal.",
                "priority": "media",
            },
        ],
        top_n=10,
    )

    assert "Total de nao lidos (24h): 2" in summary
    assert "Urgente: 1" in summary
    assert "Media: 1" in summary
    assert "Diretoria" in summary


def test_format_daily_email_summary_handles_empty_input():
    summary = format_daily_email_summary([])
    assert "Nenhum email novo nao lido nas ultimas 24 horas" in summary


def test_send_whatsapp_via_evolution_api_success(monkeypatch):
    monkeypatch.setattr("app.utils.EVOLUTION_API_URL", "https://evolution.example")
    monkeypatch.setattr("app.utils.EVOLUTION_API_KEY", "abc123")
    monkeypatch.setattr("app.utils.EVOLUTION_INSTANCE", "outlook-bot")
    monkeypatch.setattr("app.utils.EVOLUTION_DEFAULT_NUMBER", "5511999999999")

    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"status": "PENDING"}

    with patch("app.utils.requests.post", return_value=mock_response) as mock_post:
        result = send_whatsapp_via_evolution_api("Resumo diario")

    assert result["ok"] is True
    assert result["status_code"] == 201
    assert (
        mock_post.call_args.args[0]
        == "https://evolution.example/message/sendText/outlook-bot"
    )
    assert mock_post.call_args.kwargs["headers"]["apikey"] == "abc123"
    assert mock_post.call_args.kwargs["json"]["number"] == "5511999999999"
    assert mock_post.call_args.kwargs["json"]["text"] == "Resumo diario"


def test_send_whatsapp_via_evolution_api_raises_when_not_configured(monkeypatch):
    monkeypatch.setattr("app.utils.EVOLUTION_API_URL", "")

    from fastapi import HTTPException

    with pytest.raises(HTTPException) as exc_info:
        send_whatsapp_via_evolution_api("Resumo diario")

    assert exc_info.value.status_code == 500


def test_send_whatsapp_via_evolution_api_raises_on_provider_error(monkeypatch):
    monkeypatch.setattr("app.utils.EVOLUTION_API_URL", "https://evolution.example")
    monkeypatch.setattr("app.utils.EVOLUTION_API_KEY", "abc123")
    monkeypatch.setattr("app.utils.EVOLUTION_INSTANCE", "outlook-bot")
    monkeypatch.setattr("app.utils.EVOLUTION_DEFAULT_NUMBER", "5511999999999")

    mock_response = MagicMock()
    mock_response.status_code = 500
    mock_response.text = "Internal Server Error"

    with patch("app.utils.requests.post", return_value=mock_response):
        from fastapi import HTTPException

        with pytest.raises(HTTPException) as exc_info:
            send_whatsapp_via_evolution_api("Resumo diario")

    assert exc_info.value.status_code == 502


def test_notifications_daily_summary_requires_token_when_enabled(monkeypatch):
    monkeypatch.setattr("app.routes.notifications.NOTIFICATIONS_REQUIRE_AUTH", True)
    monkeypatch.setattr(
        "app.routes.notifications.NOTIFICATIONS_AUTOMATION_TOKEN", "secret-token"
    )

    response = client.post("/notifications/daily-summary")
    assert response.status_code == 401


def test_notifications_daily_summary_returns_summary_and_delivery(monkeypatch):
    monkeypatch.setattr("app.routes.notifications.NOTIFICATIONS_REQUIRE_AUTH", True)
    monkeypatch.setattr(
        "app.routes.notifications.NOTIFICATIONS_AUTOMATION_TOKEN", "secret-token"
    )

    monkeypatch.setattr(
        "app.routes.notifications.get_latest_local_access_token", lambda: "token"
    )
    monkeypatch.setattr(
        "app.routes.notifications.fetch_unread_inbox_emails",
        lambda access_token, window_hours=24, max_items=10, include_read=False: [
            {
                "subject": "Urgente: revisar contrato",
                "sender_name": "Diretoria",
                "sender_email": "diretoria@empresa.com",
                "received_at": datetime.now(timezone.utc).isoformat(),
                "preview": "Precisamos da sua aprovacao hoje.",
                "priority": "urgente",
            }
        ],
    )
    monkeypatch.setattr(
        "app.routes.notifications.format_daily_email_summary",
        lambda emails, top_n=10, include_read=False: "Resumo pronto",
    )
    monkeypatch.setattr(
        "app.routes.notifications.send_whatsapp_via_evolution_api",
        lambda message: {"ok": True, "status_code": 200, "provider_response": "OK"},
    )

    response = client.post(
        "/notifications/daily-summary",
        headers={"Authorization": "Bearer secret-token"},
    )

    assert response.status_code == 200
    body = response.json()
    assert body["emails_processed"] == 1
    assert body["delivery"]["ok"] is True
    assert body["summary_preview"] == "Resumo pronto"


def test_notifications_daily_summary_accepts_local_session_cookie(monkeypatch):
    monkeypatch.setattr("app.routes.notifications.NOTIFICATIONS_REQUIRE_AUTH", True)
    monkeypatch.setattr(
        "app.routes.notifications.NOTIFICATIONS_AUTOMATION_TOKEN", "secret-token"
    )

    monkeypatch.setattr(
        "app.routes.notifications.get_local_access_token", lambda request: "token"
    )
    monkeypatch.setattr(
        "app.routes.notifications.fetch_unread_inbox_emails",
        lambda access_token, window_hours=24, max_items=10, include_read=False: [],
    )
    monkeypatch.setattr(
        "app.routes.notifications.format_daily_email_summary",
        lambda emails, top_n=10, include_read=False: "Sem emails",
    )
    monkeypatch.setattr(
        "app.routes.notifications.send_whatsapp_via_evolution_api",
        lambda message: {"ok": True, "status_code": 200, "provider_response": "OK"},
    )

    response = client.post(
        "/notifications/daily-summary",
        cookies={"local_session_id": "session-id"},
    )

    assert response.status_code == 200
    assert response.json()["delivery"]["ok"] is True


def test_notifications_command_send_summary_now(monkeypatch):
    monkeypatch.setattr("app.routes.notifications.NOTIFICATIONS_REQUIRE_AUTH", True)
    monkeypatch.setattr(
        "app.routes.notifications.get_local_access_token", lambda request: "token"
    )
    monkeypatch.setattr(
        "app.routes.notifications.fetch_unread_inbox_emails",
        lambda access_token, window_hours=24, max_items=10, include_read=False: [
            {"subject": "Assunto", "priority": "media"}
        ],
    )
    monkeypatch.setattr(
        "app.routes.notifications.format_daily_email_summary",
        lambda emails, top_n=10, include_read=False: "Resumo agora",
    )
    monkeypatch.setattr(
        "app.routes.notifications.send_whatsapp_via_evolution_api",
        lambda message: {"ok": True, "status_code": 200, "provider_response": "OK"},
    )

    response = client.post(
        "/notifications/command",
        json={"action": "send_summary_now"},
        cookies={"local_session_id": "session-id"},
    )

    assert response.status_code == 200
    assert response.json()["status"] == "ok"
    assert response.json()["result"]["summary_preview"] == "Resumo agora"


def test_notifications_settings_roundtrip(monkeypatch):
    monkeypatch.setattr("app.routes.notifications.NOTIFICATIONS_REQUIRE_AUTH", True)
    monkeypatch.setattr(
        "app.routes.notifications.get_local_access_token", lambda request: "token"
    )

    monkeypatch.setattr(
        "app.routes.notifications.get_user_settings",
        lambda user_id: {
            "user_id": user_id,
            "max_emails_in_summary": 10,
            "include_read_emails": False,
            "preferred_channel": "whatsapp",
            "priority_senders": [],
        },
    )

    monkeypatch.setattr(
        "app.routes.notifications.save_user_settings",
        lambda user_id, settings: {
            "user_id": user_id,
            "max_emails_in_summary": settings["max_emails_in_summary"],
            "include_read_emails": settings["include_read_emails"],
            "preferred_channel": settings["preferred_channel"],
            "priority_senders": settings["priority_senders"],
        },
    )

    get_response = client.get(
        "/notifications/settings",
        cookies={"local_session_id": "session-id"},
    )
    assert get_response.status_code == 200
    assert get_response.json()["settings"]["max_emails_in_summary"] == 10

    put_response = client.put(
        "/notifications/settings",
        cookies={"local_session_id": "session-id"},
        json={
            "max_emails_in_summary": 5,
            "include_read_emails": True,
            "preferred_channel": "whatsapp",
            "priority_senders": ["chefia@fai.ufscar.br"],
        },
    )
    assert put_response.status_code == 200
    assert put_response.json()["settings"]["max_emails_in_summary"] == 5
    assert put_response.json()["settings"]["include_read_emails"] is True

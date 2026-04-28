"""Testes para fluxo de resumo diário e envio no WhatsApp."""

from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.utils import (
    fetch_unread_inbox_emails,
    format_daily_email_summary,
    send_whatsapp_via_callmebot,
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


def test_send_whatsapp_via_callmebot_success(monkeypatch):
    monkeypatch.setattr(
        "app.utils.CALLMEBOT_API_URL", "https://api.callmebot.com/whatsapp.php"
    )
    monkeypatch.setattr("app.utils.CALLMEBOT_PHONE", "5511999999999")
    monkeypatch.setattr("app.utils.CALLMEBOT_API_KEY", "abc123")

    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.text = "OK"

    with patch("app.utils.requests.get", return_value=mock_response) as mock_get:
        result = send_whatsapp_via_callmebot("Resumo diario")

    assert result["ok"] is True
    assert result["status_code"] == 200
    assert "phone" in mock_get.call_args.kwargs["params"]


def test_send_whatsapp_via_callmebot_raises_when_not_configured(monkeypatch):
    monkeypatch.setattr("app.utils.CALLMEBOT_PHONE", "")

    from fastapi import HTTPException

    with pytest.raises(HTTPException) as exc_info:
        send_whatsapp_via_callmebot("Resumo diario")

    assert exc_info.value.status_code == 500


def test_send_whatsapp_via_callmebot_raises_on_provider_error(monkeypatch):
    monkeypatch.setattr(
        "app.utils.CALLMEBOT_API_URL", "https://api.callmebot.com/whatsapp.php"
    )
    monkeypatch.setattr("app.utils.CALLMEBOT_PHONE", "5511999999999")
    monkeypatch.setattr("app.utils.CALLMEBOT_API_KEY", "abc123")

    mock_response = MagicMock()
    mock_response.status_code = 429
    mock_response.text = "Too Many Requests"

    with patch("app.utils.requests.get", return_value=mock_response):
        from fastapi import HTTPException

        with pytest.raises(HTTPException) as exc_info:
            send_whatsapp_via_callmebot("Resumo diario")

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
        lambda access_token, window_hours=24, max_items=10: [
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
        lambda emails, top_n=10: "Resumo pronto",
    )
    monkeypatch.setattr(
        "app.routes.notifications.send_whatsapp_via_callmebot",
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

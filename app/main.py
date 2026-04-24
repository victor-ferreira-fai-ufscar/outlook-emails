import json
import os
import uuid
from html import escape
from urllib.parse import parse_qs, urlparse
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import msal
import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse

load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["User.Read", "Mail.Read"]

client_id = os.getenv("MS_CLIENT_ID", "")
client_secret = os.getenv("MS_CLIENT_SECRET", "")
tenant_id = os.getenv("MS_TENANT_ID", "common")
redirect_uri = os.getenv("MS_REDIRECT_URI", "http://localhost:8000/auth/callback")
session_secret_key = os.getenv("SESSION_SECRET_KEY", "change-this-in-production")

app = FastAPI(title="Outlook Profile Integration")


def _authority_url() -> str:
    return f"https://login.microsoftonline.com/{tenant_id}"


def _sessions_dir() -> Path:
    directory = Path("sessions")
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def _write_session_file(file_name: str, payload: dict[str, Any]) -> None:
    file_path = _sessions_dir() / file_name
    with file_path.open("w", encoding="utf-8") as fp:
        json.dump(payload, fp, indent=2, ensure_ascii=False)


def _read_session_file(file_name: str) -> dict[str, Any] | None:
    file_path = _sessions_dir() / file_name
    if not file_path.exists():
        return None

    with file_path.open("r", encoding="utf-8") as fp:
        return json.load(fp)


def _build_msal_app() -> msal.ConfidentialClientApplication:
    if not client_id or not client_secret:
        raise HTTPException(
            status_code=500,
            detail="Set MS_CLIENT_ID and MS_CLIENT_SECRET in .env before authenticating.",
        )
    return msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=_authority_url(),
        client_credential=client_secret,
    )


def _fetch_outlook_profile(access_token: str) -> dict[str, Any]:
    response = requests.get(
        f"{GRAPH_BASE_URL}/me",
        headers={"Authorization": f"Bearer {access_token}"},
        params={
            "$select": "id,displayName,mail,userPrincipalName,givenName,surname,jobTitle,department,officeLocation,mobilePhone,businessPhones,preferredLanguage"
        },
        timeout=30,
    )

    if response.status_code >= 400:
        raise HTTPException(status_code=response.status_code, detail=response.text)

    return response.json()


def _fetch_latest_sent_email(access_token: str) -> dict[str, Any]:
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/mailFolders/SentItems/messages",
        headers={"Authorization": f"Bearer {access_token}"},
        params={
            "$top": "1",
            "$orderby": "sentDateTime desc",
            "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,createdDateTime,sentDateTime,receivedDateTime,bodyPreview,conversationId,importance,webLink,isRead",
        },
        timeout=30,
    )

    if response.status_code >= 400:
        raise HTTPException(status_code=response.status_code, detail=response.text)

    payload = response.json()
    messages = payload.get("value", [])
    if not messages:
        return {"message": "No sent emails found."}

    return messages[0]


def _save_profile_json(profile_data: dict[str, Any]) -> str:
    output_dir = Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)

    user_id = profile_data.get("id", "unknown-user")
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output_path = output_dir / f"outlook-profile-{user_id}-{timestamp}.json"

    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(profile_data, fp, indent=2, ensure_ascii=False)

    return str(output_path)


@app.get("/")
def root() -> dict[str, str]:
    return {
        "message": "Outlook integration is running.",
        "next_step": "Open /auth/login to authenticate your account.",
    }


@app.get("/auth/login")
def auth_login(request: Request) -> RedirectResponse:
    configured_host = urlparse(redirect_uri).netloc
    request_host = request.headers.get("host", "")

    # Keep the same host used in redirect_uri, otherwise session cookie is lost on callback.
    if configured_host and request_host and configured_host != request_host:
        normalized_login_url = (
            f"{request.url.scheme}://{configured_host}{request.url.path}"
        )
        return RedirectResponse(url=normalized_login_url, status_code=302)

    msal_app = _build_msal_app()

    auth_flow = msal_app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPES,
        redirect_uri=redirect_uri,
        response_mode="form_post",
    )
    flow_state = auth_flow.get("state")
    if not flow_state:
        raise HTTPException(status_code=500, detail="OAuth state not generated.")

    _write_session_file(
        file_name=f"flow-{flow_state}.json",
        payload={
            "created_at": datetime.now(timezone.utc).isoformat(),
            "auth_flow": auth_flow,
        },
    )

    response = RedirectResponse(url=auth_flow["auth_uri"], status_code=302)
    response.set_cookie(
        key="oauth_state",
        value=flow_state,
        httponly=True,
        samesite="lax",
    )
    return response


@app.api_route("/auth/callback", methods=["GET", "POST"])
async def auth_callback(request: Request) -> HTMLResponse:
    body_params: dict[str, list[str]] = {}
    if request.method == "POST":
        raw_body = (await request.body()).decode("utf-8")
        body_params = parse_qs(raw_body)

    code = request.query_params.get("code") or body_params.get("code", [None])[0]
    state = (
        request.query_params.get("state")
        or body_params.get("state", [None])[0]
        or request.cookies.get("oauth_state")
    )

    if not code or not state:
        raise HTTPException(
            status_code=400, detail="Missing authorization code or state."
        )

    flow_record = _read_session_file(file_name=f"flow-{state}.json")
    if not flow_record:
        raise HTTPException(
            status_code=400,
            detail="Auth flow not found in local storage. Start at /auth/login.",
        )

    auth_flow = flow_record.get("auth_flow")
    if not auth_flow:
        raise HTTPException(
            status_code=400, detail="Invalid auth flow data in sessions."
        )

    msal_app = _build_msal_app()
    token_result = msal_app.acquire_token_by_auth_code_flow(
        auth_code_flow=auth_flow,
        auth_response={"code": code, "state": state},
    )

    if "access_token" not in token_result:
        error_detail = token_result.get("error_description", token_result)
        raise HTTPException(status_code=401, detail=error_detail)

    access_token = token_result["access_token"]
    local_session_id = str(uuid.uuid4())
    _write_session_file(
        file_name=f"session-{local_session_id}.json",
        payload={
            "created_at": datetime.now(timezone.utc).isoformat(),
            "state": state,
            "access_token": access_token,
            "token_result": token_result,
        },
    )

    profile = _fetch_outlook_profile(access_token)
    json_path = _save_profile_json(profile)
    latest_email = _fetch_latest_sent_email(access_token)

    base_url = str(request.base_url).rstrip("/")
    profile_url = f"{base_url}/profile"
    latest_email_url = f"{base_url}/messages/sent/latest"
    export_url = f"{base_url}/profile/export"

    user_name = escape(profile.get("displayName", "User"))
    user_mail = escape(profile.get("mail") or profile.get("userPrincipalName") or "")
    last_subject = escape(latest_email.get("subject", "(sem assunto)"))

    html = f"""
    <html>
      <head>
        <title>Outlook Integration - Auth Success</title>
      </head>
      <body style=\"font-family: Arial, sans-serif; max-width: 760px; margin: 32px auto; line-height: 1.5;\">
        <h1>Autenticacao concluida com sucesso</h1>
        <p><strong>Usuario:</strong> {user_name}</p>
        <p><strong>Email:</strong> {user_mail}</p>
        <p><strong>Ultimo email enviado:</strong> {last_subject}</p>
        <p><strong>JSON de perfil salvo em:</strong> {escape(json_path)}</p>
        <hr />
        <p>Links para teste rapido:</p>
        <ul>
          <li><a href=\"{profile_url}\" target=\"_blank\">Ver perfil (/profile)</a></li>
          <li><a href=\"{latest_email_url}\" target=\"_blank\">Ver ultimo email enviado (/messages/sent/latest)</a></li>
          <li><a href=\"{export_url}\" target=\"_blank\">Exportar perfil novamente (/profile/export)</a></li>
        </ul>
      </body>
    </html>
    """

    response = HTMLResponse(content=html)
    response.set_cookie(
        key="local_session_id",
        value=local_session_id,
        httponly=True,
        samesite="lax",
    )
    return response


def _get_local_access_token(request: Request) -> str:
    local_session_id = request.cookies.get("local_session_id")
    if not local_session_id:
        raise HTTPException(
            status_code=401,
            detail="Not authenticated. Open /auth/login first.",
        )

    session_data = _read_session_file(file_name=f"session-{local_session_id}.json")
    if not session_data:
        raise HTTPException(
            status_code=401,
            detail="Local session not found. Authenticate again at /auth/login.",
        )

    access_token = session_data.get("access_token")
    if not access_token:
        raise HTTPException(
            status_code=401,
            detail="Local access token missing. Authenticate again at /auth/login.",
        )

    return access_token


@app.get("/profile")
def get_profile(request: Request) -> dict[str, Any]:
    access_token = _get_local_access_token(request)
    return _fetch_outlook_profile(access_token)


@app.get("/profile/export")
def export_profile_json(request: Request) -> dict[str, str]:
    access_token = _get_local_access_token(request)
    profile = _fetch_outlook_profile(access_token)
    json_path = _save_profile_json(profile)

    return {"message": "Profile exported successfully.", "json_path": json_path}


@app.get("/messages/sent/latest")
def get_latest_sent_email(request: Request) -> dict[str, Any]:
    access_token = _get_local_access_token(request)
    return _fetch_latest_sent_email(access_token)

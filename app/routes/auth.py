"""
Rotas de autenticacao OAuth2 com Microsoft Graph.
"""

import uuid
from datetime import datetime, timedelta, timezone
from html import escape
from urllib.parse import parse_qs, urlparse
from typing import Any

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.config import GRAPH_SCOPES, MS_REDIRECT_URI
from app.utils import (
    build_msal_app,
    fetch_latest_sent_email,
    fetch_outlook_profile,
    read_session_file,
    save_profile_json,
    write_session_file,
)

router = APIRouter(prefix="/auth", tags=["authentication"])


@router.get("/login")
def auth_login(request: Request) -> RedirectResponse:
    """Inicia fluxo OAuth2 com Microsoft."""
    configured_host = urlparse(MS_REDIRECT_URI).netloc
    request_host = request.headers.get("host", "")

    # Keep the same host used in redirect_uri, otherwise session cookie is lost on callback.
    if configured_host and request_host and configured_host != request_host:
        normalized_login_url = (
            f"{request.url.scheme}://{configured_host}{request.url.path}"
        )
        return RedirectResponse(url=normalized_login_url, status_code=302)

    msal_app = build_msal_app()

    auth_flow = msal_app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPES,
        redirect_uri=MS_REDIRECT_URI,
        response_mode="form_post",
    )
    flow_state = auth_flow.get("state")
    if not flow_state:
        raise HTTPException(status_code=500, detail="OAuth state not generated.")

    write_session_file(
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


@router.api_route("/callback", methods=["GET", "POST"])
async def auth_callback(request: Request) -> HTMLResponse:
    """Processa callback do OAuth2 e retorna pagina de sucesso."""
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

    flow_record = read_session_file(file_name=f"flow-{state}.json")
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

    msal_app = build_msal_app()
    token_result = msal_app.acquire_token_by_auth_code_flow(
        auth_code_flow=auth_flow,
        auth_response={"code": code, "state": state},
    )

    if "access_token" not in token_result:
        error_detail = token_result.get("error_description", token_result)
        raise HTTPException(status_code=401, detail=error_detail)

    access_token = token_result["access_token"]
    expires_in: int = token_result.get("expires_in", 3600)
    expires_at = (
        datetime.now(timezone.utc) + timedelta(seconds=expires_in)
    ).isoformat()

    local_session_id = str(uuid.uuid4())
    write_session_file(
        file_name=f"session-{local_session_id}.json",
        payload={
            "created_at": datetime.now(timezone.utc).isoformat(),
            "state": state,
            "access_token": access_token,
            "expires_at": expires_at,
            "refresh_token": token_result.get("refresh_token"),
            "token_result": token_result,
        },
    )

    profile = fetch_outlook_profile(access_token)
    json_path = save_profile_json(profile)
    latest_email = fetch_latest_sent_email(access_token)

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

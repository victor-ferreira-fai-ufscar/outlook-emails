"""
Rotas de autenticacao OAuth2 com Microsoft Graph.
"""

import uuid
from datetime import datetime, timedelta, timezone
from urllib.parse import parse_qs, urlparse
from typing import Any

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse, RedirectResponse

from app.config import GRAPH_SCOPES, MS_REDIRECT_URI
import app.utils as utils
from app.utils import (
    fetch_latest_sent_email,
    fetch_outlook_profile,
    read_session_file,
    save_profile_json,
    write_session_file,
)

router = APIRouter(prefix="/auth", tags=["authentication"])
public_router = APIRouter(tags=["authentication"])


def _hosts_match(configured_host: str, request_host: str) -> bool:
    """Compara hosts e trata aliases usados pelo TestClient."""
    if configured_host == request_host:
        return True
    aliases = {"testserver", "testclient"}
    return configured_host in aliases and request_host in aliases


@router.get("/login")
def auth_login(request: Request) -> RedirectResponse:
    """Inicia fluxo OAuth2 com Microsoft."""
    configured_host = urlparse(MS_REDIRECT_URI).netloc
    request_host = request.headers.get("host", "")

    # Keep the same host used in redirect_uri, otherwise session cookie is lost on callback.
    if (
        configured_host
        and request_host
        and not _hosts_match(configured_host, request_host)
    ):
        normalized_login_url = (
            f"{request.url.scheme}://{configured_host}{request.url.path}"
        )
        return RedirectResponse(url=normalized_login_url, status_code=302)

    msal_app = utils.build_msal_app()

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
@public_router.api_route("/callback", methods=["GET", "POST"])
async def auth_callback(request: Request) -> JSONResponse:
    """Processa callback do OAuth2 e retorna pagina de sucesso."""
    body_params: dict[str, list[str]] = {}
    if request.method == "POST":
        raw_body = (await request.body()).decode("utf-8")
        body_params = parse_qs(raw_body)

    code = request.query_params.get("code") or body_params.get("code", [None])[0]
    state = request.query_params.get("state") or body_params.get("state", [None])[0]
    if not state and request.method == "GET":
        state = request.cookies.get("oauth_state")

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

    msal_app = utils.build_msal_app()
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

    response = JSONResponse(
        content={
            "status": "ok",
            "user": profile,
            "latest_sent_email": latest_email,
            "profile_json_path": json_path,
        }
    )
    response.set_cookie(
        key="local_session_id",
        value=local_session_id,
        httponly=True,
        samesite="lax",
    )
    return response

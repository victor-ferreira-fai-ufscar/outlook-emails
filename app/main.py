import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import msal
import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import RedirectResponse
from starlette.middleware.sessions import SessionMiddleware

load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["User.Read", "Mail.Read"]

client_id = os.getenv("MS_CLIENT_ID", "")
client_secret = os.getenv("MS_CLIENT_SECRET", "")
tenant_id = os.getenv("MS_TENANT_ID", "common")
redirect_uri = os.getenv("MS_REDIRECT_URI", "http://localhost:8000/auth/callback")
session_secret_key = os.getenv("SESSION_SECRET_KEY", "change-this-in-production")

app = FastAPI(title="Outlook Profile Integration")
app.add_middleware(SessionMiddleware, secret_key=session_secret_key)


def _authority_url() -> str:
    return f"https://login.microsoftonline.com/{tenant_id}"


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
    msal_app = _build_msal_app()

    auth_flow = msal_app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPES,
        redirect_uri=redirect_uri,
    )
    request.session["auth_flow"] = auth_flow

    return RedirectResponse(url=auth_flow["auth_uri"], status_code=302)


@app.get("/auth/callback")
def auth_callback(
    request: Request, code: str | None = None, state: str | None = None
) -> dict[str, Any]:
    auth_flow = request.session.get("auth_flow")
    if not auth_flow:
        raise HTTPException(
            status_code=400,
            detail="Auth flow not found in session. Start at /auth/login.",
        )

    if not code or not state:
        raise HTTPException(
            status_code=400, detail="Missing authorization code or state."
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
    request.session["access_token"] = access_token

    profile = _fetch_outlook_profile(access_token)
    json_path = _save_profile_json(profile)

    return {
        "message": "Authentication successful. Outlook profile collected.",
        "json_path": json_path,
        "profile": profile,
    }


@app.get("/profile")
def get_profile(request: Request) -> dict[str, Any]:
    access_token = request.session.get("access_token")
    if not access_token:
        raise HTTPException(
            status_code=401, detail="Not authenticated. Open /auth/login first."
        )

    return _fetch_outlook_profile(access_token)


@app.get("/profile/export")
def export_profile_json(request: Request) -> dict[str, str]:
    access_token = request.session.get("access_token")
    if not access_token:
        raise HTTPException(
            status_code=401, detail="Not authenticated. Open /auth/login first."
        )

    profile = _fetch_outlook_profile(access_token)
    json_path = _save_profile_json(profile)

    return {"message": "Profile exported successfully.", "json_path": json_path}


@app.get("/messages/sent/latest")
def get_latest_sent_email(request: Request) -> dict[str, Any]:
    access_token = request.session.get("access_token")
    if not access_token:
        raise HTTPException(
            status_code=401, detail="Not authenticated. Open /auth/login first."
        )

    return _fetch_latest_sent_email(access_token)

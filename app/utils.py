"""
Funcoes auxiliares para persistencia local, Graph API e MSAL.
"""

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import msal
import requests
from fastapi import HTTPException

from app.config import (
    GRAPH_BASE_URL,
    GRAPH_SCOPES,
    MS_CLIENT_ID,
    MS_CLIENT_SECRET,
    MS_TENANT_ID,
    SUPABASE_ENABLED,
)
from app.supabase_client import get_supabase_client


def authority_url() -> str:
    """Constroi URL de autoridade do Azure AD."""
    return f"https://login.microsoftonline.com/{MS_TENANT_ID}"


def sessions_dir() -> Path:
    """Cria e retorna diretorio de sessoes locais."""
    directory = Path("sessions")
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def write_session_file(file_name: str, payload: dict[str, Any]) -> None:
    """Salva dados de sessao em arquivo JSON local."""
    file_path = sessions_dir() / file_name
    with file_path.open("w", encoding="utf-8") as fp:
        json.dump(payload, fp, indent=2, ensure_ascii=False)

    _sync_session_to_supabase(file_name=file_name, payload=payload)


def read_session_file(file_name: str) -> dict[str, Any] | None:
    """Le dados de sessao de arquivo JSON local."""
    file_path = sessions_dir() / file_name
    if not file_path.exists():
        return _read_session_from_supabase(file_name=file_name)

    with file_path.open("r", encoding="utf-8") as fp:
        return json.load(fp)


def build_msal_app() -> msal.ConfidentialClientApplication:
    """Constroi cliente MSAL para OAuth2."""
    if not MS_CLIENT_ID or not MS_CLIENT_SECRET:
        raise HTTPException(
            status_code=500,
            detail="Set MS_CLIENT_ID and MS_CLIENT_SECRET in .env before authenticating.",
        )
    return msal.ConfidentialClientApplication(
        client_id=MS_CLIENT_ID,
        authority=authority_url(),
        client_credential=MS_CLIENT_SECRET,
    )


def fetch_outlook_profile(access_token: str) -> dict[str, Any]:
    """Busca dados do perfil do usuario via Microsoft Graph."""
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


def fetch_latest_sent_email(access_token: str) -> dict[str, Any]:
    """Busca ultimo email enviado via Microsoft Graph."""
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


def save_profile_json(profile_data: dict[str, Any]) -> str:
    """Salva snapshot do perfil em arquivo JSON em data/."""
    output_dir = Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)

    user_id = profile_data.get("id", "unknown-user")
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output_path = output_dir / f"outlook-profile-{user_id}-{timestamp}.json"

    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(profile_data, fp, indent=2, ensure_ascii=False)

    _sync_profile_to_supabase(
        user_id=user_id,
        path=str(output_path),
        payload=profile_data,
    )

    return str(output_path)


def _sync_session_to_supabase(file_name: str, payload: dict[str, Any]) -> None:
    """Sincroniza sessão em Supabase sem interromper fluxo local."""
    if not SUPABASE_ENABLED:
        return

    try:
        client = get_supabase_client()
        client.table("sessions").upsert(
            {
                "file_name": file_name,
                "payload": payload,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            },
            on_conflict="file_name",
        ).execute()
    except Exception:
        # Mantém operação local mesmo se Supabase estiver indisponível.
        return


def _read_session_from_supabase(file_name: str) -> dict[str, Any] | None:
    """Busca sessão no Supabase quando arquivo local não existe."""
    if not SUPABASE_ENABLED:
        return None

    try:
        client = get_supabase_client()
        response = (
            client.table("sessions")
            .select("payload")
            .eq("file_name", file_name)
            .limit(1)
            .execute()
        )
        rows = response.data or []
        if not rows:
            return None
        return rows[0].get("payload")
    except Exception:
        return None


def _sync_profile_to_supabase(user_id: str, path: str, payload: dict[str, Any]) -> None:
    """Sincroniza snapshot de perfil em Supabase sem quebrar execução local."""
    if not SUPABASE_ENABLED:
        return

    try:
        client = get_supabase_client()
        client.table("profiles").insert(
            {
                "user_id": user_id,
                "path": path,
                "payload": payload,
                "created_at": datetime.now(timezone.utc).isoformat(),
            }
        ).execute()
    except Exception:
        return


def get_local_access_token(request) -> str:
    """
    Retorna access token da sessao local, com refresh automatico se expirado.
    """
    from fastapi import Request

    local_session_id = request.cookies.get("local_session_id")
    if not local_session_id:
        raise HTTPException(
            status_code=401,
            detail="Not authenticated. Open /auth/login first.",
        )

    session_data = read_session_file(file_name=f"session-{local_session_id}.json")
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

    # Refresh token if expired or about to expire (within 60 seconds).
    expires_at_raw = session_data.get("expires_at")
    if expires_at_raw:
        expires_at = datetime.fromisoformat(expires_at_raw)
        if datetime.now(timezone.utc) >= expires_at - timedelta(seconds=60):
            refresh_token = session_data.get("refresh_token")
            if not refresh_token:
                raise HTTPException(
                    status_code=401,
                    detail="Session expired and no refresh token available. Authenticate again at /auth/login.",
                )

            msal_app = build_msal_app()
            refreshed = msal_app.acquire_token_by_refresh_token(
                refresh_token=refresh_token,
                scopes=GRAPH_SCOPES,
            )

            if "access_token" not in refreshed:
                raise HTTPException(
                    status_code=401,
                    detail="Token refresh failed. Authenticate again at /auth/login.",
                )

            access_token = refreshed["access_token"]
            new_expires_in: int = refreshed.get("expires_in", 3600)
            session_data["access_token"] = access_token
            session_data["expires_at"] = (
                datetime.now(timezone.utc) + timedelta(seconds=new_expires_in)
            ).isoformat()
            session_data["refresh_token"] = refreshed.get(
                "refresh_token", refresh_token
            )
            write_session_file(
                file_name=f"session-{local_session_id}.json",
                payload=session_data,
            )

    return access_token


def get_latest_local_access_token() -> str:
    """Retorna access token da sessao local mais recente (sem cookie)."""
    session_files = sorted(
        sessions_dir().glob("session-*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if not session_files:
        raise HTTPException(
            status_code=401,
            detail="Nenhuma sessao local encontrada. Faça login em /auth/login.",
        )

    latest_file = session_files[0]
    session_data = read_session_file(latest_file.name)
    if not session_data:
        raise HTTPException(
            status_code=401,
            detail="Sessao local invalida. Faça login novamente em /auth/login.",
        )

    access_token = session_data.get("access_token")
    if not access_token:
        raise HTTPException(
            status_code=401,
            detail="Token de acesso ausente na sessao local.",
        )

    expires_at_raw = session_data.get("expires_at")
    if expires_at_raw:
        expires_at = datetime.fromisoformat(expires_at_raw)
        if datetime.now(timezone.utc) >= expires_at - timedelta(seconds=60):
            refresh_token = session_data.get("refresh_token")
            if not refresh_token:
                raise HTTPException(
                    status_code=401,
                    detail="Sessao expirada e sem refresh token. Faça login novamente.",
                )

            msal_app = build_msal_app()
            refreshed = msal_app.acquire_token_by_refresh_token(
                refresh_token=refresh_token,
                scopes=GRAPH_SCOPES,
            )

            if "access_token" not in refreshed:
                raise HTTPException(
                    status_code=401,
                    detail="Falha ao atualizar token. Faça login novamente.",
                )

            access_token = refreshed["access_token"]
            new_expires_in: int = refreshed.get("expires_in", 3600)
            session_data["access_token"] = access_token
            session_data["expires_at"] = (
                datetime.now(timezone.utc) + timedelta(seconds=new_expires_in)
            ).isoformat()
            session_data["refresh_token"] = refreshed.get(
                "refresh_token", refresh_token
            )
            write_session_file(file_name=latest_file.name, payload=session_data)

    return access_token

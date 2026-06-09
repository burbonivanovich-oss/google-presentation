from __future__ import annotations

import json
import os
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

SECRETS_DIR = Path(__file__).resolve().parent.parent / "secrets"
CLIENT_SECRET = SECRETS_DIR / "client_secret.json"
TOKEN_FILE = SECRETS_DIR / "token.json"

ENV_CLIENT_JSON = "GOOGLE_OAUTH_CLIENT_JSON"
ENV_TOKEN_JSON = "GOOGLE_OAUTH_TOKEN_JSON"


def get_credentials() -> Credentials:
    """OAuth-credentials от имени пользователя.

    Источники, в порядке приоритета:
      1) env GOOGLE_OAUTH_TOKEN_JSON — целый token.json (для CI).
      2) secrets/token.json — после локального run_local_server.
    Если токена нет / просрочен — интерактивный браузерный логин с
    client_secret из env GOOGLE_OAUTH_CLIENT_JSON или secrets/client_secret.json.
    """
    creds = _load_token()

    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _persist_token(creds)
        return creds

    client_config = _load_client_config()
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    creds = flow.run_local_server(port=0)
    _persist_token(creds)
    return creds


def _load_token() -> Credentials | None:
    raw = os.environ.get(ENV_TOKEN_JSON)
    if raw:
        return Credentials.from_authorized_user_info(json.loads(raw), SCOPES)
    if TOKEN_FILE.exists():
        return Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    return None


def _load_client_config() -> dict:
    raw = os.environ.get(ENV_CLIENT_JSON)
    if raw:
        return json.loads(raw)
    if CLIENT_SECRET.exists():
        return json.loads(CLIENT_SECRET.read_text())
    raise FileNotFoundError(
        f"Нужен OAuth client_secret. Положите его в {CLIENT_SECRET} "
        f"или передайте через env {ENV_CLIENT_JSON}."
    )


def _persist_token(creds: Credentials) -> None:
    try:
        SECRETS_DIR.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(creds.to_json())
    except OSError:
        pass

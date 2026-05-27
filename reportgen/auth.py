from __future__ import annotations

import json
import os
from pathlib import Path

from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

SECRETS_DIR = Path(__file__).resolve().parent.parent / "secrets"
SERVICE_ACCOUNT_FILE = SECRETS_DIR / "service_account.json"

ENV_SERVICE_ACCOUNT_JSON = "GCP_SERVICE_ACCOUNT_JSON"


def get_credentials() -> Credentials:
    """Возвращает credentials сервис-аккаунта.

    Источники, в порядке приоритета:
      1) env GCP_SERVICE_ACCOUNT_JSON — полный JSON ключа (для GitHub Actions).
      2) Файл secrets/service_account.json (для локального запуска).
    """
    raw = os.environ.get(ENV_SERVICE_ACCOUNT_JSON)
    if raw:
        info = json.loads(raw)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    if SERVICE_ACCOUNT_FILE.exists():
        return Credentials.from_service_account_file(
            str(SERVICE_ACCOUNT_FILE), scopes=SCOPES
        )

    raise FileNotFoundError(
        f"Нужен JSON-ключ сервис-аккаунта: положите его в {SERVICE_ACCOUNT_FILE} "
        f"или передайте через env {ENV_SERVICE_ACCOUNT_JSON}."
    )


def service_account_email() -> str | None:
    """Email сервис-аккаунта — удобно показать пользователю, что шарить."""
    raw = os.environ.get(ENV_SERVICE_ACCOUNT_JSON)
    if raw:
        return json.loads(raw).get("client_email")
    if SERVICE_ACCOUNT_FILE.exists():
        return json.loads(SERVICE_ACCOUNT_FILE.read_text()).get("client_email")
    return None

from __future__ import annotations

import re

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

MIME_SHEET = "application/vnd.google-apps.spreadsheet"
MIME_SLIDES = "application/vnd.google-apps.presentation"


class DriveClient:
    def __init__(self, creds: Credentials):
        self._drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    def find_in_folder(
        self,
        folder_id: str,
        name_pattern: str,
        mime_type: str | None = None,
    ) -> dict:
        """Ищет файл в папке по regex/подстроке имени, возвращает самый свежий.

        name_pattern сравнивается как регулярка (re.search, IGNORECASE) с именами
        файлов внутри папки. Поднимает FileNotFoundError, если совпадений нет.
        """
        q_parts = [f"'{folder_id}' in parents", "trashed = false"]
        if mime_type:
            q_parts.append(f"mimeType = '{mime_type}'")
        q = " and ".join(q_parts)

        resp = (
            self._drive.files()
            .list(
                q=q,
                fields="files(id,name,mimeType,modifiedTime)",
                orderBy="modifiedTime desc",
                pageSize=200,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files = resp.get("files", [])
        rx = re.compile(name_pattern, re.IGNORECASE)
        matches = [f for f in files if rx.search(f["name"])]
        if not matches:
            names = ", ".join(f["name"] for f in files[:10]) or "(пусто)"
            raise FileNotFoundError(
                f"В папке {folder_id} нет файла по паттерну {name_pattern!r}. "
                f"Видимые файлы: {names}"
            )
        return matches[0]

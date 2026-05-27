from __future__ import annotations

import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


class SheetsClient:
    def __init__(self, creds: Credentials):
        self._svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    def read_range(self, spreadsheet_id: str, a1_range: str) -> list[list[str]]:
        resp = (
            self._svc.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=a1_range)
            .execute()
        )
        return resp.get("values", [])

    def read_table(self, spreadsheet_id: str, a1_range: str) -> pd.DataFrame:
        """Читает диапазон, где первая строка — заголовки."""
        rows = self.read_range(spreadsheet_id, a1_range)
        if not rows:
            return pd.DataFrame()
        header, *data = rows
        width = len(header)
        normalized = [r + [""] * (width - len(r)) for r in data]
        return pd.DataFrame(normalized, columns=header)

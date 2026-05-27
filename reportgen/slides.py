from __future__ import annotations

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


class SlidesClient:
    """Тонкая обёртка над Slides + Drive API.

    Сценарий:
      1) copy_presentation(template_id, title) → копия шаблона
      2) replace_placeholders({"{{kpi_revenue}}": "12.3M ₽", ...})
      3) duplicate_slide(slide_id) для каждого дополнительного инсайта
    """

    def __init__(self, creds: Credentials):
        self._slides = build("slides", "v1", credentials=creds, cache_discovery=False)
        self._drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    def copy_presentation(self, template_id: str, title: str) -> str:
        copy = self._drive.files().copy(fileId=template_id, body={"name": title}).execute()
        return copy["id"]

    def get_presentation(self, presentation_id: str) -> dict:
        return self._slides.presentations().get(presentationId=presentation_id).execute()

    def replace_placeholders(self, presentation_id: str, mapping: dict[str, str]) -> None:
        if not mapping:
            return
        requests = [
            {
                "replaceAllText": {
                    "containsText": {"text": placeholder, "matchCase": True},
                    "replaceText": value,
                }
            }
            for placeholder, value in mapping.items()
        ]
        self._slides.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests}
        ).execute()

    def duplicate_slide(self, presentation_id: str, slide_id: str) -> str:
        resp = (
            self._slides.presentations()
            .batchUpdate(
                presentationId=presentation_id,
                body={"requests": [{"duplicateObject": {"objectId": slide_id}}]},
            )
            .execute()
        )
        return resp["replies"][0]["duplicateObject"]["objectId"]

    def delete_slide(self, presentation_id: str, slide_id: str) -> None:
        self._slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"deleteObject": {"objectId": slide_id}}]},
        ).execute()

    def presentation_url(self, presentation_id: str) -> str:
        return f"https://docs.google.com/presentation/d/{presentation_id}/edit"

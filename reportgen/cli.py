from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console

from .auth import get_credentials, service_account_email
from .config import ReportConfig, SheetSource
from .drive import MIME_SHEET, MIME_SLIDES, DriveClient
from .insights import qoq_changes, roas_below_benchmark, sigma_anomalies
from .sheets import SheetsClient
from .slides import SlidesClient

app = typer.Typer(help="Генератор квартальных Google Slides отчётов из Google Sheets.")
console = Console()


@app.command()
def whoami() -> None:
    """Показать email сервис-аккаунта — его нужно добавить в шаринг папки Drive."""
    get_credentials()  # упадёт с понятной ошибкой, если ключа нет
    email = service_account_email()
    console.print(f"Service account email: [bold]{email}[/bold]")
    console.print("Расшарьте на этот email вашу папку Drive (роль Editor).")


@app.command()
def generate(
    config_path: Path = typer.Option(..., "--config", "-c", help="YAML конфиг отчёта"),
    current_period: str = typer.Option(..., "--period", help="например Q1-2026"),
    previous_period: str = typer.Option(..., "--prev", help="например Q4-2025"),
) -> None:
    """Сгенерировать презентацию по конфигу."""
    cfg = ReportConfig.load(config_path)
    creds = get_credentials()
    sheets = SheetsClient(creds)
    slides = SlidesClient(creds)
    drive = DriveClient(creds)

    console.print(f"[bold]Отчёт:[/bold] {cfg.name} ({previous_period} → {current_period})")

    # 1. Читаем данные. spreadsheet_id может быть задан явно или найден по
    #    имени в расшаренной папке (folder_id).
    data = {}
    for key, src in cfg.sources.items():
        sid = _resolve_spreadsheet(src, cfg.folder_id, drive)
        data[key] = sheets.read_table(sid, src.range)
        console.print(f"  • {key}: {data[key].shape[0]} строк из {sid}")
    console.print(f"Прочитано таблиц: {len(data)}")

    template_id = cfg.presentation_template_id or drive.find_in_folder(
        cfg.folder_id, cfg.presentation_template_name, MIME_SLIDES
    )["id"]

    # 2. Ищем инсайты. Ожидаем источник 'channels' с колонками:
    #    period, channel, spend, revenue, roas, leads
    insights = []
    if "channels" in data and not data["channels"].empty:
        ch = data["channels"]
        for metric in ("spend", "revenue", "leads"):
            if metric in ch.columns:
                ch[metric] = ch[metric].apply(_to_num)
                insights += qoq_changes(
                    ch,
                    entity_col="channel",
                    metric_col=metric,
                    current_period=current_period,
                    previous_period=previous_period,
                    threshold_pct=cfg.insights.qoq_threshold_pct,
                )
        cur = ch[ch.get("period") == current_period].copy()
        if not cur.empty:
            if "spend" in cur.columns:
                insights += sigma_anomalies(
                    cur, entity_col="channel", metric_col="spend", sigma=cfg.insights.sigma
                )
            if "roas" in cur.columns:
                cur["roas"] = cur["roas"].apply(_to_num)
                insights += roas_below_benchmark(
                    cur, entity_col="channel", benchmark=cfg.insights.roas_benchmark
                )

    console.print(f"Найдено инсайтов: {len(insights)}")

    # 3. Копируем шаблон в ту же папку, иначе у пользователя не будет доступа
    #    к копии (она окажется в "My Drive" сервис-аккаунта).
    title = f"{cfg.name} — {current_period}"
    pres_id = slides.copy_presentation(template_id, title, parent_folder_id=cfg.folder_id)
    console.print(f"Создана копия шаблона: {slides.presentation_url(pres_id)}")

    # 4. Дублируем insight-слайд и подставляем тексты в каждый дубликат.
    #    Сначала дублируем (placeholder остаётся {{insight_headline}}), потом
    #    замена выполняется один раз — но т.к. в копиях значения должны
    #    отличаться, делаем уникальные маркеры через replaceAllText по индексу.
    if insights:
        # для каждого инсайта создаём слайд с уникальными маркерами
        for idx, ins in enumerate(insights):
            new_id = slides.duplicate_slide(pres_id, cfg.insight_slide_id)
            # вместо общего {{insight_headline}} ставим маркер с индексом, который
            # есть только в новом слайде. Решение: дублирующий слайд содержит
            # те же плейсхолдеры, поэтому используем pageObjectIds в batchUpdate.
            # Проще — обновить текст в конкретном слайде через replaceAllText
            # с pageObjectIds=[new_id].
            _replace_in_pages(
                slides,
                pres_id,
                [new_id],
                {
                    "{{insight_headline}}": ins.headline,
                    "{{insight_detail}}": ins.detail,
                },
            )
        # удаляем исходный insight-слайд (он остался пустым шаблоном)
        slides.delete_slide(pres_id, cfg.insight_slide_id)

    # 5. Подставляем глобальные плейсхолдеры.
    mapping = {
        "{{period}}": current_period,
        "{{previous_period}}": previous_period,
        "{{report_name}}": cfg.name,
        **cfg.static_placeholders,
    }
    slides.replace_placeholders(pres_id, mapping)

    console.print(f"[green]Готово[/green] → {slides.presentation_url(pres_id)}")


def _resolve_spreadsheet(src: SheetSource, folder_id: str | None, drive: DriveClient) -> str:
    if src.spreadsheet_id:
        return src.spreadsheet_id
    return drive.find_in_folder(folder_id, src.name_pattern, MIME_SHEET)["id"]


def _replace_in_pages(
    slides: SlidesClient,
    presentation_id: str,
    page_ids: list[str],
    mapping: dict[str, str],
) -> None:
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": placeholder, "matchCase": True},
                "replaceText": value,
                "pageObjectIds": page_ids,
            }
        }
        for placeholder, value in mapping.items()
    ]
    slides._slides.presentations().batchUpdate(  # noqa: SLF001
        presentationId=presentation_id, body={"requests": requests}
    ).execute()


def _to_num(x):
    try:
        return float(str(x).replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except (ValueError, TypeError):
        return float("nan")


if __name__ == "__main__":
    app()

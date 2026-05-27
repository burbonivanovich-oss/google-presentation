from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from .auth import get_credentials, service_account_email
from .config import ReportConfig, SheetSource
from .drive import MIME_SHEET, MIME_SLIDES, DriveClient
from .insights import qoq_changes, roas_below_benchmark, sigma_anomalies
from .sheets import SheetsClient
from .slides import SlidesClient

app = typer.Typer(help="Генератор квартальных Google Slides отчётов из Google Sheets.")
console = Console()

INSIGHT_HEADLINE = "{{insight_headline}}"
INSIGHT_DETAIL = "{{insight_detail}}"


@app.command()
def whoami() -> None:
    """Показать email сервис-аккаунта — его нужно добавить в шаринг папки Drive."""
    get_credentials()
    email = service_account_email()
    console.print(f"Service account email: [bold]{email}[/bold]")
    console.print("Расшарьте на этот email вашу папку Drive (роль Editor).")


@app.command(name="list-folder")
def list_folder(folder_id: str = typer.Argument(..., help="ID папки в Drive")) -> None:
    """Показать всё, что сервис-аккаунт видит в папке."""
    creds = get_credentials()
    drive = DriveClient(creds)
    resp = (
        drive._drive.files()  # noqa: SLF001
        .list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="files(id,name,mimeType,modifiedTime)",
            orderBy="modifiedTime desc",
            pageSize=200,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = resp.get("files", [])
    if not files:
        console.print("[yellow]Папка пуста или нет доступа.[/yellow]")
        return
    table = Table(title=f"Файлы в папке {folder_id}")
    table.add_column("Имя"); table.add_column("Тип"); table.add_column("ID", style="dim")
    for f in files:
        kind = {
            "application/vnd.google-apps.spreadsheet": "Sheet",
            "application/vnd.google-apps.presentation": "Slides",
            "application/vnd.google-apps.folder": "Folder",
        }.get(f["mimeType"], f["mimeType"].split(".")[-1])
        table.add_row(f["name"], kind, f["id"])
    console.print(table)


@app.command(name="list-slides")
def list_slides(presentation_id: str = typer.Argument(..., help="ID презентации")) -> None:
    """Показать слайды и их превью — удобно, чтобы найти objectId insight-слайда."""
    creds = get_credentials()
    slides = SlidesClient(creds)
    table = Table(title=f"Слайды презентации {presentation_id}")
    table.add_column("objectId", style="dim"); table.add_column("Превью")
    for oid, preview in slides.list_slides(presentation_id):
        table.add_row(oid, preview or "(пусто)")
    console.print(table)


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

    # 1. Данные
    data: dict = {}
    for key, src in cfg.sources.items():
        try:
            sid = _resolve_spreadsheet(src, cfg.folder_id, drive)
            data[key] = sheets.read_table(sid, src.range)
            console.print(f"  • {key}: {data[key].shape[0]} строк")
        except Exception as e:  # noqa: BLE001
            console.print(f"  [yellow]пропуск {key}: {e}[/yellow]")

    # 2. Инсайты (только из источника 'channels', если есть)
    insights = _collect_insights(data, cfg, current_period, previous_period)
    console.print(f"Найдено инсайтов: {len(insights)}")

    # 3. Копия шаблона
    template_id = cfg.presentation_template_id or drive.find_in_folder(
        cfg.folder_id, cfg.presentation_template_name, MIME_SLIDES
    )["id"]
    title = f"{cfg.name} — {current_period}"
    pres_id = slides.copy_presentation(template_id, title, parent_folder_id=cfg.folder_id)
    console.print(f"Копия шаблона: {slides.presentation_url(pres_id)}")

    # 4. Insight-слайды
    insight_slide_id = cfg.insight_slide_id or slides.find_slide_with_text(
        pres_id, INSIGHT_HEADLINE
    )
    if insights and insight_slide_id:
        for ins in insights:
            new_id = slides.duplicate_slide(pres_id, insight_slide_id)
            _replace_in_pages(
                slides,
                pres_id,
                [new_id],
                {INSIGHT_HEADLINE: ins.headline, INSIGHT_DETAIL: ins.detail},
            )
        slides.delete_slide(pres_id, insight_slide_id)
    elif insights and not insight_slide_id:
        console.print(
            "[yellow]Инсайты найдены, но в шаблоне нет слайда с "
            "{{insight_headline}} — пропускаю.[/yellow]"
        )

    # 5. Глобальные плейсхолдеры
    mapping = {
        "{{period}}": current_period,
        "{{previous_period}}": previous_period,
        "{{report_name}}": cfg.name,
        **_kpi_placeholders(data, current_period),
        **cfg.static_placeholders,
    }
    slides.replace_placeholders(pres_id, mapping)

    console.print(f"[green]Готово[/green] → {slides.presentation_url(pres_id)}")


def _collect_insights(data, cfg, current_period, previous_period):
    insights = []
    ch = data.get("channels")
    if ch is None or ch.empty:
        return insights
    for metric in ("spend", "revenue", "leads"):
        if metric not in ch.columns:
            continue
        ch[metric] = ch[metric].apply(_to_num)
        insights += qoq_changes(
            ch,
            entity_col="channel",
            metric_col=metric,
            current_period=current_period,
            previous_period=previous_period,
            threshold_pct=cfg.insights.qoq_threshold_pct,
        )
    cur = ch[ch.get("period") == current_period].copy() if "period" in ch.columns else ch.copy()
    if cur.empty:
        return insights
    if "spend" in cur.columns:
        insights += sigma_anomalies(
            cur, entity_col="channel", metric_col="spend", sigma=cfg.insights.sigma
        )
    if "roas" in cur.columns:
        cur["roas"] = cur["roas"].apply(_to_num)
        insights += roas_below_benchmark(
            cur, entity_col="channel", benchmark=cfg.insights.roas_benchmark
        )
    return insights


def _kpi_placeholders(data, current_period) -> dict[str, str]:
    """Простейшие KPI из источника 'channels': суммы за текущий период."""
    ch = data.get("channels")
    if ch is None or ch.empty:
        return {}
    cur = ch[ch.get("period") == current_period] if "period" in ch.columns else ch
    out: dict[str, str] = {}
    for col in ("spend", "revenue", "leads"):
        if col in cur.columns:
            total = cur[col].apply(_to_num).sum()
            out[f"{{{{total_{col}}}}}"] = _fmt(total)
    return out


def _resolve_spreadsheet(src: SheetSource, folder_id, drive: DriveClient) -> str:
    if src.spreadsheet_id:
        return src.spreadsheet_id
    return drive.find_in_folder(folder_id, src.name_pattern, MIME_SHEET)["id"]


def _replace_in_pages(slides: SlidesClient, presentation_id, page_ids, mapping):
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": k, "matchCase": True},
                "replaceText": v,
                "pageObjectIds": page_ids,
            }
        }
        for k, v in mapping.items()
    ]
    slides._slides.presentations().batchUpdate(  # noqa: SLF001
        presentationId=presentation_id, body={"requests": requests}
    ).execute()


def _to_num(x):
    try:
        return float(str(x).replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except (ValueError, TypeError):
        return float("nan")


def _fmt(x: float) -> str:
    if x != x:  # NaN
        return "—"
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.0f}K"
    return f"{x:.0f}"


if __name__ == "__main__":
    app()

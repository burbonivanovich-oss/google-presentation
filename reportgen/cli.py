from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from .auth import get_credentials
from .composer import compose_report
from .config import ReportConfig, SheetSource
from .design_test import build_design_test_deck
from .drive import MIME_SHEET, MIME_SLIDES, DriveClient
from .insights import qoq_changes, roas_below_benchmark, sigma_anomalies
from .planner import build_plan
from .sheets import SheetsClient
from .slides import SlidesClient
from .template_index import build_template_index, fetch_slides

app = typer.Typer(help="Генератор квартальных Google Slides отчётов из Google Sheets.")
console = Console()


@app.command()
def auth() -> None:
    """Локальный браузерный логин в Google и сохранение token.json."""
    get_credentials()
    console.print("[green]OK[/green] токен сохранён в secrets/token.json")
    console.print(
        "Дальше: скопируйте содержимое secrets/token.json в GitHub Secret "
        "GOOGLE_OAUTH_TOKEN_JSON."
    )


@app.command(name="list-folder")
def list_folder(folder_id: str = typer.Argument(..., help="ID папки в Drive")) -> None:
    """Показать всё, что доступно в папке."""
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
            "application/vnd.google-apps.spreadsheet": "Google Sheet",
            "application/vnd.google-apps.presentation": "Google Slides",
            "application/vnd.google-apps.folder": "Folder",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx (не Google!)",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx (не Google!)",
        }.get(f["mimeType"], f["mimeType"])
        table.add_row(f["name"], kind, f["id"])
    console.print(table)


@app.command(name="list-slides")
def list_slides(presentation_id: str = typer.Argument(..., help="ID презентации")) -> None:
    """Показать слайды и их превью."""
    creds = get_credentials()
    slides = SlidesClient(creds)
    table = Table(title=f"Слайды презентации {presentation_id}")
    table.add_column("objectId", style="dim"); table.add_column("Превью")
    for oid, preview in slides.list_slides(presentation_id):
        table.add_row(oid, preview or "(пусто)")
    console.print(table)


@app.command(name="inspect-template")
def inspect_template(
    presentation_id: str = typer.Argument(..., help="ID шаблона презентации"),
    raw: bool = typer.Option(False, "--raw", help="Печать всех слайдов с layout-именами"),
) -> None:
    """Распознать какие layout-эталоны из шаблона будут использованы."""
    creds = get_credentials()
    slides = SlidesClient(creds)
    _, slides_list = fetch_slides(slides._slides, presentation_id)  # noqa: SLF001

    if raw:
        rt = Table(title=f"Все слайды шаблона ({len(slides_list)})")
        rt.add_column("#"); rt.add_column("Layout"); rt.add_column("Slide ID", style="dim")
        rt.add_column("1-й shape (превью)")
        for i, s in enumerate(slides_list, 1):
            preview = ""
            for sh in s.shapes:
                if sh.text:
                    preview = sh.text.replace("\n", " ")[:60]
                    break
            rt.add_row(str(i), s.layout_name or "(пусто)", s.object_id, preview)
        console.print(rt)
        return

    index = build_template_index(slides_list)
    table = Table(title=f"Распознанные эталоны в {presentation_id}")
    table.add_column("Роль", style="bold")
    table.add_column("Layout")
    table.add_column("Slide ID", style="dim")
    table.add_column("Slots / Найдено")
    for role, entry in index.items():
        slots_total = sum(1 for _ in entry.shape_by_role) or 0
        from .template_map import by_role
        spec = by_role(role)
        expected = len(spec.slots) if spec else 0
        table.add_row(role, entry.slide.layout_name, entry.slide.object_id,
                      f"{slots_total}/{expected}")
    console.print(table)

    from .template_map import LAYOUTS
    missing = [s.role for s in LAYOUTS if s.role not in index]
    if missing:
        console.print(f"[yellow]Не нашлось:[/yellow] {', '.join(missing)}")
        console.print("Запустите с --raw, чтобы увидеть все layout-имена в шаблоне.")


@app.command(name="design-test")
def design_test(
    folder_id: str | None = typer.Option(
        None, "--folder", help="ID папки в Drive"
    ),
) -> None:
    """Создать презентацию-стенд для проверки дизайн-фиделити Slides API."""
    creds = get_credentials()
    slides = SlidesClient(creds)
    sheets = SheetsClient(creds)
    console.print("Собираю design-fidelity тест...")
    pres_id = build_design_test_deck(
        slides._slides, sheets._svc, slides._drive,  # noqa: SLF001
        parent_folder_id=folder_id,
    )
    console.print(f"[green]Готово[/green] → {slides.presentation_url(pres_id)}")


@app.command()
def generate(
    config_path: Path = typer.Option(..., "--config", "-c", help="YAML конфиг отчёта"),
    current_period: str = typer.Option(..., "--period", help="например Q1-2026"),
    previous_period: str = typer.Option(..., "--prev", help="например Q4-2025"),
) -> None:
    """Сгенерировать отчёт автономно: planner выбирает слайды, composer
    собирает их из шаблона Контура и заполняет данными."""
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
            cols = ", ".join(data[key].columns[:10])
            console.print(f"  • {key}: {data[key].shape[0]} строк, колонки: [{cols}]")
        except Exception as e:  # noqa: BLE001
            console.print(f"  [yellow]пропуск {key}: {e}[/yellow]")

    # 2. Инсайты
    insights = _collect_insights(data, cfg, current_period, previous_period)
    console.print(f"Найдено инсайтов: {len(insights)}")
    data["_insights"] = insights

    # 3. План
    plan = build_plan(
        report_name=cfg.name,
        period=current_period,
        previous_period=previous_period,
        data=data,
    )
    console.print(f"План: {len(plan)} слайдов — " + ", ".join(s.role for s in plan))

    # 4. Композиция
    template_id = cfg.presentation_template_id or drive.find_in_folder(
        cfg.folder_id, cfg.presentation_template_name, MIME_SLIDES
    )["id"]
    console.print(f"Шаблон: {template_id}")
    _, slides_list = fetch_slides(slides._slides, template_id)  # noqa: SLF001
    index_preview = build_template_index(slides_list)
    console.print(f"Распознано ролей в шаблоне: {len(index_preview)} — "
                  + ", ".join(index_preview.keys()))
    title = f"{cfg.name} — {current_period}"
    pres_id = compose_report(
        slides_svc=slides._slides,  # noqa: SLF001
        sheets_svc=sheets._svc,  # noqa: SLF001
        drive_svc=slides._drive,  # noqa: SLF001
        template_id=template_id,
        plan=plan,
        title=title,
        parent_folder_id=cfg.folder_id,
    )
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
            ch, entity_col="channel", metric_col=metric,
            current_period=current_period, previous_period=previous_period,
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


def _resolve_spreadsheet(src: SheetSource, folder_id, drive: DriveClient) -> str:
    if src.spreadsheet_id:
        return src.spreadsheet_id
    return drive.find_in_folder(folder_id, src.name_pattern, MIME_SHEET)["id"]


def _to_num(x):
    try:
        return float(str(x).replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except (ValueError, TypeError):
        return float("nan")


if __name__ == "__main__":
    app()

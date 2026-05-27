# google-presentation

Автоматическая сборка квартальных Google Slides отчётов из данных в Google Sheets.

Текущий статус: **MVP** — CLI + один отчёт (digital marketing). Авторизация через
Google **Service Account**. Шаблон презентации копируется, в копию подставляются
значения KPI и динамически дублируются «слайды-инсайты».

## Установка

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Подключение к Google API (Service Account)

1. В Google Cloud Console включите три API в вашем проекте
   (*APIs & Services → Library*): **Google Sheets API**, **Google Slides API**,
   **Google Drive API**.
2. *IAM & Admin → Service Accounts → Create service account*:
   - Name: `report-generator` (любое).
   - Роли проекта можно не назначать — доступ выдаётся через шаринг файлов.
   - Создать.
3. У созданного аккаунта откройте вкладку **Keys → Add key → Create new key → JSON**.
   Скачается файл с client_email и приватным ключом.
4. Положите его в репозиторий локально как `secrets/service_account.json`
   (он в `.gitignore`).
5. Запомните email сервис-аккаунта (вида
   `report-generator@<project>.iam.gserviceaccount.com`) — на него нужно расшарить
   папку Drive. Подсказать его умеет CLI:

   ```bash
   python -m reportgen whoami
   ```

### Запуск из GitHub Actions

В **Settings → Secrets and variables → Actions** добавьте один секрет:

| Secret name                | Что туда положить                                  |
|----------------------------|----------------------------------------------------|
| `GCP_SERVICE_ACCOUNT_JSON` | целиком содержимое `secrets/service_account.json`  |

Дальше отчёт можно запустить вручную через вкладку **Actions → Generate
quarterly report → Run workflow**, передав `config`, `period`, `prev`.

## Папка-приёмник в Google Drive

Чтобы не править ID файлов в коде каждый квартал:

1. Создайте папку в Drive, например `Quarterly Reports / Digital Marketing`.
2. Положите в неё:
   - **Шаблон презентации** — имя должно матчиться regex из конфига, например
     `Digital Marketing — template`.
   - **Таблицы-источники** — например `channels Q4-2025`, `channels Q1-2026`.
     Код возьмёт самый свежий файл, попавший под `name_pattern`.
3. В URL папки скопируйте `<FOLDER_ID>` и пропишите его в конфиг.
4. **Расшарьте папку на email сервис-аккаунта** (роль **Editor**). Доступ
   автоматически распространится на всё содержимое — и на шаблон, и на таблицы,
   и на будущие копии презентаций.

> **Важно про владельца файлов.** Сгенерированная презентация будет числиться
> за сервис-аккаунтом (у личного Gmail передать владение SA → пользователю
> нельзя — это ограничение Google). Но т.к. файл создаётся в вашей расшаренной
> папке, вы видите его и можете редактировать. Если нужен файл «за вашим
> именем» — `File → Make a copy` в Google Slides.

## Подготовка шаблона презентации

В вашей презентации-шаблоне:

- Везде, где нужно вставить значение, пишите плейсхолдер вида `{{period}}`,
  `{{company}}`, и т.п. — он будет заменён на текст из конфига или вычисленный.
- Один слайд сделайте «заготовкой под инсайт» с двумя плейсхолдерами:
  `{{insight_headline}}` (заголовок) и `{{insight_detail}}` (подпись).
  Этот слайд будет копироваться столько раз, сколько найдено инсайтов
  в данных, а исходный — удаляться.

Узнать `objectId` нужного слайда можно через API:

```bash
python - <<'PY'
from reportgen.auth import get_credentials
from reportgen.slides import SlidesClient
sc = SlidesClient(get_credentials())
for s in sc.get_presentation("YOUR_TEMPLATE_ID")["slides"]:
    print(s["objectId"])
PY
```

## Подготовка данных в Sheets

Минимально нужный лист `channels` с заголовками:

| period   | channel  | spend | revenue | roas | leads |
|----------|----------|-------|---------|------|-------|
| Q4-2025  | Google   | 500000| 1800000 | 3.6  | 230   |
| Q1-2026  | Google   | 620000| 2100000 | 3.4  | 245   |
| …        | …        | …     | …       | …    | …     |

## Генерация отчёта

```bash
cp config/digital_marketing.example.yaml config/digital_marketing.yaml
# вписать ID шаблона, ID таблицы, id insight-слайда

python -m reportgen generate \
    --config config/digital_marketing.yaml \
    --period Q1-2026 \
    --prev   Q4-2025
```

На выходе — ссылка на новую копию презентации в вашем Google Drive.

## Какие инсайты ищутся по умолчанию

| Правило            | Когда срабатывает                                              |
|--------------------|----------------------------------------------------------------|
| `qoq_jump`         | Метрика канала Q/Q изменилась больше чем на N% (по умолч. 20%) |
| `anomaly`          | Значение выходит за ±3σ от среднего по каналам в периоде       |
| `roas_below`       | ROAS канала ниже бенчмарка из конфига (по умолч. 3.0)          |

Пороги настраиваются в секции `insights:` конфига.

## Что планируется дальше

1. Второй и третий отчёты (под ваши направления, после получения шаблонов).
2. Автозапуск по расписанию (GitHub Actions / cron).
3. Веб-UI с одной кнопкой «Сгенерировать».
4. LLM-комментарии к инсайтам через Claude API (опционально).

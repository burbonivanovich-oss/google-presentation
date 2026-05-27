# google-presentation

Автоматическая сборка квартальных Google Slides отчётов из данных в Google Sheets.

Текущий статус: **MVP** — CLI + один отчёт (digital marketing). Авторизация OAuth 2.0
от имени пользователя. Шаблон презентации копируется, в копию подставляются
значения KPI и динамически дублируются «слайды-инсайты».

## Установка

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Подключение к Google API

1. В Google Cloud Console включите **Google Sheets API**, **Google Slides API** и
   **Google Drive API** в вашем проекте.
2. В разделе *APIs & Services → Credentials* создайте **OAuth client ID** типа
   *Desktop app* и скачайте `client_secret_*.json`.
3. Положите файл сюда: `secrets/client_secret.json`.
4. Выполните однократный логин — откроется браузер Google:

   ```bash
   python -m reportgen auth
   ```

   Токен сохранится в `secrets/token.json` и будет автоматически обновляться.

Файлы в `secrets/` уже в `.gitignore` и в репозиторий не попадут.

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

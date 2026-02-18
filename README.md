# Обробка даних Google Sheets та експорт до ArcGIS

Скрипт для обробки даних з Google Sheets та експорту їх до ArcGIS Feature Layer.

## Що робить скрипт

1. **Зчитує дані** з Google Sheets за вказаним spreadsheet_id
2. **Обробляє дані**: трансформує числові значення в 0-1 (на основі максимального значення)
3. **Створює нову таблицю** Google Sheets з обробленими даними
4. **Експортує дані** до ArcGIS Feature Layer як геопросторові об'єкти

## Необхідне для запуску

### Файли налаштувань

1. **credentials.json** — файл облікових даних Google OAuth 2.0
   - Отримати можна в [Google Cloud Console](https://console.cloud.google.com/)
   - Увімкніть Google Sheets API
   - Створіть OAuth 2.0 credentials та завантажте JSON

2. **.env** — файл змінних середовища (скопіюйте з .env-dist):

```bash
cp .env-dist .env
```

Оновіть конфігурації

```
GIS_FEATURE_LAYER_ID=<ID шару ArcGIS>
GIS_API_TOKEN=<API токен ArcGIS>
```

### Встановлення залежностей

```bash
pip install -r requirements.txt
```

## Перший запуск

1. Встановіть залежності
2. Скопіюйте `.env-dist` в `.env` та заповніть значення
3. Помістіть `credentials.json` в корінь проєкту
4. Запустіть скрипт — відкриється браузер для авторизації Google
5. Після авторизації створиться `token.json` (зберігайте його для наступних запусків)

## Використання

```bash
python main.py --spreadsheet_id <ID_таблиці>
```

### Приклад

```bash
python main.py --spreadsheet_id 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
```

### Аргументи

- `--spreadsheet_id` (обов'язковий) — ID Google Sheets документа з вихідними даними

## Структура вхідних даних

Таблиця повинна мати такі колонки:

- A: Дата
- B: Регіон
- C: Місто
- D-M: Числові значення (10 колонок)
- N: Довгота (long)
- O: Широта (lat)

## Результат

### Google Sheets

- Створюється нова таблиця з назвою "New Dataset [дата_час]"
- Заморожується перший рядок (заголовки)
- Заголовки форматуються жирним шрифтом з нижньою межею

### ArcGIS

- Дані додаються до вказаного Feature Layer
- Кожен рядок стає точковим об'єктом з координатами
- Атрибути включають: date, region, city, value_1-10, long, lat

## Логування

Логи записуються в файл `app.log`

## Помилки

- Перевірте наявність `credentials.json` та правильність `token.json`
- Переконайтесь, що Google Sheets API увімкнено
- Перевірте правильність ID шару ArcGIS та валідність токена

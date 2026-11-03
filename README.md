# Wildberries — Google Sheets (Apps Script)

Набор скриптов для **Google Таблиц**: выгрузки и отчёты по API Wildberries, меню в таблице, опционально уведомления в Telegram.

Токены **не хранятся в коде** — вводятся через меню (`PropertiesService`).

## Файлы (папка `gas/`, порядок в проекте Apps Script)

| Файл | Назначение |
|------|------------|
| `00_headers_ru_*.js` | Русские заголовки колонок WB |
| `01_downloads_*.js` | Главное меню, токен API, основные загрузки |
| `02_reports_*.js` | Отчёты |
| `03_postavka*.js` | Поставки FBW |
| `04_print*.js` | Печать / этикетки |
| `05_seller_stocks_*.js` | Остатки |
| `06_fbs_orders_*.js` | Заказы FBS |
| `07_fbs_supplies_*.js` | Поставки FBS |
| `08_fbs_compose_*.js` | Сборка FBS |

Файл `00_headers` должен идти первым (как в комментариях в `01_downloads`).

## Разработка и деплой (clasp)

Локальная копия привязана к проекту Apps Script через `.clasp.json` (не коммитится).

```bash
npx @google/clasp pull   # скачать из облака
npx @google/clasp push   # залить изменения
```

Первый запуск: `npx @google/clasp login`.

## Установка в новую таблицу

1. Google Таблица → **Расширения** → **Apps Script**.
2. Создайте файлы по таблице выше и вставьте код из `gas/`, либо привяжите репозиторий через clasp (`clasp clone` / настройка `scriptId`).
3. Сохраните проект, обновите таблицу — появится меню **WB Аналитика**.
4. **Ввести API токен** — токен из личного кабинета WB (категории по подсказке в диалоге).

## Стек

- Google Apps Script (JavaScript)
- API Wildberries
- Google Sheets
- Telegram Bot API (опционально, через настройку в меню)

## Автор

Pet-project / фриланс. Без гарантии совместимости с будущими версиями API WB.

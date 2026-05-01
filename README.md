# Electrocech Apps Script

Google Apps Script project for the `Електроцех` spreadsheet.

## Sheets

- `Замовлення`
- `Документ для підпису`

## Features

- Package orders by date when column `J` is checked.
- Send one Telegram message per package to the manager.
- Clear `J` after successful send.
- Write package ID into column `L`.
- Send status updates from column `K` to the engineer.
- Rebuild the `Документ для підпису` sheet automatically.

## Required columns in `Замовлення`

- `A`: ID
- `B`: Дата
- `C`: Назва
- `D`: Фірма виробник
- `E`: Каталожний номер
- `F`: Прев'ю (Фото)
- `G`: Кількість
- `H`: Терміновість
- `I`: Зауваження
- `J`: Надіслати
- `K`: Стан замовлення
- `L`: Пакет

## Setup

1. Open the spreadsheet-bound Apps Script project.
2. Copy `Code.gs` into the project.
3. Set script properties:
   - `TELEGRAM_BOT_TOKEN`
   - `TELEGRAM_MANAGER_CHAT_ID`
   - `TELEGRAM_ENGINEER_CHAT_ID`
4. Create an installable trigger:
   - function: `onEdit`
   - source: `From spreadsheet`
   - event type: `On edit`

## GitHub

Do not commit real Telegram tokens or chat IDs.

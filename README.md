# Acoustic Invoice Bot

Telegram bot for automated invoice generation for Acoustic.ge.

## Features

- Fetches live product data from `https://acoustic.ge/data/products.json`
- Matches products by SKU or product_id
- Generates Excel invoices from a master template
- Sends the generated invoice file back via Telegram

## Setup

1. Create virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure:
   - Place the master Excel template as `ინვოისი - Marneuli.xlsx`
   - Set your Telegram bot token in `config.py`

## Usage

Run the bot:
```bash
python bot.py
```

Or run in background:
```bash
screen -S invoice-bot -d -m python bot.py
```

## Telegram Commands

- `/start` - Welcome message with usage instructions
- Send text in this format to generate an invoice:
  ```
  Client Name
  13333 - 1
  14522 - 2
  ```

## Files

- `config.py` - Configuration (token, paths)
- `generator.py` - Invoice generation logic
- `bot.py` - Telegram bot interface
- `ინვოისი - Marneuli.xlsx` - Master Excel template

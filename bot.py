"""
Telegram bot entry point. python-telegram-bot v20+ async API.
"""
from __future__ import annotations

import logging
import os
import re

from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from config import TELEGRAM_TOKEN
from pdf_generator import generate_invoice

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("acoustic-invoice-bot")

# Matches "CC546424075DE-2 - 1", "13333 - 1", "CC-5464_24 - 1" etc.
# Captures everything before the last dash as SKU, and the number after as quantity
ITEM_LINE_RE = re.compile(r"^(.*)\s*[-–—xX×]\s*(\d+)\s*$")

WELCOME = (
    "👋 გამარჯობა, Irma!\n\n"
    "ეს ბოტი ავტომატურად ქმნის Acoustic.ge-ის ინვოისებს.\n\n"
    "📝 *როგორ გამოვიყენო:*\n"
    "გამომიგზავნე შეტყობინება ამ ფორმატით:\n\n"
    "```\n"
    "მარნეულის კულტურის ცენტრი\n"
    "CC546424075DE-2 - 1\n"
    "13333 - 4\n"
    "```\n"
    "• პირველი ხაზი — კლიენტის სახელი\n"
    "• თითო ხაზი = `პროდუქტის_ID - რაოდენობა`\n"
    "  (პროდუქტის კოდში შეიძლება ტირები, ქვეები ან სივრციები)\n\n"
    "მე გამოვწევ პროდუქტის სახელს და ფასს ცოცხალი მონაცემებიდან "
    "და გამოგიგზავნი მზა .xlsx ფაილს."
)


async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(WELCOME, parse_mode="Markdown")


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(WELCOME, parse_mode="Markdown")


def _parse_message(text: str) -> tuple[str, list[dict]]:
    """
    Parse the multi-line user input into (client_info, items).
    Raises ValueError if the input is malformed.
    """
    raw_lines = [ln.strip() for ln in text.splitlines()]
    lines = [ln for ln in raw_lines if ln]
    if len(lines) < 2:
        raise ValueError(
            "❗ ფორმატი არასწორია.\n"
            "პირველი ხაზი უნდა იყოს კლიენტი, შემდეგ თითო ხაზი — `ID - QTY`."
        )

    client_info = lines[0]
    items: list[dict] = []
    bad: list[str] = []
    for ln in lines[1:]:
        m = ITEM_LINE_RE.match(ln)
        if not m:
            bad.append(ln)
            continue
        pid = m.group(1).strip()  # SKU (may contain dashes)
        qty_str = m.group(2)     # Quantity string
        try:
            qty = int(qty_str)
        except ValueError:
            bad.append(ln)
            continue
        if qty <= 0:
            bad.append(ln)
            continue
        items.append({"id": pid, "qty": qty})

    if bad:
        raise ValueError(
            "❗ ვერ გავარჩიე შემდეგი ხაზები:\n• "
            + "\n• ".join(bad)
            + "\n\nმოსალოდნელი ფორმატია: `CC546424075DE-2 - 1`"
        )
    if not items:
        raise ValueError("❗ პროდუქტის არცერთი ხაზი არ მოვიდა.")

    return client_info, items


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.message
    if not msg or not msg.text:
        return

    try:
        client_info, items = _parse_message(msg.text)
    except ValueError as e:
        await msg.reply_text(str(e), parse_mode="Markdown")
        return

    await msg.chat.send_action(ChatAction.UPLOAD_DOCUMENT)
    status = await msg.reply_text("⏳ ვქმნი ინვოისს, გთხოვ მოიცადო...")

    try:
        path = await generate_invoice(client_info, items)
    except FileNotFoundError as e:
        log.exception("template missing")
        await status.edit_text(f"❌ შაბლონის ფაილი ვერ მოიძებნა:\n{e}")
        return
    except ValueError as e:
        log.warning("invoice error: %s", e)
        await status.edit_text(f"❌ {e}")
        return
    except Exception as e:
        log.exception("unexpected error")
        await status.edit_text(f"❌ მოულოდნელი შეცდომა: {e}")
        return

    try:
        with open(path, "rb") as f:
            await msg.reply_document(
                document=f,
                filename=path.rsplit("/", 1)[-1],
                caption=f"✅ ინვოისი მზადაა: {client_info}",
            )
        await status.delete()
        os.remove(path)  # Cleanup: delete generated file after successful send
    except Exception as e:
        log.exception("send failed")
        await status.edit_text(f"❌ ფაილის გაგზავნა ვერ მოხერხდა: {e}")


def main() -> None:
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    log.info("Bot starting (long polling)...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from telegram.request import HTTPXRequest

from tracker import (
    BROKERS,
    DATA_DIR,
    INACTIVE_DAYS,
    detect_broker,
    format_telegram,
    process_file,
)

load_dotenv(ROOT / ".env")

BOT_TOKEN = (os.getenv("BOT_TOKEN") or "").strip()
ALLOWED = {
    int(x.strip())
    for x in (os.getenv("ALLOWED_USER_IDS") or "").split(",")
    if x.strip().isdigit()
}
INACTIVE = int((os.getenv("INACTIVE_DAYS") or str(INACTIVE_DAYS)).strip() or INACTIVE_DAYS)
START_TEXT = (
    "Send one Excel at a time. Keep the broker name in the file name.\n\n"
    "Examples:\n"
    "• ib_accounts_2026-08-24 PU Prime.xlsx\n"
    "• ib_accounts_2026-08-24 Vantage.xlsx\n\n"
    "Report shows only recent items:\n"
    "• Inactive ~7 days (not over 1 month)\n"
    "• No funds, but traded in the last 30 days\n"
    "• Unlinked User IDs (one line per user)"
)
UNKNOWN_BROKER = (
    "Please put the broker name in the file name, then send again.\n\n"
    "Example:\n"
    "• ... PU Prime.xlsx\n"
    "• ... Vantage.xlsx"
)


def allowed(user_id: int | None) -> bool:
    if not ALLOWED:
        return True
    return user_id in ALLOWED


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update.effective_user.id if update.effective_user else None):
        return
    await update.message.reply_text(START_TEXT)


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update.effective_user.id if update.effective_user else None):
        return
    doc = update.message.document
    name = (doc.file_name or "ib.xlsx").lower()
    if not name.endswith((".xlsx", ".xls")):
        await update.message.reply_text("Please send the IB Excel file (.xlsx).")
        return

    caption = (update.message.caption or "").strip()
    broker = detect_broker(doc.file_name or name, caption)
    if not broker:
        await update.message.reply_text(UNKNOWN_BROKER)
        return

    await update.message.reply_text(f"Checking {BROKERS[broker]} file...")
    incoming = DATA_DIR / "incoming"
    incoming.mkdir(parents=True, exist_ok=True)
    local = incoming / (doc.file_name or "ib.xlsx")
    tg_file = await doc.get_file()
    await tg_file.download_to_drive(str(local))

    try:
        result, report_path = process_file(local, broker=broker, inactive_days=INACTIVE)
    except Exception as e:
        await update.message.reply_text(f"Could not read this file: {e}")
        return

    text = format_telegram(result, doc.file_name or name)
    if len(text) > 4000:
        text = text[:3900] + "\n\n...full details in the Excel report."
    await update.message.reply_text(text)
    with report_path.open("rb") as f:
        await update.message.reply_document(document=f, filename=report_path.name)


def main() -> None:
    if not BOT_TOKEN:
        print("Missing BOT_TOKEN in .env")
        sys.exit(1)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    request = HTTPXRequest(connect_timeout=30.0, read_timeout=30.0, write_timeout=30.0)
    app = Application.builder().token(BOT_TOKEN).request(request).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    print("IB Tracker bot running...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()

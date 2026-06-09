import os
import sqlite3
from datetime import datetime, time
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

BOT_TOKEN = os.getenv("8371412181:AAFItyQlef5pa2oQj6UxyLCDN4OmStDzv0M")

TIMEZONE = ZoneInfo("Asia/Phnom_Penh")
DB_FILE = "telegram_group_messages.db"


# =========================
# DATABASE
# =========================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id TEXT,
            group_name TEXT,
            sender_name TEXT,
            sender_username TEXT,
            message_text TEXT,
            has_photo TEXT,
            message_id TEXT,
            message_link TEXT,
            date_text TEXT,
            time_text TEXT,
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS config (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    conn.commit()
    conn.close()


def set_config(key, value):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO config (key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    """, (key, value))

    conn.commit()
    conn.close()


def get_config(key):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("SELECT value FROM config WHERE key = ?", (key,))
    row = cur.fetchone()

    conn.close()
    return row[0] if row else None


def save_message(
    group_id,
    group_name,
    sender_name,
    sender_username,
    message_text,
    has_photo,
    message_id,
    message_link,
    date_text,
    time_text,
    created_at
):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO messages (
            group_id,
            group_name,
            sender_name,
            sender_username,
            message_text,
            has_photo,
            message_id,
            message_link,
            date_text,
            time_text,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        group_id,
        group_name,
        sender_name,
        sender_username,
        message_text,
        has_photo,
        message_id,
        message_link,
        date_text,
        time_text,
        created_at
    ))

    conn.commit()
    conn.close()


def get_today_messages():
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
        SELECT
            group_name,
            sender_name,
            sender_username,
            message_text,
            has_photo,
            date_text,
            time_text,
            message_id,
            group_id,
            message_link
        FROM messages
        WHERE date_text = ?
        ORDER BY created_at ASC
    """, (today,))

    rows = cur.fetchall()
    conn.close()
    return rows


# =========================
# EXCEL
# =========================

def create_excel_file(rows):
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    file_name = f"telegram_group_messages_{today}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Telegram Messages"

    headers = [
        "Name Group",
        "Name Sender",
        "Username",
        "Text",
        "Photo",
        "Date",
        "Time",
        "Message ID",
        "Group ID",
        "Message Link"
    ]

    ws.append(headers)

    header_fill = PatternFill(
        start_color="D9EAF7",
        end_color="D9EAF7",
        fill_type="solid"
    )

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        ws.append(row)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    ws.freeze_panes = "A2"
    wb.save(file_name)

    return file_name


# =========================
# HELPERS
# =========================

def build_message_link(chat, message_id):
    try:
        if chat.username:
            return f"https://t.me/{chat.username}/{message_id}"

        chat_id = str(chat.id)

        if chat_id.startswith("-100"):
            internal_id = chat_id.replace("-100", "")
            return f"https://t.me/c/{internal_id}/{message_id}"

        return ""
    except Exception:
        return ""


def split_long_text(text, max_length=3900):
    parts = []

    while len(text) > max_length:
        split_at = text.rfind("\n", 0, max_length)

        if split_at == -1:
            split_at = max_length

        parts.append(text[:split_at])
        text = text[split_at:].strip()

    if text:
        parts.append(text)

    return parts


# =========================
# COMMANDS
# =========================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type == "private":
        await update.message.reply_text(
            "Hello Sophann!\n\n"
            "This bot can alert you when someone sends a message in your Telegram group.\n\n"
            f"Your personal Chat ID is:\n{chat.id}\n\n"
            "Send /setme to make this chat receive alerts and daily Excel files."
        )
    else:
        await update.message.reply_text(
            "Bot is active in this group.\n\n"
            "Important: disable privacy mode in BotFather so I can read group messages."
        )


async def setme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type != "private":
        await update.message.reply_text("Please use /setme in private chat with the bot.")
        return

    set_config("personal_chat_id", str(chat.id))

    await update.message.reply_text(
        "Done! This chat will receive:\n\n"
        "1. Instant group message alerts\n"
        "2. Daily Excel file at 5:00 PM Cambodia time"
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    today_rows = get_today_messages()

    text = (
        "Bot Status\n\n"
        f"Personal Chat ID saved: {'Yes' if personal_chat_id else 'No'}\n"
        f"Messages saved today: {len(today_rows)}\n"
        "Daily Excel time: 5:00 PM Cambodia time"
    )

    await update.message.reply_text(text)


async def sendnow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_today_messages()

    if not rows:
        await update.message.reply_text("No messages saved today.")
        return

    file_name = create_excel_file(rows)

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="Today Telegram group messages Excel file."
        )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "Available commands:\n\n"
        "/start - Show bot info and your chat ID\n"
        "/setme - Save your personal chat to receive alerts\n"
        "/status - Check bot status\n"
        "/sendnow - Send today's Excel file now\n"
        "/help - Show help message"
    )

    await update.message.reply_text(text)


# =========================
# GROUP MESSAGE COLLECTOR
# =========================

async def collect_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message

    if not message:
        return

    chat = update.effective_chat
    user = update.effective_user

    now = datetime.now(TIMEZONE)

    group_id = str(chat.id)
    group_name = chat.title or "Unknown Group"

    sender_name = user.full_name if user else "Unknown Sender"
    sender_username = f"@{user.username}" if user and user.username else ""

    message_text = message.text or message.caption or ""
    has_photo = "Yes" if message.photo else "No"
    message_id = str(message.message_id)
    message_link = build_message_link(chat, message.message_id)

    date_text = now.strftime("%Y-%m-%d")
    time_text = now.strftime("%H:%M:%S")
    created_at = now.isoformat()

    save_message(
        group_id=group_id,
        group_name=group_name,
        sender_name=sender_name,
        sender_username=sender_username,
        message_text=message_text,
        has_photo=has_photo,
        message_id=message_id,
        message_link=message_link,
        date_text=date_text,
        time_text=time_text,
        created_at=created_at
    )

    personal_chat_id = get_config("personal_chat_id")

    if personal_chat_id:
        alert_text = (
            "New group message alert\n\n"
            f"Name group: {group_name}\n"
            f"Name sender: {sender_name}\n"
            f"Username: {sender_username if sender_username else '-'}\n"
            f"Text: {message_text if message_text else '[No text]'}\n"
            f"Photo: {has_photo}\n"
            f"Date: {date_text}\n"
            f"Time: {time_text}\n"
            f"Message ID: {message_id}\n"
            f"Message link: {message_link if message_link else '-'}"
        )

        for part in split_long_text(alert_text):
            await context.bot.send_message(
                chat_id=personal_chat_id,
                text=part
            )


# =========================
# DAILY EXCEL JOB
# =========================

async def send_daily_excel(context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")

    if not personal_chat_id:
        print("No personal chat ID saved. Please send /setme to the bot in private chat.")
        return

    rows = get_today_messages()
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")

    if not rows:
        await context.bot.send_message(
            chat_id=personal_chat_id,
            text=f"No Telegram group messages saved today: {today}"
        )
        return

    file_name = create_excel_file(rows)

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=personal_chat_id,
            document=f,
            caption=f"Telegram group messages report for {today}"
        )


# =========================
# MAIN
# =========================

def main():
    if not BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN is missing. Please set it in Railway Variables.")

    init_db()

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("setme", setme))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("sendnow", sendnow))
    app.add_handler(CommandHandler("help", help_command))

    app.add_handler(
        MessageHandler(
            filters.ChatType.GROUPS
            & (filters.TEXT | filters.PHOTO)
            & ~filters.COMMAND,
            collect_group_message
        )
    )

    app.job_queue.run_daily(
        send_daily_excel,
        time=time(hour=17, minute=0, tzinfo=TIMEZONE),
        name="daily_excel_5pm"
    )

    print("Bot is running on Railway...")
    print("Daily Excel will be sent at 5:00 PM Cambodia time.")

    app.run_polling()


if __name__ == "__main__":
    main()

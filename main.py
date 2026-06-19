import os
import sqlite3
from datetime import datetime, time, timedelta
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


BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TIMEZONE = ZoneInfo("Asia/Phnom_Penh")
DB_FILE = "telegram_group_messages.db"


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

    if row:
        return row[0]
    return None


def save_message(
    group_id,
    group_name,
    sender_name,
    sender_username,
    message_text,
    has_photo,
    message_id,
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
        date_text,
        time_text,
        created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        group_id,
        group_name,
        sender_name,
        sender_username,
        message_text,
        has_photo,
        message_id,
        date_text,
        time_text,
        created_at
    ))

    conn.commit()
    conn.close()


def get_last_7_days_messages():
    now = datetime.now(TIMEZONE)
    seven_days_ago = now - timedelta(days=7)

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
        created_at
    FROM messages
    WHERE created_at >= ?
    ORDER BY created_at ASC
    """, (seven_days_ago.isoformat(),))

    rows = cur.fetchall()
    conn.close()

    return rows


def group_messages_by_same_datetime(rows):
    """
    Group messages by:
    Name Group + Name Sender + Username + Date + Time HH:MM:SS

    If text and photos are sent in the same date and same time HH:MM:SS,
    they will become only one row in Excel.

    Photo column will show quantity:
    No
    1 photo
    2 photos
    3 photos
    """
    grouped = {}

    for row in rows:
        (
            group_name,
            sender_name,
            sender_username,
            message_text,
            has_photo,
            date_text,
            time_text,
            created_at
        ) = row

        time_hhmmss = time_text[:8]
        full_date_time = f"{date_text} {time_hhmmss}"

        key = (
            group_name,
            sender_name,
            sender_username,
            full_date_time
        )

        if key not in grouped:
            grouped[key] = {
                "group_name": group_name,
                "sender_name": sender_name,
                "sender_username": sender_username,
                "texts": [],
                "photo_count": 0,
                "date": full_date_time
            }

        if message_text:
            grouped[key]["texts"].append(message_text)

        if has_photo == "Yes":
            grouped[key]["photo_count"] += 1

    return grouped


def create_excel_file(rows):
    now = datetime.now(TIMEZONE)
    file_name = f"telegram_group_messages_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

    grouped = group_messages_by_same_datetime(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Telegram Messages"

    headers = [
        "Name Group",
        "Name Sender",
        "Username",
        "Text",
        "Photo",
        "Date"
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

    for key in sorted(grouped.keys(), key=lambda x: x[3]):
        data = grouped[key]

        combined_text = "\n".join(data["texts"]) if data["texts"] else ""

        photo_count = data["photo_count"]
        if photo_count == 0:
            photo_text = "No"
        elif photo_count == 1:
            photo_text = "1 photo"
        else:
            photo_text = f"{photo_count} photos"

        ws.append([
            data["group_name"],
            data["sender_name"],
            data["sender_username"],
            combined_text,
            photo_text,
            data["date"]
        ])

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                value_length = len(str(cell.value))
                if value_length > max_length:
                    max_length = value_length

        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    ws.freeze_panes = "A2"

    wb.save(file_name)
    return file_name


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type == "private":
        await update.message.reply_text(
            f"Hello Sophann!\n\n"
            f"Your Chat ID: {chat.id}\n\n"
            f"Send /setme to save your chat ID for weekly Friday 4 PM report.\n"
            f"Send /sendnow to get Excel immediately."
        )
    else:
        await update.message.reply_text("Bot is active in this group.")


async def setme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type != "private":
        await update.message.reply_text("Please use /setme in private chat with the bot.")
        return

    set_config("personal_chat_id", str(chat.id))

    await update.message.reply_text(
        "Done! Your chat ID is saved.\n"
        "You will receive Excel report every Friday at 4 PM Cambodia time."
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    rows = get_last_7_days_messages()

    await update.message.reply_text(
        "Bot Status\n\n"
        f"Chat ID saved: {'Yes' if personal_chat_id else 'No'}\n"
        f"Messages in last 7 days: {len(rows)}\n"
        "Auto report: Every Friday at 4 PM Cambodia time\n"
        "Excel columns: Name Group, Name Sender, Username, Text, Photo, Date"
    )


async def sendnow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_last_7_days_messages()

    if not rows:
        await update.message.reply_text("No messages in the last 7 days.")
        return

    file_name = create_excel_file(rows)

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption=f"Excel report sent now.\nMessages collected: {len(rows)}"
        )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "/start - Show bot info\n"
        "/setme - Save your private chat ID for auto report\n"
        "/status - Check bot status\n"
        "/sendnow - Send Excel report now\n"
        "/help - Show help"
    )


async def collect_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message

    if not message:
        return

    chat = update.effective_chat
    user = update.effective_user
    now = datetime.now(TIMEZONE)

    group_id = str(chat.id)
    group_name = chat.title or "Unknown Group"

    if user:
        sender_name = user.full_name
        sender_username = f"@{user.username}" if user.username else ""
    else:
        sender_name = "Unknown Sender"
        sender_username = ""

    message_text = message.text or message.caption or ""
    has_photo = "Yes" if message.photo else "No"
    message_id = str(message.message_id)

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
        date_text=date_text,
        time_text=time_text,
        created_at=created_at
    )

    print(
        f"Saved | Group: {group_name} | Sender: {sender_name} | "
        f"Photo: {has_photo} | Date: {date_text} {time_text}"
    )


async def send_weekly_excel(context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    now = datetime.now(TIMEZONE)

    if not personal_chat_id:
        print("No personal chat ID saved. Please send /setme to the bot in private chat.")
        return

    rows = get_last_7_days_messages()

    if not rows:
        await context.bot.send_message(
            chat_id=personal_chat_id,
            text=f"No messages in the last 7 days.\nChecked: {now.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        return

    file_name = create_excel_file(rows)

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=personal_chat_id,
            document=f,
            caption=(
                "Weekly Telegram Excel Report\n"
                f"Auto sent: {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Messages collected: {len(rows)}"
            )
        )

    print(f"Weekly Excel report sent at {now.strftime('%Y-%m-%d %H:%M:%S')}")


def main():
    if not BOT_TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN is not set.")
        print("Please set TELEGRAM_BOT_TOKEN environment variable.")
        return

    init_db()

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("setme", setme))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("sendnow", sendnow))
    app.add_handler(CommandHandler("help", help_command))

    app.add_handler(MessageHandler(
        filters.ChatType.GROUPS
        & (filters.TEXT | filters.PHOTO)
        & ~filters.COMMAND,
        collect_group_message
    ))

    # Auto send every Friday at 4:00 PM Cambodia time
    # Monday = 0
    # Tuesday = 1
    # Wednesday = 2
    # Thursday = 3
    # Friday = 4
    app.job_queue.run_daily(
        send_weekly_excel,
        time=time(hour=16, minute=0, second=0, tzinfo=TIMEZONE),
        days=(4,),
        name="weekly_excel_friday_4pm"
    )

    print("Bot is running...")
    print("Auto Excel report: Every Friday at 4 PM Cambodia time.")
    print("Use /sendnow to send Excel immediately.")
    print("Collecting group messages silently...")

    app.run_polling()


if __name__ == "__main__":
    main()

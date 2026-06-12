import os
import sqlite3
import html
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


# =========================
# CONFIG
# =========================

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
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

    # Add missing columns if your old database has different structure
    cur.execute("PRAGMA table_info(messages)")
    existing_columns = [col[1] for col in cur.fetchall()]

    required_columns = {
        "group_id": "TEXT",
        "group_name": "TEXT",
        "sender_name": "TEXT",
        "sender_username": "TEXT",
        "message_text": "TEXT",
        "has_photo": "TEXT",
        "message_id": "TEXT",
        "date_text": "TEXT",
        "time_text": "TEXT",
        "created_at": "TEXT"
    }

    for column_name, column_type in required_columns.items():
        if column_name not in existing_columns:
            cur.execute(f"ALTER TABLE messages ADD COLUMN {column_name} {column_type}")

    # Remove duplicate saved messages if any
    cur.execute("""
    DELETE FROM messages
    WHERE id NOT IN (
        SELECT MIN(id)
        FROM messages
        GROUP BY group_id, message_id
    )
    """)

    # Prevent duplicate message saving
    cur.execute("""
    CREATE UNIQUE INDEX IF NOT EXISTS idx_unique_group_message
    ON messages (group_id, message_id)
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
    date_text,
    time_text,
    created_at
):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    INSERT OR IGNORE INTO messages (
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


# =========================
# REPORT PERIOD
# =========================

def get_latest_thursday_3pm(now=None):
    """
    Return the latest Thursday 3:00 PM before or equal to now.

    Example:
    Friday 10 AM       -> yesterday Thursday 3 PM
    Wednesday 10 AM    -> previous Thursday 3 PM
    Thursday 2 PM      -> previous Thursday 3 PM
    Thursday 4 PM      -> today Thursday 3 PM
    """

    if now is None:
        now = datetime.now(TIMEZONE)

    # Monday = 0, Tuesday = 1, Wednesday = 2, Thursday = 3
    days_since_thursday = (now.weekday() - 3) % 7

    thursday_date = now.date() - timedelta(days=days_since_thursday)

    thursday_3pm = datetime.combine(
        thursday_date,
        time(hour=15, minute=0),
        tzinfo=TIMEZONE
    )

    if now < thursday_3pm:
        thursday_3pm -= timedelta(days=7)

    return thursday_3pm


def get_sendnow_period(now=None):
    """
    /sendnow period:
    latest Thursday 3 PM until now.
    """

    if now is None:
        now = datetime.now(TIMEZONE)

    start_datetime = get_latest_thursday_3pm(now)
    end_datetime = now

    return start_datetime, end_datetime


def get_weekly_auto_period(now=None):
    """
    Automatic Thursday 3 PM report period:
    previous Thursday 3 PM until now.
    """

    if now is None:
        now = datetime.now(TIMEZONE)

    end_datetime = now
    start_datetime = get_latest_thursday_3pm(now) - timedelta(days=7)

    return start_datetime, end_datetime


# =========================
# GET REPORT DATA
# =========================

def get_grouped_messages_between(start_datetime, end_datetime):
    """
    Get report rows between start_datetime and end_datetime.

    Group condition:
    Same group + same sender + same username + same date + same HH:MM:SS
    will show only one Excel row.

    Excel output:
    Name Group | Name Sender | Username | All Text | Photo | Photo Qty
    """

    start_iso = start_datetime.isoformat()
    end_iso = end_datetime.isoformat()

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    SELECT
        group_name AS name_group,
        sender_name AS name_sender,
        sender_username AS username,

        COALESCE(
            GROUP_CONCAT(
                CASE
                    WHEN message_text IS NOT NULL
                         AND TRIM(message_text) != ''
                    THEN message_text
                END,
                CHAR(10)
            ),
            ''
        ) AS all_text,

        CASE
            WHEN SUM(CASE WHEN has_photo = 'Yes' THEN 1 ELSE 0 END) > 0
            THEN 'Yes'
            ELSE 'No'
        END AS photo,

        COALESCE(
            SUM(CASE WHEN has_photo = 'Yes' THEN 1 ELSE 0 END),
            0
        ) AS photo_qty

    FROM messages
    WHERE created_at >= ?
      AND created_at <= ?

    GROUP BY
        group_id,
        sender_name,
        sender_username,
        date_text,
        time_text

    ORDER BY
        MIN(created_at) ASC
    """, (start_iso, end_iso))

    rows = cur.fetchall()
    conn.close()

    return rows


# =========================
# EXCEL FILE
# =========================

def create_excel_file(rows, report_type="report"):
    now = datetime.now(TIMEZONE)
    today = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%H%M%S")

    file_name = f"telegram_group_messages_{report_type}_{today}_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Telegram Messages"

    headers = [
        "Name Group",
        "Name Sender",
        "Username",
        "All Text",
        "Photo",
        "Photo Qty"
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
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(max_length + 2, 70)

    ws.freeze_panes = "A2"

    wb.save(file_name)

    return file_name


# =========================
# COMMANDS
# =========================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type == "private":
        await update.message.reply_text(
            f"Hello Sophann!\n\n"
            f"Your Chat ID: {chat.id}\n\n"
            f"Use /setme to save this chat for automatic weekly Excel report.\n"
            f"Use /sendnow to receive Excel from latest Thursday 3 PM until now.\n\n"
            f"No group message alerts will be sent."
        )
    else:
        await update.message.reply_text(
            "Bot is active in this group.\n"
            "Group messages/photos will be saved silently."
        )


async def setme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat

    if chat.type != "private":
        await update.message.reply_text(
            "Please use /setme in private chat with the bot."
        )
        return

    set_config("personal_chat_id", str(chat.id))

    await update.message.reply_text(
        "Done!\n\n"
        "This chat will receive automatic Excel report every Thursday at 3 PM Cambodia time."
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")

    now = datetime.now(TIMEZONE)
    start_datetime, end_datetime = get_sendnow_period(now)

    rows = get_grouped_messages_between(start_datetime, end_datetime)

    text = (
        "Bot Status\n\n"
        f"Chat ID saved: {'Yes' if personal_chat_id else 'No'}\n"
        f"/sendnow From: {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"/sendnow To: {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"Rows ready: {len(rows)}\n\n"
        "Automatic report: Every Thursday at 3 PM Cambodia time\n"
        "Alert message: OFF"
    )

    await update.message.reply_text(text)


async def sendnow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = datetime.now(TIMEZONE)
    start_datetime, end_datetime = get_sendnow_period(now)

    rows = get_grouped_messages_between(start_datetime, end_datetime)

    if not rows:
        await update.message.reply_text(
            "No messages found.\n\n"
            f"From: {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"To: {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        return

    file_name = create_excel_file(rows, report_type="manual")

    caption = (
        "Telegram group messages report\n\n"
        f"From: {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"To: {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}"
    )

    with open(file_name, "rb") as file:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=file,
            caption=caption
        )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "/start - Show bot information\n"
        "/setme - Save your private chat for automatic Thursday report\n"
        "/status - Check report period and row count\n"
        "/sendnow - Send Excel from latest Thursday 3 PM until now\n"
        "/help - Show help\n\n"
        "Note: Bot saves group messages silently. No alerts."
    )

    await update.message.reply_text(text)


# =========================
# COLLECT GROUP MESSAGES
# =========================

async def collect_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message

    if not message:
        return

    chat = update.effective_chat
    user = update.effective_user

    # Use real Telegram message time, converted to Cambodia time
    if message.date:
        message_datetime = message.date.astimezone(TIMEZONE)
    else:
        message_datetime = datetime.now(TIMEZONE)

    group_id = str(chat.id)
    group_name = html.unescape(chat.title or "Unknown")

    sender_name = html.unescape(user.full_name) if user else "Unknown"
    sender_username = f"@{user.username}" if user and user.username else ""

    message_text = message.text or message.caption or ""
    message_text = html.unescape(message_text)

    has_photo = "Yes" if message.photo else "No"

    message_id = str(message.message_id)

    date_text = message_datetime.strftime("%Y-%m-%d")
    time_text = message_datetime.strftime("%H:%M:%S")
    created_at = message_datetime.isoformat()

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

    # No alert.
    # Bot only saves data silently.


# =========================
# WEEKLY AUTO REPORT
# =========================

async def send_weekly_excel(context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")

    if not personal_chat_id:
        print("No personal chat ID saved. Please send /setme to the bot.")
        return

    now = datetime.now(TIMEZONE)
    start_datetime, end_datetime = get_weekly_auto_period(now)

    rows = get_grouped_messages_between(start_datetime, end_datetime)

    if not rows:
        await context.bot.send_message(
            chat_id=personal_chat_id,
            text=(
                "No messages found for weekly report.\n\n"
                f"From: {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"To: {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}"
            )
        )
        return

    file_name = create_excel_file(rows, report_type="weekly")

    caption = (
        "Weekly Telegram group messages report\n\n"
        f"From: {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"To: {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}"
    )

    with open(file_name, "rb") as file:
        await context.bot.send_document(
            chat_id=personal_chat_id,
            document=file,
            caption=caption
        )


# =========================
# MAIN
# =========================

def main():
    if not BOT_TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN not set!")
        print("Please set TELEGRAM_BOT_TOKEN environment variable.")
        return

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
        send_weekly_excel,
        time=time(hour=15, minute=0, tzinfo=TIMEZONE),
        days=(3,),
        name="weekly_excel_thursday_3pm"
    )

    print("Bot is running...")
    print("No alerts will be sent when group messages arrive.")
    print("Use /sendnow to get Excel from latest Thursday 3 PM until now.")
    print("Automatic Excel report: every Thursday at 3 PM Cambodia time.")

    app.run_polling()


if __name__ == "__main__":
    main()

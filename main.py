import os
import sqlite3
import html
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


# =========================
# REPORT DATA
# =========================

def get_unsent_grouped_messages():
    """
    Get all messages after last_report_at.
    If last_report_at does not exist, get all saved messages.
    Group by:
    - group_id
    - sender_name
    - sender_username
    - date_text
    - time_text

    This means messages/photos sent in the same HH:MM:SS become one Excel row.
    """

    last_report_at = get_config("last_report_at")

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    if last_report_at:
        cur.execute("""
        SELECT
            group_name,
            sender_name,
            sender_username,

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

            SUM(CASE WHEN has_photo = 'Yes' THEN 1 ELSE 0 END) AS photo_qty,

            MAX(created_at) AS max_created_at

        FROM messages
        WHERE created_at > ?
        GROUP BY
            group_id,
            sender_name,
            sender_username,
            date_text,
            time_text
        ORDER BY MIN(created_at) ASC
        """, (last_report_at,))
    else:
        cur.execute("""
        SELECT
            group_name,
            sender_name,
            sender_username,

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

            SUM(CASE WHEN has_photo = 'Yes' THEN 1 ELSE 0 END) AS photo_qty,

            MAX(created_at) AS max_created_at

        FROM messages
        GROUP BY
            group_id,
            sender_name,
            sender_username,
            date_text,
            time_text
        ORDER BY MIN(created_at) ASC
        """)

    rows = cur.fetchall()
    conn.close()

    return rows


def update_last_report_at(rows):
    """
    Save the latest created_at from reported rows.
    Next report will only include data after this time.
    """

    if not rows:
        return

    max_created_at = max(row[6] for row in rows if row[6])
    set_config("last_report_at", max_created_at)


# =========================
# EXCEL
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
        # row has 7 fields.
        # The last field is max_created_at.
        # Excel only needs first 6 fields.
        ws.append(row[:6])

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
            f"Hello Sophann!\n\n"
            f"Your Chat ID: {chat.id}\n\n"
            f"Send /setme to receive alerts and weekly Excel reports."
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
        "Done!\n\n"
        "You will receive alerts and Excel reports every Thursday at 3 PM Cambodia time."
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    last_report_at = get_config("last_report_at")
    rows = get_unsent_grouped_messages()

    text = (
        "Bot Status\n\n"
        f"Chat ID saved: {'Yes' if personal_chat_id else 'No'}\n"
        f"Unsent grouped rows: {len(rows)}\n"
        f"Last report at: {last_report_at or 'Never'}\n"
        "Weekly report: Every Thursday at 3 PM Cambodia time"
    )

    await update.message.reply_text(text)


async def sendnow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_unsent_grouped_messages()

    if not rows:
        await update.message.reply_text("No new messages since last report.")
        return

    file_name = create_excel_file(rows, "manual")

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="Messages since last report"
        )

    update_last_report_at(rows)

    await update.message.reply_text("Done. Last report time has been updated.")


async def resetlast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Optional command.
    Use this if you want next report to include all old data again.
    """

    set_config("last_report_at", "")

    await update.message.reply_text(
        "Done. Last report time has been reset.\n"
        "Next /sendnow or weekly report will include all saved data."
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "/start - Show bot info\n"
        "/setme - Save your private chat ID for alerts and weekly reports\n"
        "/status - Check bot status\n"
        "/sendnow - Send Excel now with all new data since last report\n"
        "/resetlast - Reset last report time\n"
        "/help - Show help"
    )

    await update.message.reply_text(text)


# =========================
# COLLECT GROUP MESSAGE
# =========================

async def collect_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message

    if not message:
        return

    chat = update.effective_chat
    user = update.effective_user
    now = datetime.now(TIMEZONE)

    group_id = str(chat.id)
    group_name = html.unescape(chat.title or "Unknown")

    sender_name = html.unescape(user.full_name) if user else "Unknown"
    sender_username = f"@{user.username}" if user and user.username else ""

    message_text = message.text or message.caption or ""
    message_text = html.unescape(message_text)

    has_photo = "Yes" if message.photo else "No"

    message_id = str(message.message_id)
    message_link = build_message_link(chat, message.message_id)

    date_text = now.strftime("%Y-%m-%d")
    time_text = now.strftime("%H:%M:%S")
    created_at = now.isoformat()

    save_message(
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

    personal_chat_id = get_config("personal_chat_id")

    if personal_chat_id:
        alert_text = (
            "New message\n\n"
            f"Group: {group_name}\n"
            f"Sender: {sender_name}\n"
            f"Username: {sender_username or '-'}\n"
            f"Text: {message_text or '[No text]'}\n"
            f"Photo: {has_photo}\n"
            f"Date: {date_text}\n"
            f"Time: {time_text}"
        )

        for part in split_long_text(alert_text):
            await context.bot.send_message(
                chat_id=personal_chat_id,
                text=part
            )


# =========================
# WEEKLY REPORT
# =========================

async def send_weekly_excel(context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")

    if not personal_chat_id:
        print("No personal chat ID saved. Please send /setme to the bot.")
        return

    rows = get_unsent_grouped_messages()
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")

    if not rows:
        await context.bot.send_message(
            chat_id=personal_chat_id,
            text=f"No new messages since last report: {today}"
        )
        return

    file_name = create_excel_file(rows, "weekly")

    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=personal_chat_id,
            document=f,
            caption=f"Weekly report: messages since last report until {today}"
        )

    update_last_report_at(rows)


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
    app.add_handler(CommandHandler("resetlast", resetlast))
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
    print("Weekly Excel report will be sent every Thursday at 3 PM Cambodia time.")

    app.run_polling()


if __name__ == "__main__":
    main()

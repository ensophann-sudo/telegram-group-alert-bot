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

def save_message(group_id, group_name, sender_name, sender_username, message_text, has_photo, message_id, message_link, date_text, time_text, created_at):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO messages (group_id, group_name, sender_name, sender_username, message_text, has_photo, message_id, message_link, date_text, time_text, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (group_id, group_name, sender_name, sender_username, message_text, has_photo, message_id, message_link, date_text, time_text, created_at))
    conn.commit()
    conn.close()

def get_last_7_days_messages():
    """Get messages from last 7 days (since last Thursday 3 PM)"""
    now = datetime.now(TIMEZONE)
    seven_days_ago = now - timedelta(days=7)
    
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
    SELECT group_name, sender_name, sender_username, message_text, has_photo, date_text, time_text, message_id, group_id, message_link
    FROM messages 
    WHERE created_at >= ? 
    ORDER BY created_at ASC
    """, (seven_days_ago.isoformat(),))
    rows = cur.fetchall()
    conn.close()
    return rows

def create_excel_file(rows):
    now = datetime.now(TIMEZONE)
    file_name = f"telegram_group_messages_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Telegram Messages"
    headers = ["Name Group", "Name Sender", "Username", "Text", "Photo", "Date", "Time", "Message ID", "Group ID", "Message Link"]
    ws.append(headers)
    header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
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

def build_message_link(chat, message_id):
    try:
        if chat.username:
            return f"https://t.me/{chat.username}/{message_id}"
        chat_id = str(chat.id)
        if chat_id.startswith("-100"):
            internal_id = chat_id.replace("-100", "")
            return f"https://t.me/c/{internal_id}/{message_id}"
        return ""
    except:
        return ""

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    if chat.type == "private":
        await update.message.reply_text(f"Hello Sophann!\n\nYour Chat ID: {chat.id}\n\nSend /setme to receive weekly reports.")
    else:
        await update.message.reply_text("Bot is active in this group.")

async def setme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    if chat.type != "private":
        await update.message.reply_text("Please use /setme in private chat.")
        return
    set_config("personal_chat_id", str(chat.id))
    await update.message.reply_text("Done! You will receive weekly Excel reports every Thursday at 3 PM Cambodia time.")

async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    rows = get_last_7_days_messages()
    text = f"Bot Status\n\nChat ID saved: {'Yes' if personal_chat_id else 'No'}\nMessages (last 7 days): {len(rows)}\nWeekly report: Every Thursday at 3 PM"
    await update.message.reply_text(text)

async def sendnow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_last_7_days_messages()
    if not rows:
        await update.message.reply_text("No messages in the last 7 days.")
        return
    file_name = create_excel_file(rows)
    with open(file_name, "rb") as f:
        await context.bot.send_document(chat_id=update.effective_chat.id, document=f, caption=f"Messages from last 7 days ({len(rows)} messages)")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "/start - Show info\n/setme - Save chat\n/status - Check status\n/sendnow - Send Excel now (last 7 days)\n/help - Help"
    await update.message.reply_text(text)

async def collect_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Collect messages from group without sending alerts"""
    message = update.message
    if not message:
        return
    chat = update.effective_chat
    user = update.effective_user
    now = datetime.now(TIMEZONE)
    group_id = str(chat.id)
    group_name = chat.title or "Unknown"
    sender_name = user.full_name if user else "Unknown"
    sender_username = f"@{user.username}" if user and user.username else ""
    message_text = message.text or message.caption or ""
    has_photo = "Yes" if message.photo else "No"
    message_id = str(message.message_id)
    message_link = build_message_link(chat, message.message_id)
    date_text = now.strftime("%Y-%m-%d")
    time_text = now.strftime("%H:%M:%S")
    created_at = now.isoformat()
    
    # Save to database (no alert sent)
    save_message(group_id, group_name, sender_name, sender_username, message_text, has_photo, message_id, message_link, date_text, time_text, created_at)
    print(f"Message saved from {sender_name} in {group_name}")

async def send_weekly_excel(context: ContextTypes.DEFAULT_TYPE):
    personal_chat_id = get_config("personal_chat_id")
    if not personal_chat_id:
        print("No personal chat ID saved. Please send /setme to the bot.")
        return
    
    rows = get_last_7_days_messages()
    now = datetime.now(TIMEZONE)
    
    if not rows:
        await context.bot.send_message(chat_id=personal_chat_id, text=f"No messages in the last 7 days ({now.strftime('%Y-%m-%d %H:%M:%S')})")
        return
    
    file_name = create_excel_file(rows)
    with open(file_name, "rb") as f:
        await context.bot.send_document(
            chat_id=personal_chat_id, 
            document=f, 
            caption=f"📊 Weekly report: {now.strftime('%Y-%m-%d %H:%M:%S')}\n{len(rows)} messages from last 7 days"
        )
    print(f"Weekly Excel report sent with {len(rows)} messages")

def main():
    if not BOT_TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN not set!")
        return
    
    init_db()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    
    # Add handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("setme", setme))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("sendnow", sendnow))
    app.add_handler(CommandHandler("help", help_command))
    
    # Message handler for group messages (collects silently, no alerts)
    app.add_handler(MessageHandler(
        filters.ChatType.GROUPS & (filters.TEXT | filters.PHOTO) & ~filters.COMMAND, 
        collect_group_message
    ))
    
    # Weekly job - Thursday 3 PM
    app.job_queue.run_daily(
        send_weekly_excel, 
        time=time(hour=15, minute=0, tzinfo=TIMEZONE), 
        days=(3,),  # 3 = Thursday
        name="weekly_excel_thursday_3pm"
    )
    
    print("Bot is running...")
    print("Weekly Excel report will be sent every Thursday at 3 PM Cambodia time")
    print("Messages are collected silently (no alerts)")
    app.run_polling()

if __name__ == "__main__":
    main()

import telebot
import os
import fitz
import pandas as pd
import re

TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise ValueError("BOT_TOKEN environment variable is not set.")

bot = telebot.TeleBot(TOKEN, parse_mode='HTML')

REFERENCE_EXCEL_PATH = 'ref-refinary.xlsx'
if not os.path.exists(REFERENCE_EXCEL_PATH):
    raise FileNotFoundError(f"Reference Excel file not found: {REFERENCE_EXCEL_PATH}")

def extract_mid_rate(text, code):
    pattern = rf'{re.escape(code)}\s+(?:[\d.]+[\u2013\-][\d.]+\s+)?([\d.]+)'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return 'Not Found'

def extract_date(text):
    match = re.search(r'December\s+(\d{1,2}),\s+(\d{4})', text)
    if match:
        day = int(match.group(1))
        year = int(match.group(2))
        return year, 12, day
    return None

def to_shamsi_simple(gy, gm, gd):
    sh_year = gy - 621
    if gd < 22:
        sh_month = 9
        sh_day = gd + 9
    else:
        sh_month = 10
        sh_day = gd - 21
    return f"{sh_year:04d}/{sh_month:02d}/{sh_day:02d}"

def extract_and_format_rates(pdf_bytes):
    df = pd.read_excel(REFERENCE_EXCEL_PATH)
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        full_text = "\n".join([page.get_text() for page in doc])

    date_parts = extract_date(full_text)
    if date_parts:
        shamsi_date = to_shamsi_simple(*date_parts)
        header = f"<b>تاریخ:</b> {shamsi_date}\n"
    else:
        header = "<b>تاریخ:</b> نامشخص\n"

    output_lines = [header]

    for _, row in df.iterrows():
        product = str(row.iloc[0]).strip()
        code = str(row.iloc[2]).strip()
        unit = str(row.iloc[1]).strip()
        rate = extract_mid_rate(full_text, code)
        output_lines.append(f"<b>{product}</b>: {rate} {unit}")

    return "\n".join(output_lines)

@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        response_text = extract_and_format_rates(downloaded_file)
        bot.send_message(message.chat.id, response_text)
    except Exception as e:
        bot.send_message(message.chat.id, f"خطا در پردازش فایل: {str(e)}")

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    welcome_text = (
        "سلام! به بات محاسبه نرخ ارز خوش آمدید ✨\n"
        "این بات می‌تونه در گروه‌ها هم استفاده بشه.\n"
        "\n<b>فرمان‌ها:</b>\n"
        "/start - معرفی بات\n"
        "/help - راهنمای استفاده\n"
        "/menu - منو دسترسی سریع"
    )
    bot.reply_to(message, welcome_text)

@bot.message_handler(commands=['menu'])
def send_menu(message):
    menu_text = (
        "<b>منو دسترسی سریع:</b>\n"
        "• /start - معرفی بات\n"
        "• /help - راهنما\n"
        "• ارسال فایل PDF - استخراج نرخ محصولات"
    )
    bot.send_message(message.chat.id, menu_text)

@bot.my_chat_member_handler()
def on_added_to_group(message):
    if message.new_chat_member.user.id == bot.get_me().id:
        bot.send_message(message.chat.id, "سلام بچه‌ها! من الان اضافه شدم به گروه‌تون تا با ارسال فایل PDF نرخ محصولات رو استخراج کنم ✨")

bot.infinity_polling()

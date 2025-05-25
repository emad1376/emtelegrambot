
import telebot
import os
import fitz
import pandas as pd
import re
from io import BytesIO

TOKEN = os.getenv("BOT_TOKEN")
bot = telebot.TeleBot(TOKEN)

REFERENCE_EXCEL_PATH = 'ref-refinary.xlsx'

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
    return f"{sh_year:04d}{sh_month:02d}{sh_day:02d}"

def extract_all_mid_rates_from_file(pdf_bytes, original_filename):
    df = pd.read_excel(REFERENCE_EXCEL_PATH)
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        full_text = "\n".join([page.get_text() for page in doc])

    date_parts = extract_date(full_text)
    base_name = os.path.splitext(original_filename)[0]
    if date_parts:
        shamsi_date = to_shamsi_simple(*date_parts)
        output_filename = f"{base_name}_{shamsi_date}.xlsx"
    else:
        output_filename = f"{base_name}_unknown.xlsx"

    df['Mid Rate'] = df.iloc[:, 2].apply(lambda code: extract_mid_rate(full_text, str(code).strip()))

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output, output_filename

@bot.message_handler(content_types=['document'])
def handle_document(message):
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    output_io, output_name = extract_all_mid_rates_from_file(downloaded_file, message.document.file_name)
    bot.send_document(message.chat.id, output_io, visible_file_name=output_name)

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "سلام! فایل PDF مجله‌ رو بفرست تا نرخ‌ها استخراج بشن و فایل اکسل برگرده ✨")

bot.infinity_polling()

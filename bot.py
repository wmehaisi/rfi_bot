import os
from flask import Flask, request
from telegram import Bot, Update
from telegram.ext import Dispatcher, MessageHandler, Filters, CommandHandler
import pdfplumber
import openpyxl
import tempfile

BOT_TOKEN = os.environ.get("BOT_TOKEN")
bot = Bot(token=BOT_TOKEN)

app = Flask(__name__)

# Permanent directory on Render
UPLOAD_DIR = "/data"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

EXCEL_FILE = os.path.join(UPLOAD_DIR, "excel.xlsx")


def start(update, context):
    update.message.reply_text(
        "Welcome! 👋\n\n"
        "1️⃣ Upload your Excel file first.\n"
        "2️⃣ Then upload your PDF RFI files.\n"
        "I will update Excel automatically."
    )


# ---------------------------
# FIXED EXCEL FILE HANDLER
# ---------------------------
def handle_excel(update, context):
    file = update.message.document
    file_name = file.file_name.lower()

    # Accept ALL Excel MIME types coming from Telegram
    allowed_excel = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/octet-stream",
        "application/vnd.ms-excel",
    ]

    if file.mime_type not in allowed_excel or not file_name.endswith(".xlsx"):
        update.message.reply_text("❌ Please upload a valid Excel (.xlsx) file.")
        return

    # Save Excel permanently on render
    file.get_file().download(EXCEL_FILE)
    update.message.reply_text("✔ Excel uploaded successfully.\nNow send me PDF files!")


# ---------------------------
# PDF PROCESSING
# ---------------------------
def extract_rfi_info(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        rfi_number = None
        description = None

        for line in text.split("\n"):
            line = line.strip()

            if "Inspection Request" in line or "RFI" in line:
                description = line

            if "Rev" in line:
                digits = "".join(filter(str.isdigit, line))
                if digits:
                    rfi_number = digits

        return rfi_number, description

    except Exception as e:
        print("PDF ERROR:", e)
        return None, None


def handle_pdf(update, context):
    if not os.path.exists(EXCEL_FILE):
        update.message.reply_text("❌ Please upload Excel first!")
        return

    file = update.message.document

    if not file.file_name.lower().endswith(".pdf"):
        update.message.reply_text("❌ Please upload a PDF file.")
        return

    update.message.reply_text("📄 Extracting information from PDF...")

    # Save PDF temporarily
    pdf_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    file.get_file().download(pdf_temp.name)

    rfi_number, description = extract_rfi_info(pdf_temp.name)

    if not rfi_number or not description:
        update.message.reply_text("❌ Could not extract RFI info from this PDF.")
        return

    # Load Excel and update it
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    ws.append([rfi_number, description])
    wb.save(EXCEL_FILE)

    update.message.reply_text("✔ PDF processed and Excel updated!")


def webhook_route():
    if request.method == "POST":
        update = Update.de_json(request.json, bot)
        dispatcher.process_update(update)
    return "OK", 200


# Dispatcher
dispatcher = Dispatcher(bot, None, use_context=True)
dispatcher.add_handler(CommandHandler("start", start))

# Accept Excel using flexible MIME types
dispatcher.add_handler(MessageHandler(Filters.document, handle_excel))
dispatcher.add_handler(MessageHandler(Filters.document.mime_type("application/pdf"), handle_pdf))


@app.route("/", methods=["GET"])
def home():
    return "Bot is running!"


@app.route(f"/webhook/{BOT_TOKEN}", methods=["POST"])
def webhook_handler():
    return webhook_route()


# Debug version route
@app.route("/debug", methods=["GET"])
def debug():
    return "VERSION: FIXED-1.0", 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

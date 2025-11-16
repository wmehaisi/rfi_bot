import os
import re
import tempfile

from flask import Flask, request
from telegram import Bot, Update
from telegram.ext import Dispatcher, CommandHandler, MessageHandler, Filters

import pdfplumber
import openpyxl

# --- Telegram bot setup ---

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Environment variable BOT_TOKEN is missing")

bot = Bot(BOT_TOKEN)
app = Flask(__name__)


# --- File storage paths ---

# Render keeps /data between restarts on free tier
UPLOAD_DIR = "/data"
os.makedirs(UPLOAD_DIR, exist_ok=True)

EXCEL_FILE = os.path.join(UPLOAD_DIR, "rfi_log.xlsx")


def log(*args):
    """Print to logs immediately (visible in Render Logs)."""
    print(*args, flush=True)


# --- Handlers ---

def start(update, context):
    update.message.reply_text(
        "Welcome! 👋\n\n"
        "1️⃣ Upload your *Excel* file first (.xlsx).\n"
        "2️⃣ Then upload your *PDF RFI* files.\n"
        "I will extract info and update the Excel file automatically.",
        parse_mode="Markdown",
    )


def save_excel(doc, update):
    """Download and save the Excel file permanently."""
    doc.get_file().download(EXCEL_FILE)
    log(f"Saved Excel to {EXCEL_FILE}")
    update.message.reply_text("✔ Excel uploaded successfully.\nNow send me PDF files!")


def extract_rfi_info_from_pdf(pdf_path, filename):
    """
    Extract RFI number and description from the PDF.

    - RFI number: first group of digits in the filename
      e.g. WIR-CIV-OHTL-855 Rev.00.pdf -> 855
    - Description: first line containing 'Inspection Request' or 'RFI'
    """
    rfi_number = None
    m = re.search(r"(\d+)", filename)
    if m:
        rfi_number = m.group(1)

    description = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texts = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texts.append(t)
            full_text = "\n".join(texts)

        for line in full_text.splitlines():
            if "Inspection Request" in line or "RFI" in line:
                description = line.strip()
                break
    except Exception as e:
        log("Error reading PDF:", e)

    return rfi_number, description


def handle_document(update, context):
    """Handle ANY document (Excel or PDF). Decide by file extension."""
    if not update.message or not update.message.document:
        return

    doc = update.message.document
    fname = (doc.file_name or "").lower()
    log("Received document:", fname)

    # 1) Excel files
    if fname.endswith((".xlsx", ".xlsm", ".xls")):
        save_excel(doc, update)
        return

    # 2) PDF files
    if fname.endswith(".pdf"):
        if not os.path.exists(EXCEL_FILE):
            update.message.reply_text(
                "❌ Please upload the Excel file first, then send PDFs."
            )
            return

        update.message.reply_text("📄 Extracting information from PDF...")

        # Save PDF temporarily
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        doc.get_file().download(tmp.name)
        log("Saved temp PDF:", tmp.name)

        rfi_number, description = extract_rfi_info_from_pdf(tmp.name, fname)

        if not (rfi_number and description):
            update.message.reply_text(
                "❌ I couldn't extract RFI info from this PDF."
            )
            return

        # Update Excel
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            sheet = wb.active
            sheet.append([rfi_number, description])
            wb.save(EXCEL_FILE)
            log("Excel updated with:", rfi_number, description)
        except Exception as e:
            log("Error updating Excel:", e)
            update.message.reply_text(
                "⚠ An error happened while updating the Excel file."
            )
            return

        update.message.reply_text("✔ PDF processed and Excel updated!")
        return

    # 3) Any other document type
    update.message.reply_text(
        "❌ Please send either an Excel (.xlsx) file or a PDF file."
    )


# --- Dispatcher setup (python-telegram-bot) ---

dispatcher = Dispatcher(bot, None, workers=0, use_context=True)
dispatcher.add_handler(CommandHandler("start", start))
dispatcher.add_handler(MessageHandler(Filters.document, handle_document))


# --- Flask routes ---

@app.route("/", methods=["GET"])
def index():
    return "Bot is running!", 200


@app.route(f"/webhook/{BOT_TOKEN}", methods=["POST"])
def webhook():
    """Telegram will POST updates here."""
    if request.method == "POST":
        update = Update.de_json(request.get_json(force=True), bot)
        dispatcher.process_update(update)
    return "OK", 200


# --- Local run (not used on Render, but useful for debugging) ---

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

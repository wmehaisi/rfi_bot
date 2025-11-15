import os
import logging
from flask import Flask, request
from telegram import Update
from telegram.ext import Dispatcher, MessageHandler, Filters, CallbackContext, CommandHandler
import pdfplumber
import openpyxl
from werkzeug.utils import secure_filename

# Logging
logging.basicConfig(level=logging.INFO)

TOKEN = os.getenv("BOT_TOKEN")

app = Flask(__name__)

# ================= TELEGRAM DISPATCHER =====================
from telegram import Bot
bot = Bot(token=TOKEN)
dispatcher = Dispatcher(bot, None, workers=0)

# ============================================================
# =================== BOT COMMANDS ===========================
# ============================================================

def start(update: Update, context: CallbackContext):
    update.message.reply_text(
        "Welcome! 👋\n\n"
        "1️⃣ Upload your Excel file first.\n"
        "2️⃣ Then upload any number of PDF RFI files.\n"
        "I will extract info and update Excel automatically."
    )

dispatcher.add_handler(CommandHandler("start", start))


# ============================================================
# ============== HANDLE EXCEL UPLOAD =========================
# ============================================================

EXCEL_PATH = "main.xlsx"

def handle_excel(update: Update, context: CallbackContext):
    file = update.message.document.get_file()
    filename = secure_filename(update.message.document.file_name)

    if not filename.endswith(".xlsx"):
        update.message.reply_text("❌ Please upload a valid Excel (.xlsx) file")
        return

    file.download(EXCEL_PATH)
    update.message.reply_text("✔ Excel uploaded successfully.\nNow send me PDF files!")

dispatcher.add_handler(MessageHandler(Filters.document.category("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_excel))


# ============================================================
# ============== HANDLE PDF UPLOAD ===========================
# ============================================================

def extract_from_pdf(pdf_path):
    """Extract RFI number, description, date, location, drawing number"""
    data = {
        "rfi_no": "",
        "description": "",
        "date": "",
        "location": "",
        "drawing": "",
    }

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # RFI number from filename
    base = os.path.basename(pdf_path)
    parts = base.split("-")
    data["rfi_no"] = parts[-1].replace("Rev.00.pdf", "").strip()

    # crude extraction (you can customize patterns as needed)
    lines = text.split("\n")
    for line in lines:
        if "Inspection Request" in line:
            data["description"] = line.strip()

        if "Tower" in line:
            data["location"] = line.strip()

        if "/" in line and "-" in line:
            data["date"] = line.strip()

        if "DWG" in line or "Drawing" in line:
            data["drawing"] = line.strip()

    return data


def handle_pdf(update: Update, context: CallbackContext):
    if not os.path.exists(EXCEL_PATH):
        update.message.reply_text("❌ Please upload Excel first!")
        return

    file = update.message.document.get_file()
    filename = secure_filename(update.message.document.file_name)
    pdf_path = f"input/{filename}"

    os.makedirs("input", exist_ok=True)
    file.download(pdf_path)

    update.message.reply_text("📄 Extracting information from PDF...")

    info = extract_from_pdf(pdf_path)

    # Update Excel
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    next_row = ws.max_row + 1
    ws[f"A{next_row}"] = info["rfi_no"]
    ws[f"B{next_row}"] = info["description"]
    ws[f"C{next_row}"] = info["date"]
    ws[f"D{next_row}"] = info["location"]
    ws[f"E{next_row}"] = info["drawing"]

    # Sort Excel by RFI number
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    body = rows[1:]

    body_sorted = sorted(body, key=lambda x: int(x[0]))

    ws.delete_rows(1, ws.max_row)
    ws.append(header)
    for r in body_sorted:
        ws.append(r)

    wb.save(EXCEL_PATH)

    update.message.reply_text(
        f"✔ RFI {info['rfi_no']} added!\nSend more PDFs or type /start."
    )

dispatcher.add_handler(MessageHandler(Filters.document.pdf, handle_pdf))


# ============================================================
# ================== WEBHOOK HANDLER =========================
# ============================================================

@app.route(f"/webhook/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), bot)
    dispatcher.process_update(update)
    return "OK", 200


@app.route("/", methods=["GET"])
def home():
    return "Bot running successfully!", 200


# ============================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

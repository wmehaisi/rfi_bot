import os
import re
import tempfile
import pdfplumber
import openpyxl

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
)
from telegram.ext import (
    Updater,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    Filters,
    CallbackContext,
)

# =========================
# CONFIG
# =========================
BOT_TOKEN = os.getenv("BOT_TOKEN")

START_ROW = 866
TEMPLATE_ROW_INDEX = 865
FIXED_PROJECT_NO = "4400021143"
FIXED_CLASSIFICATION = "OHTL"
FIXED_DISCIPLINE = "Civil"

# user data in memory
USERS = {}  # user_id -> dict(tmpdir, excel_path, pdfs, preview, generated)


def ensure_user(user_id: int):
    if user_id not in USERS:
        tmp = tempfile.mkdtemp(prefix=f"user_{user_id}_")
        USERS[user_id] = {
            "tmpdir": tmp,
            "excel_path": None,
            "pdfs": [],
            "preview": [],
            "generated": None,
        }
    return USERS[user_id]


# =========================
# PDF HELPERS
# =========================
def extract_text_from_pdf(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            text += p.extract_text() or ""
    return text


def parse_pdf(path: str):
    base = os.path.basename(path)
    m = re.search(r"(\d{1,6})", base)
    rfi_num = m.group(1) if m else ""

    txt = " ".join(extract_text_from_pdf(path).split())

    # drawing number like CA-1581064
    dm = re.search(r"([A-Z]{1,6}-\d{2,7})", txt)
    drawing = dm.group(1) if dm else ""

    # description
    desc = ""
    d1 = re.search(r"(Inspection Request for[^\n]{5,120})", txt)
    if d1:
        desc = d1.group(1).strip()

    # date like 02-Nov-25 or 02 Nov 2025
    date = ""
    dt = re.search(
        r"(\d{1,2}[-/ ](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-/ ]\d{2,4})",
        txt,
        re.I,
    )
    if dt:
        date = dt.group(1)

    return {
        "rfi": rfi_num,
        "drawing": drawing,
        "description": desc,
        "date": date,
    }


# =========================
# EXCEL HELPER
# =========================
def update_excel(template_path: str, out_path: str, rows):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # insert new rows to keep formatting below
    ws.insert_rows(START_ROW, len(rows))

    for i, r in enumerate(rows):
        row = START_ROW + i

        def sc(col_letter, value):
            ws["%s%d" % (col_letter, row)].value = value

        sc("C", "IRFI-C-%s" % r["rfi"])
        sc("D", FIXED_PROJECT_NO)
        sc("E", FIXED_CLASSIFICATION)
        sc("F", FIXED_DISCIPLINE)
        sc("G", "RFI-C-%s" % r["rfi"])
        sc("H", r["description"])
        sc("I", r["drawing"])
        sc("J", r["date"])
        sc("K", r["date"])

    wb.save(out_path)
    return out_path


# =========================
# TELEGRAM HANDLERS (SYNC)
# =========================
def start(update: Update, context: CallbackContext):
    ensure_user(update.effective_user.id)
    update.message.reply_text(
        "👋 Welcome!\n"
        "1️⃣ Send the Excel file (.xlsx)\n"
        "2️⃣ Then send one or more PDF RFI files.\n"
        "When both are uploaded, I will show buttons to preview and generate."
    )


def handle_docs(update: Update, context: CallbackContext):
    user_id = update.effective_user.id
    u = ensure_user(user_id)

    doc = update.message.document
    if not doc:
        return

    file = context.bot.get_file(doc.file_id)
    dest = os.path.join(u["tmpdir"], doc.file_name)
    file.download(dest)

    if dest.lower().endswith(".xlsx"):
        u["excel_path"] = dest
        update.message.reply_text("📘 Excel template uploaded.")
    elif dest.lower().endswith(".pdf"):
        u["pdfs"].append(dest)
        update.message.reply_text("📄 PDF uploaded.")
    else:
        update.message.reply_text("❌ Only .xlsx and .pdf files are supported.")
        return

    if u["excel_path"] and u["pdfs"]:
        keyboard = [
            [InlineKeyboardButton("Preview", callback_data="preview")],
            [InlineKeyboardButton("Generate Excel", callback_data="generate")],
            [InlineKeyboardButton("Download Excel", callback_data="download")],
        ]
        update.message.reply_text(
            "✅ Files ready. Choose an action:",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )


def button_callback(update: Update, context: CallbackContext):
    query = update.callback_query
    user_id = query.from_user.id
    u = ensure_user(user_id)

    data = query.data

    if data == "preview":
        # parse PDFs
        u["preview"] = [parse_pdf(p) for p in u["pdfs"]]
        lines = []
        for i, r in enumerate(u["preview"]):
            line = "Row %d: RFI=%s | %s" % (
                START_ROW + i,
                r["rfi"],
                (r["description"] or "")[:60],
            )
            lines.append(line)
        text = "🔍 Preview of rows to be written:\n\n" + "\n".join(lines)
        query.edit_message_text(text)

    elif data == "generate":
        if not u.get("preview"):
            u["preview"] = [parse_pdf(p) for p in u["pdfs"]]
        out_path = os.path.join(u["tmpdir"], "updated.xlsx")
        update_excel(u["excel_path"], out_path, u["preview"])
        u["generated"] = out_path
        query.edit_message_text("✅ Excel generated. Use 'Download Excel' to get it.")

    elif data == "download":
        if not u.get("generated"):
            query.edit_message_text("❌ Please generate the Excel file first.")
            return
        with open(u["generated"], "rb") as f:
            context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(f, filename="updated.xlsx"),
            )
        query.edit_message_text("📤 File sent.")


def main():
    token = BOT_TOKEN
    if not token:
        raise RuntimeError("BOT_TOKEN environment variable is not set.")

    updater = Updater(token=token, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.document, handle_docs))
    dp.add_handler(CallbackQueryHandler(button_callback))

    # long polling
    updater.start_polling()
    updater.idle()


if __name__ == "__main__":
    main()

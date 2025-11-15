import os
import re
import pdfplumber
import tempfile
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
)

# BOT TOKEN WILL COME FROM RENDER ENV VARIABLE
BOT_TOKEN = os.getenv("BOT_TOKEN")

START_ROW = 866
TEMPLATE_ROW_INDEX = 865
FIXED_PROJECT_NO = "4400021143"
FIXED_CLASSIFICATION = "OHTL"
FIXED_DISCIPLINE = "Civil"

# USER TEMP STORAGE
USERS = {}

def ensure_user(uid):
    if uid not in USERS:
        tmp = tempfile.mkdtemp(prefix=f"user_{uid}_")
        USERS[uid] = {
            "tmpdir": tmp,
            "excel_path": None,
            "pdfs": [],
            "preview": [],
            "generated": None,
        }
    return USERS[uid]

# PDF TEXT EXTRACTION
def extract_text_from_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            text += p.extract_text() or ""
    return text

# PARSE PDF CONTENT
def parse_pdf(path):
    base = os.path.basename(path)
    rfi = re.search(r"(\d{1,6})", base)
    rfi_num = rfi.group(1) if rfi else ""

    txt = " ".join(extract_text_from_pdf(path).split())

    drawing = re.search(r"([A-Z]{1,6}-\d{2,7})", txt)
    drawing = drawing.group(1) if drawing else ""

    desc = ""
    d = re.search(r"(Inspection Request for[^\n]{5,120})", txt)
    if d:
        desc = d.group(1).strip()

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

# UPDATE EXCEL FILE
def update_excel(template, out_path, rows):
    wb = openpyxl.load_workbook(template)
    ws = wb.active

    ws.insert_rows(START_ROW, len(rows))

    for i, r in enumerate(rows):
        row = START_ROW + i

        def sc(c, v):
            ws[f"{c}{row}"].value = v

        sc("C", f"IRFI-C-{r['rfi']}")
        sc("D", FIXED_PROJECT_NO)
        sc("E", FIXED_CLASSIFICATION)
        sc("F", FIXED_DISCIPLINE)
        sc("G", f"RFI-C-{r['rfi']}")
        sc("H", r["description"])
        sc("I", r["drawing"])
        sc("J", r["date"])
        sc("K", r["date"])

    wb.save(out_path)
    return out_path

# TELEGRAM HANDLERS
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ensure_user(update.effective_user.id)
    await update.message.reply_text(
        "👋 Welcome!\nSend your Excel file first, then upload the PDF files."
    )

async def handle_docs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    u = ensure_user(update.effective_user.id)
    doc = update.message.document
    f = await doc.get_file()

    dest = os.path.join(u["tmpdir"], doc.file_name)
    await f.download_to_drive(dest)

    if dest.endswith(".xlsx"):
        u["excel_path"] = dest
        await update.message.reply_text("📘 Excel uploaded.")
    elif dest.endswith(".pdf"):
        u["pdfs"].append(dest)
        await update.message.reply_text("📄 PDF uploaded.")
    else:
        await update.message.reply_text("❌ Only PDF and Excel files are supported.")
        return

    if u["excel_path"] and u["pdfs"]:
        kb = [
            [InlineKeyboardButton("Preview", callback_data="preview")],
            [InlineKeyboardButton("Generate Excel", callback_data="generate")],
            [InlineKeyboardButton("Download Excel", callback_data="download")],
        ]
        await update.message.reply_text(
            "Choose action:", reply_markup=InlineKeyboardMarkup(kb)
        )

async def buttons(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    u = ensure_user(q.from_user.id)

    if q.data == "preview":
        u["preview"] = [parse_pdf(p) for p in u["pdfs"]]
        txt = "\n".join(
            [
                f"Row {START_ROW+i}: RFI={r['rfi']} | {r['description'][:40]}"
                for i, r in enumerate(u["preview"])
            ]
        )
        await q.edit_message_text("🔍 Preview:\n" + txt)

    elif q.data == "generate":
        out_path = os.path.join(u["tmpdir"], "updated.xlsx")
        update_excel(u["excel_path"], out_path, u["preview"])
        u["generated"] = out_path
        await q.edit_message_text("✅ Excel generated successfully!")

    elif q.data == "download":
        if not u["generated"]:
            await q.edit_message_text("❌ Please generate the Excel first.")
            return

        with open(u["generated"], "rb") as f:
            await q.message.reply_document(InputFile(f, "updated.xlsx"))

        await q.edit_message_text("📤 File sent successfully!")

# START BOT
app = ApplicationBuilder().token(BOT_TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.Document.ALL, handle_docs))
app.add_handler(CallbackQueryHandler(buttons))

app.run_polling()

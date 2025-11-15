import os
from flask import Flask, request
from telegram import Update, Bot
from telegram.ext import Dispatcher, MessageHandler, Filters, CallbackContext

TOKEN = os.environ.get("BOT_TOKEN")
bot = Bot(token=TOKEN)

app = Flask(__name__)

dispatcher = Dispatcher(bot, update_queue=None, use_context=True)

# ---- YOUR MESSAGE HANDLER ----
def echo(update: Update, context: CallbackContext):
    update.message.reply_text("Bot is working on Render webhook!")

dispatcher.add_handler(MessageHandler(Filters.text & ~Filters.command, echo))

# ---- WEBHOOK ENDPOINT ----
@app.route(f"/webhook/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), bot)
    dispatcher.process_update(update)
    return "OK", 200

# ---- HEALTH CHECK ----
@app.route("/")
def home():
    return "Bot is running", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

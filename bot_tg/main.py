from dotenv import load_dotenv
from telegram.ext import ApplicationBuilder, CallbackQueryHandler, CommandHandler, MessageHandler, filters

from common.config import get_telegram_token
from common.logger import log
from bot_tg.handlers import start, help_command, status, sync, handle_text, handle_callback, handle_document


def run() -> None:
    load_dotenv()
    app = (
        ApplicationBuilder()
        .token(get_telegram_token())
        .concurrent_updates(True)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("sync", sync))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    log.info("Telegram bot started")
    app.run_polling()


if __name__ == "__main__":
    run()

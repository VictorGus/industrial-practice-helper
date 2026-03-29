from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

from common.config import get_telegram_token
from common.logger import log
from bot_tg.handlers import start, echo, handle_document


def run() -> None:
    app = (
        ApplicationBuilder()
        .token(get_telegram_token())
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, echo))

    log.info("Telegram bot started")
    app.run_polling()


if __name__ == "__main__":
    run()

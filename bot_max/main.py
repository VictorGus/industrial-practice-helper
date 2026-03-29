import aiomax

from common.config import get_max_token
from bot_max.handlers import setup_handlers


def run() -> None:
    bot = aiomax.Bot(get_max_token())
    setup_handlers(bot)

    print("Max bot started")
    bot.run()


if __name__ == "__main__":
    run()

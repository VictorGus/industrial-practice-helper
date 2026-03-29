import aiomax


def setup_handlers(bot: aiomax.Bot) -> None:
    @bot.on_bot_start()
    async def on_start(payload: aiomax.BotStartPayload):
        await payload.send("Hello! I'm your Max bot.")

    @bot.on_message()
    async def echo(message: aiomax.Message):
        await message.reply(message.body.text)

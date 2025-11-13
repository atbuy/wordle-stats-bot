import asyncio

from loguru import logger

from wordle.bot import WordleStatsBot
from wordle.cogs.events import EventsCog

bot = WordleStatsBot()


async def main() -> None:
    """Initializes and starts the bot."""

    await bot.add_cog(EventsCog(bot))


if __name__ == "__main__":
    try:
        asyncio.run(main())
        bot.run(bot.settings.token)
    except KeyboardInterrupt:
        logger.info("Exiting...")

import asyncio
from argparse import ArgumentParser
from datetime import datetime

from loguru import logger

from wordle.bot import WordleStatsBot
from wordle.cogs.events import EventsCog

bot = WordleStatsBot()


async def main(args) -> None:
    """Initializes and starts the bot."""

    await bot.add_cog(EventsCog(bot, args.m))


def cli():
    current_date = datetime.now().strftime("%Y-%m")

    parser = ArgumentParser()
    parser.add_argument(
        "-m",
        default=current_date,
        type=str,
        help="The month you want to check results for, in the format '%Y-%m', e.g. '2025-04'",
    )

    args = parser.parse_args()

    try:
        asyncio.run(main(args))
        bot.run(bot.settings.token)
    except KeyboardInterrupt:
        logger.info("Exiting...")


if __name__ == "__main__":
    cli()

import asyncio

import discord
from discord import Activity, ActivityType
from discord.ext import commands
from loguru import logger

from wordle.cogs.events import EventsCog
from wordle.settings import get_settings


async def main() -> None:
    """Initializes and starts the bot."""

    settings = get_settings()

    activity = Activity(
        type=ActivityType.listening,
        name=f"{settings.command_prefix}help",
    )

    intents = discord.Intents.default()

    bot = commands.Bot(
        command_prefix=settings.command_prefix,
        activity=activity,
        intents=intents,
    )

    await bot.add_cog(EventsCog(bot))

    await bot.start(settings.token)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Exiting...")

import asyncio
from datetime import datetime, timedelta, timezone

from discord.ext import commands
from loguru import logger

from wordle.settings import get_settings


class EventsCog(commands.Cog):
    """This is a cog for event listeners.

    Evnt listeners are executed automatically on a specific event.
    They can be used to handle errors, guild changes and more.
    """

    def __init__(self, bot: commands.Bot):
        self.bot = bot

    @commands.Cog.listener()
    async def on_ready(self) -> None:
        logger.info("Bot is ready.")

        asyncio.create_task(self.parse_wordle_stats())

    async def parse_wordle_stats(self):
        """Read messages from the wordle bot and generate statistics."""

        await asyncio.sleep(0)

        settings = get_settings()

        guild = self.bot.get_guild(settings.guild_id)
        if guild is None:
            logger.error("Guild not found. Is the bot in the server?")
            await self.bot.close()
            return

        channel = guild.get_channel(settings.channel_id)
        if channel is None:
            logger.error("Channel not found.")
            await self.bot.close()
            return

        now = datetime.now(timezone.utc)
        one_month_ago = now - timedelta(days=30)

        results = []

        channel_messages = channel.history(
            limit=None,
            after=one_month_ago,
            oldest_first=True,
        )

        async for message in channel_messages:
            if message.author.id != settings.app_id:
                continue

            content = message.content.strip()
            logger.info(f"Got message from '{message.author}': '{content}'")

            results.append(content)

        logger.info("Finished. Shutting down...")

        await self.bot.close()

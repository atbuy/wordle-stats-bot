import asyncio
import re
from datetime import datetime, timedelta, timezone

from discord.ext import commands
from loguru import logger

from wordle.settings import get_settings

POINT_MAP = {
    "1": 10,
    "2": 5,
    "3": 4,
    "4": 3,
    "5": 2,
    "6": 1,
    "X": -1,
}


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

        user_data = await self.get_user_data()
        if not user_data:
            await self.bot.close()
            return

        await self.generate_report(user_data)
        await self.bot.close()

    async def get_user_data(self) -> dict[str, dict[str, int]]:
        """Parse daily statistics and calculate total user score."""

        settings = get_settings()

        guild = self.bot.get_guild(settings.guild_id)
        if guild is None:
            logger.error("Guild not found. Is the bot in the server?")
            return {}

        channel = guild.get_channel(settings.channel_id)
        if channel is None:
            logger.error("Channel not found.")
            return {}

        now = datetime.now(timezone.utc)
        one_month_ago = now - timedelta(days=30)

        user_data = {}

        points_pattern = r".*([0-6X])/[0-6]:"
        users_pattern = r"<@(\d+)>"

        channel_messages = channel.history(
            limit=None,
            after=one_month_ago,
            oldest_first=True,
        )

        async for message in channel_messages:
            if message.author.id != settings.app_id:
                continue

            message_date = message.created_at
            date = ""
            if message_date is not None:
                date = message_date.strftime("%d %a")

            if date not in user_data:
                user_data[date] = {}

            content = message.content.strip()
            if "Here are yesterday's results:" not in content:
                continue

            lines = content.split("\n")
            for line in lines:
                line = line.strip()

                points = re.findall(points_pattern, line)
                if points == []:
                    continue

                score = POINT_MAP[points[0]]

                users = re.findall(users_pattern, line)
                for user in users:
                    if user not in user_data[date]:
                        user_data[date][user] = 0

                    user_data[date][user] += score

        logger.info(user_data)

        return user_data

    async def generate_report(self, user_data: dict[str, dict[str, int]]) -> None:
        """Create excel sheet with user data."""

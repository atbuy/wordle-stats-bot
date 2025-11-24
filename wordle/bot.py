import discord
from discord import Activity, ActivityType
from discord.ext import commands

from wordle.settings import get_settings


class WordleStatsBot(commands.Bot):
    def __init__(self, *args, **kwargs):
        self.settings = get_settings()

        self._prefix = self.settings.command_prefix
        self._activity = Activity(
            type=ActivityType.listening,
            name=f"{self._prefix}help",
        )

        self._intents = discord.Intents.all()

        super().__init__(
            *args,
            command_prefix=self.settings.command_prefix,
            activity=self._activity,
            intents=self._intents,
            **kwargs,
        )

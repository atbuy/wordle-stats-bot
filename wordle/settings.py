import logging
from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="WORDLE_", env_nested_delimiter="__")

    token: str
    guild_id: int
    channel_id: int
    app_id: int
    command_prefix: str = "&"
    log_level: int = logging.INFO


@lru_cache
def get_settings() -> Settings:
    return Settings()

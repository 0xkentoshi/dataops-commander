from functools import lru_cache
from pathlib import Path
import os

from pydantic import Field, SecretStr
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    telegram_bot_token: SecretStr = Field(
        validation_alias="TELEGRAM_BOT_TOKEN"
    )
    ollama_host: str = Field(
        default="http://localhost:11434",
        validation_alias="OLLAMA_HOST",
    )
    ollama_model: str = Field(
        default="qwen3:8b",
        validation_alias="OLLAMA_MODEL",
    )
    allowed_telegram_user_ids: str = Field(
        default="",
        validation_alias="ALLOWED_TELEGRAM_USER_IDS",
    )
    data_workspace_dir: str = Field(
        default="workspace_files",
        validation_alias="DATA_WORKSPACE_DIR",
    )

    @property
    def allowed_user_ids(self) -> set[int]:
        raw_values = self.allowed_telegram_user_ids.split(",")
        values = {value.strip() for value in raw_values if value.strip()}

        try:
            return {int(value) for value in values}
        except ValueError as error:
            raise ValueError(
                "ALLOWED_TELEGRAM_USER_IDS должен содержать Telegram ID "
                "через запятую"
            ) from error

    @property
    def workspace_directory(self) -> Path:
        expanded = os.path.expandvars(self.data_workspace_dir.strip())
        if not expanded:
            expanded = "workspace_files"
        path = Path(expanded).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        resolved = path.resolve()
        if resolved == Path(resolved.anchor):
            raise ValueError(
                "DATA_WORKSPACE_DIR не может указывать на корень диска. "
                "Выберите отдельную папку с рабочими данными."
            )
        return resolved


@lru_cache
def get_settings() -> Settings:
    return Settings()

"""Application settings — reads from .env file."""

import os
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict

BASE_DIR = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=str(BASE_DIR / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_name: str = "Mannings Social Dashboard"
    debug: bool = True

    excel_path: str = str(BASE_DIR / "Mannings_FB_IG_Dashboard_Feed.xlsx")
    competitors_dir: str = str(BASE_DIR / "competitors")
    screenshots_config: str = str(BASE_DIR / "screenshots_config.yaml")

    gemini_api_key: str = ""
    gemini_model: str = "gemini-2.0-flash"

    # LLM via OpenRouter (free models, works in HK)
    # Get free key: https://openrouter.ai/keys
    openrouter_api_key: str = ""
    openrouter_base_url: str = "https://openrouter.ai/api/v1"
    openrouter_model: str = "deepseek/deepseek-chat-v3-0324:free"
    openrouter_referer: str = "https://localhost:8000"
    openrouter_title: str = "Mannings Dashboard"

    fpk_email: str = ""
    fpk_password: str = ""


settings = Settings()

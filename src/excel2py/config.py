"""Configuration management for excel2py."""

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env", env_prefix="EXCEL2PY_", extra="ignore"
    )

    default_provider: str = "openai"
    openai_api_key: str | None = None
    openai_model: str = "gpt-4o"
    anthropic_api_key: str | None = None
    anthropic_model: str = "claude-sonnet-4-20250514"
    google_api_key: str | None = None
    google_model: str = "gemini-2.5-flash"
    openrouter_api_key: str | None = None
    openrouter_model: str = "openai/gpt-4o"
    agno_api_key: str | None = None
    agno_model: str = "claude-sonnet-4-5-20250929"
    correction_backend: str = "langchain"  # "langchain" or "agno"
    temperature: float = 0.2
    max_retries: int = 3
    log_level: str = "INFO"
    verify: bool = True
    max_verify_attempts: int = 5
    verify_timeout: int = 60


def get_settings() -> Settings:
    return Settings()

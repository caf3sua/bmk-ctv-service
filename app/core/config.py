import os
from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    MONGODB_URI: str = "mongodb://localhost:27017"
    DB_NAME: str = "bmk_ctv"
    PORT: int = 8000
    HOST: str = "0.0.0.0"

    SECRET_KEY: str = "change-me-to-a-random-secret"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 1440

    CORS_ORIGINS: str = "http://localhost:5173,http://127.0.0.1:5173"

    # Logging
    LOG_LEVEL: str = "INFO"  # DEBUG, INFO, WARNING, ERROR
    LOG_FORMAT: str = "text"  # "text" hoặc "json"
    LOG_FILE: str | None = None  # Nếu có, log cũng được ghi ra file này (rotating)
    LOG_MAX_BYTES: int = 1_000_000
    LOG_BACKUP_COUNT: int = 5

    # Google Social Login (ID token flow — only GOOGLE_CLIENT_ID is used for verification;
    # GOOGLE_CLIENT_SECRET is stored for completeness in case a server-side auth-code flow is added later)
    GOOGLE_CLIENT_ID: str = ""
    GOOGLE_CLIENT_SECRET: str = ""

    # S3 MinIO Configurations
    S3_ACCESS_KEY_ID: str = ""
    S3_ACCESS_SECRET: str = ""
    S3_BUCKET: str = "intranet"
    S3_ENDPOINT: str = ""
    S3_REGION: str = "hanoi"
    S3_VERIFY: bool = False

    model_config = SettingsConfigDict(
        env_file=os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), ".env"),
        env_file_encoding="utf-8",
        extra="ignore"
    )

    @property
    def cors_origins_list(self) -> list[str]:
        return [origin.strip() for origin in self.CORS_ORIGINS.split(",") if origin.strip()]

settings = Settings()

"""Centralized logging setup for the application.

Provides structured (JSON) or plain text logging, optional rotating file handler,
and a convenience function to fetch module-specific loggers.
"""
from __future__ import annotations

import json
import logging
import os
import time
from logging.handlers import RotatingFileHandler
from typing import Any, Dict

from app.core.config import settings

_LOGGING_INITIALIZED = False

_EXTRA_ATTRS = ["request_id", "path", "method", "status_code", "duration_ms"]


class JSONLogFormatter(logging.Formatter):
    """Format log records as JSON for easier ingestion by log processors."""

    def format(self, record: logging.LogRecord) -> str:  # type: ignore[override]
        base: Dict[str, Any] = {
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(record.created)),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        for attr in _EXTRA_ATTRS:
            if hasattr(record, attr):
                base[attr] = getattr(record, attr)
        if record.exc_info:
            base["exc_info"] = self.formatException(record.exc_info)
        return json.dumps(base, ensure_ascii=False)


def _make_formatter() -> logging.Formatter:
    if settings.LOG_FORMAT.lower() == "json":
        return JSONLogFormatter()
    return logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _build_handlers() -> list[logging.Handler]:
    handlers: list[logging.Handler] = []

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(_make_formatter())
    handlers.append(console_handler)

    if settings.LOG_FILE:
        log_dir = os.path.dirname(settings.LOG_FILE)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)
        file_handler = RotatingFileHandler(
            settings.LOG_FILE,
            maxBytes=settings.LOG_MAX_BYTES,
            backupCount=settings.LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setFormatter(_make_formatter())
        handlers.append(file_handler)

    return handlers


def setup_logging() -> None:
    """Configure root logging according to settings. Safe to call multiple times."""
    global _LOGGING_INITIALIZED
    if _LOGGING_INITIALIZED:
        return
    level = getattr(logging, settings.LOG_LEVEL.upper(), logging.INFO)
    logging.basicConfig(level=level, handlers=_build_handlers(), force=True)
    _LOGGING_INITIALIZED = True


def get_logger(name: str) -> logging.Logger:
    """Return a logger, ensuring logging has been initialized."""
    if not _LOGGING_INITIALIZED:
        setup_logging()
    return logging.getLogger(name)

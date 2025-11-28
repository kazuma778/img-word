"""Konfigurasi aplikasi dan utilitas pemuatan environment."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / ".env"
if ENV_PATH.exists():
    load_dotenv(ENV_PATH)


@dataclass(frozen=True)
class Config:
    """Konfigurasi dasar yang dapat dioverride via variabel environment."""

    SECRET_KEY: str = os.getenv("SECRET_KEY", "dev-secret-key")
    UPLOAD_FOLDER: str = os.getenv("UPLOAD_FOLDER", str(BASE_DIR / "uploads"))
    PROCESSED_FOLDER: str = os.getenv("PROCESSED_FOLDER", str(BASE_DIR / "processed"))
    MAX_FILES: int = int(os.getenv("MAX_PROCESSED_FILES", "5"))
    MAX_UPLOAD_FILES: int = int(os.getenv("MAX_UPLOAD_FILES", "500"))
    MAX_CONTENT_LENGTH: int = int(os.getenv("MAX_CONTENT_LENGTH", str(25 * 1024 * 1024)))
    ALLOWED_EXTENSIONS: tuple[str, ...] = tuple(
        filter(None, os.getenv("ALLOWED_EXTENSIONS", "docx,pdf,doc,jpg,jpeg,png,rtf,bmp,webp,tiff").split(","))
    )
    IMG_UPSCALER_JWT: str = os.getenv("IMG_UPSCALER_JWT", "")
    IMG_UPSCALER_SCALE: str = os.getenv("IMG_UPSCALER_SCALE", "400")
    ENABLE_FILE_BROWSER: bool = os.getenv("ENABLE_FILE_BROWSER", "false").lower() == "true"
    WORKER_POOL_SIZE: int = int(os.getenv("WORKER_POOL_SIZE", "4"))

    @property
    def allowed_extensions_set(self) -> set[str]:
        return {ext.strip().lower() for ext in self.ALLOWED_EXTENSIONS if ext.strip()}

    def to_flask_config(self) -> Dict[str, Any]:
        """Konversi ke dict untuk `app.config`."""
        return {
            "SECRET_KEY": self.SECRET_KEY,
            "UPLOAD_FOLDER": self.UPLOAD_FOLDER,
            "PROCESSED_FOLDER": self.PROCESSED_FOLDER,
            "MAX_FILES": self.MAX_FILES,
            "MAX_UPLOAD_FILES": self.MAX_UPLOAD_FILES,
            "ALLOWED_EXTENSIONS": self.allowed_extensions_set,
            "MAX_CONTENT_LENGTH": self.MAX_CONTENT_LENGTH,
            "IMG_UPSCALER_JWT": self.IMG_UPSCALER_JWT,
            "IMG_UPSCALER_SCALE": self.IMG_UPSCALER_SCALE,
            "ENABLE_FILE_BROWSER": self.ENABLE_FILE_BROWSER,
            "WORKER_POOL_SIZE": self.WORKER_POOL_SIZE,
        }


def get_config() -> Config:
    """Helper untuk mendapatkan instance Config (bisa dimemoisasi bila perlu)."""
    return Config()


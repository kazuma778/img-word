"""Utility terkait penyimpanan/upload."""

from __future__ import annotations

import os
import shutil
import time
from pathlib import Path
from typing import Iterable, Mapping, MutableMapping

from flask import current_app

ConfigType = Mapping[str, object] | MutableMapping[str, object]


def _get_cfg(config: ConfigType | None = None) -> ConfigType:
    return config or current_app.config


def allowed_file(
    filename: str,
    allowed_extensions: Iterable[str] | None = None,
    config: ConfigType | None = None,
) -> bool:
    cfg = _get_cfg(config)
    allowed = set(allowed_extensions or cfg["ALLOWED_EXTENSIONS"])
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


def safe_remove_file(file_path: str, retry: int = 3, delay: float = 0.5) -> bool:
    """Hapus file dengan retry untuk menangani race condition di Windows."""
    for _ in range(retry):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            return True
        except OSError:
            time.sleep(delay)
    return False


def _cleanup_dir(target: str, max_items: int) -> None:
    entries = []
    for name in os.listdir(target):
        path = os.path.join(target, name)
        try:
            entries.append((path, os.path.getmtime(path)))
        except FileNotFoundError:
            continue

    entries.sort(key=lambda item: item[1], reverse=True)
    for path, _ in entries[max_items:]:
        try:
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            else:
                os.remove(path)
        except OSError:
            continue


def cleanup_processed_folder(config: ConfigType | None = None) -> None:
    cfg = _get_cfg(config)
    Path(cfg["PROCESSED_FOLDER"]).mkdir(parents=True, exist_ok=True)
    _cleanup_dir(cfg["PROCESSED_FOLDER"], cfg["MAX_FILES"])


def cleanup_uploads_folder(config: ConfigType | None = None) -> None:
    cfg = _get_cfg(config)
    Path(cfg["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
    _cleanup_dir(cfg["UPLOAD_FOLDER"], cfg["MAX_UPLOAD_FILES"])


def cleanup_all_folders(config: ConfigType | None = None) -> None:
    cleanup_uploads_folder(config)
    cleanup_processed_folder(config)


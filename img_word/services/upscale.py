"""Integrasi dengan layanan upscaler eksternal."""

from __future__ import annotations

import os
import time
from typing import Optional, Tuple

import requests
from flask import current_app


UPLOAD_URL = "https://get1.imglarger.com/api/UpscalerNew/UploadNew"
STATUS_URL = "https://get1.imglarger.com/api/UpscalerNew/CheckStatusNew"


def _build_headers() -> dict[str, str]:
    cookie = current_app.config.get("IMG_UPSCALER_JWT", "")
    return {
        "Cookie": f"jwt={cookie}",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            " AppleWebKit/537.36 (KHTML, like Gecko)"
            " Chrome/115.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
        "Origin": "https://imgupscaler.com",
        "Referer": "https://imgupscaler.com/",
    }


def upload_image(file_path: str) -> Tuple[Optional[str], Optional[str]]:
    if not os.path.exists(file_path):
        return None, None

    headers = _build_headers()
    scale = current_app.config.get("IMG_UPSCALER_SCALE", "400")
    scale_map = {"200": "1", "400": "4"}
    scale_value = scale_map.get(scale, "4")

    with open(file_path, "rb") as handle:
        files = {"myfile": (os.path.basename(file_path), handle, "image/jpeg")}
        data = {"scaleRadio": scale_value}
        response = requests.post(UPLOAD_URL, headers=headers, files=files, data=data, timeout=60)
        response.raise_for_status()
        payload = response.json()

    if payload.get("msg") != "Success":
        return None, None

    task_id = payload.get("data", {}).get("code")
    return task_id, scale_value


def check_status(task_id: str, scale_value: str, timeout_seconds: int = 300) -> Optional[str]:
    headers = _build_headers()
    start = time.time()
    while time.time() - start < timeout_seconds:
        data = {"code": task_id, "scaleRadio": scale_value}
        response = requests.post(STATUS_URL, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        payload = response.json()

        if payload.get("msg") != "Success":
            return None

        status_data = payload.get("data", {})
        status = status_data.get("status")
        if status == "success":
            urls = status_data.get("downloadUrls", [])
            return urls[0] if urls else None
        if status == "fail":
            return None
        time.sleep(10)
    return None


def download_result(url: str, output_path: str) -> bool:
    headers = {"User-Agent": _build_headers()["User-Agent"]}
    response = requests.get(url, headers=headers, stream=True, timeout=60)
    response.raise_for_status()
    with open(output_path, "wb") as handle:
        for chunk in response.iter_content(8192):
            handle.write(chunk)
    return True


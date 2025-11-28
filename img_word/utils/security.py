"""Helper keamanan sederhana."""

from __future__ import annotations

from functools import wraps
from typing import Callable, TypeVar, Any, cast

from flask import abort, current_app, request

F = TypeVar("F", bound=Callable[..., Any])


def require_file_browser_token(func: F) -> F:
    """Lindungi endpoint file browser dengan token sederhana."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any):
        token = current_app.config.get("FILE_BROWSER_TOKEN")
        if not current_app.config.get("ENABLE_FILE_BROWSER", False):
            abort(404)
        provided = request.headers.get("X-File-Token") or request.args.get("token")
        if not token or provided != token:
            abort(403)
        return func(*args, **kwargs)

    return cast(F, wrapper)


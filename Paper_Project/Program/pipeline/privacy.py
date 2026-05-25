"""
privacy.py - small helpers for writing local QA/profile reports without
leaking absolute machine paths.

The pipeline still keeps full paths in raw local JSON when needed for the
builder. Public-facing reports should pass through these helpers first.
"""
from __future__ import annotations

import os
import re
import tempfile
from typing import Any


def _norm(path: str) -> str:
    return str(path or "").replace("\\", "/")


def sanitize_path(value: Any, project_root: str | None = None) -> Any:
    """Return a display-safe version of a filesystem path-like value."""
    if not isinstance(value, str):
        return value

    text = _norm(value)
    if not text:
        return text

    roots = []
    if project_root:
        roots.append(("PROJECT", project_root))
    roots.append(("TEMP", tempfile.gettempdir()))
    home = os.path.expanduser("~")
    if home and home != "~":
        roots.append(("HOME", home))

    for label, root in roots:
        root_norm = _norm(os.path.abspath(root)).rstrip("/")
        if text.lower().startswith(root_norm.lower()):
            rel = text[len(root_norm):].lstrip("/")
            return f"<{label}>/{rel}" if rel else f"<{label}>"

    drive_match = re.match(r"^[A-Za-z]:/", text)
    if drive_match:
        parts = text.split("/")
        if len(parts) >= 2 and parts[1] in {"Inputs", "Outputs", "Templates"}:
            return "/".join(parts[1:])
        return "<ABS_PATH>/" + "/".join(parts[-2:])

    return text


def sanitize_value(value: Any, project_root: str | None = None) -> Any:
    """Recursively sanitize path-like strings inside JSON-compatible values."""
    if isinstance(value, dict):
        return {k: sanitize_value(v, project_root) for k, v in value.items()}
    if isinstance(value, list):
        return [sanitize_value(v, project_root) for v in value]
    if isinstance(value, str):
        return sanitize_path(value, project_root)
    return value

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


def _replace_root_path(text: str, label: str, root: str) -> str:
    root_norm = _norm(os.path.abspath(root)).rstrip("/")
    if not root_norm:
        return text
    pattern = re.compile(re.escape(root_norm) + r"(?P<tail>(?:/[^\s`'\"<>|]+)*)", re.I)

    def repl(match: re.Match[str]) -> str:
        tail = (match.group("tail") or "").lstrip("/")
        return f"<{label}>/{tail}" if tail else f"<{label}>"

    return pattern.sub(repl, text)


def _replace_windows_abs_path(text: str) -> str:
    pattern = re.compile(r"(?<![A-Za-z0-9_])([A-Za-z]:/[^\s`'\"<>|]+)")

    def repl(match: re.Match[str]) -> str:
        path = match.group(1)
        parts = path.split("/")
        if len(parts) >= 3 and parts[1] in {"Inputs", "Outputs", "Templates"}:
            return "/".join(parts[1:])
        return "<ABS_PATH>/" + "/".join(parts[-2:])

    return pattern.sub(repl, text)


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
        text = _replace_root_path(text, label, root)

    return _replace_windows_abs_path(text)


def sanitize_value(value: Any, project_root: str | None = None) -> Any:
    """Recursively sanitize path-like strings inside JSON-compatible values."""
    if isinstance(value, dict):
        return {k: sanitize_value(v, project_root) for k, v in value.items()}
    if isinstance(value, list):
        return [sanitize_value(v, project_root) for v in value]
    if isinstance(value, str):
        return sanitize_path(value, project_root)
    return value

"""Template manifest, download, and local file helpers for the public suite."""
from __future__ import annotations

import hashlib
import json
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional

from public_template_suite_modules.paths import FILES_DIR, MANIFEST_PATH, ROOT, TEST_ROOT
from public_template_suite_modules.scenarios import DEFAULT_TEMPLATES


def safe_download_url(value: Any) -> bool:
    if not isinstance(value, str) or not value.startswith("https://"):
        return False
    if any(token in value for token in ("????", "\ufffd")):
        return False
    try:
        value.encode("ascii")
    except UnicodeEncodeError:
        return False
    return True


def read_manifest() -> Dict[str, Any]:
    default_by_id = {str(item.get("id")): item for item in DEFAULT_TEMPLATES}
    if MANIFEST_PATH.exists():
        with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("templates"):
            merged = []
            for item in data.get("templates") or []:
                item_id = str(item.get("id") or "")
                base = dict(default_by_id.get(item_id, {}))
                clean = {}
                for key, value in item.items():
                    if value in (None, ""):
                        continue
                    if key == "download_url" and not safe_download_url(value):
                        continue
                    clean[key] = value
                base.update(clean)
                merged.append(base)
            seen = {str(item.get("id")) for item in merged}
            for item in DEFAULT_TEMPLATES:
                if str(item.get("id")) not in seen:
                    merged.append(dict(item))
            data["templates"] = merged
            return data
    return {"schema_version": 1, "templates": [dict(item) for item in DEFAULT_TEMPLATES]}


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def expected_sha256(item: Dict[str, Any]) -> str:
    value = item.get("sha256") or item.get("expected_sha256")
    if not value:
        return ""
    value = str(value).strip().lower()
    if len(value) != 64 or any(ch not in "0123456789abcdef" for ch in value):
        raise RuntimeError(f"invalid sha256 for {item.get('id')}")
    return value


def verify_expected_sha256(path: Path, expected: str) -> None:
    if expected and sha256_file(path).lower() != expected:
        raise RuntimeError(f"local template sha256 mismatch: {path.name}")


def resolve_existing_file(item: Dict[str, Any]) -> Optional[Path]:
    raw = item.get("file")
    if not raw:
        return None
    path = Path(str(raw))
    candidates = [path]
    if not path.is_absolute():
        candidates.extend([ROOT / path, TEST_ROOT / path, FILES_DIR / path.name])
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def download_template(item: Dict[str, Any], force: bool = False) -> Path:
    expected = expected_sha256(item)
    existing = resolve_existing_file(item)
    if existing and existing.exists() and not force:
        verify_expected_sha256(existing, expected)
        return existing

    FILES_DIR.mkdir(parents=True, exist_ok=True)
    if item.get("name"):
        name = str(item.get("name"))
    elif item.get("file"):
        name = Path(str(item.get("file"))).name
    else:
        name = f"{item.get('id', 'template')}.docx"
    path = FILES_DIR / name
    if path.exists() and path.stat().st_size > 2000 and not force:
        verify_expected_sha256(path, expected)
        return path
    url = item.get("download_url") or item.get("url")
    if not safe_download_url(url):
        raise RuntimeError(f"missing or unsafe ASCII download_url for {item.get('id')}")
    req = urllib.request.Request(str(url), headers={"User-Agent": "word-chat-template-test/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = resp.read()
    if len(data) < 2000 or not data.startswith(b"PK"):
        raise RuntimeError(f"downloaded file is not a DOCX zip: {len(data)} bytes")
    if expected and hashlib.sha256(data).hexdigest().lower() != expected:
        raise RuntimeError(f"downloaded file sha256 mismatch for {item.get('id')}")
    path.write_bytes(data)
    return path


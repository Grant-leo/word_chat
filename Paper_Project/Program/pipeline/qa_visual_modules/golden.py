"""Golden baseline helpers for visual QA."""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from typing import Any, Dict, List


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _hash_text(text: str) -> str:
    return hashlib.sha256(str(text or "").encode("utf-8")).hexdigest()


def _hamming_hex(a: str, b: str) -> int:
    try:
        return bin(int(a, 16) ^ int(b, 16)).count("1")
    except Exception:
        return 999


def _baseline_key(out_dir: str) -> str:
    parts: List[str] = []
    for name in ("format.json", "content.json"):
        path = os.path.join(out_dir, name)
        try:
            data = _load_json(path)
            meta = data.get("_meta") or {}
            parts.append(str(meta.get("sha256") or meta.get("source") or name))
        except Exception:
            parts.append(name)
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()[:24]


def _golden_record(out_dir: str, counts: Dict[str, Any], pages_text: List[str], image_stats: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "key": _baseline_key(out_dir),
        "counts": {
            "pages": counts.get("pages"),
            "page_width_pt": counts.get("page_width_pt"),
            "page_height_pt": counts.get("page_height_pt"),
            "text_pages": counts.get("text_pages"),
            "all_page_images": counts.get("all_page_images"),
        },
        "rendered_text_hash": _hash_text("\f".join(pages_text)),
        "page_hashes": image_stats.get("page_hashes") or [],
    }


def _compare_or_update_golden(
    out_dir: str,
    counts: Dict[str, Any],
    pages_text: List[str],
    image_stats: Dict[str, Any],
    golden_dir: str | None,
    update_golden: bool,
) -> Dict[str, Any]:
    result: Dict[str, Any] = {"enabled": bool(golden_dir), "status": "disabled"}
    if not golden_dir:
        return result
    os.makedirs(golden_dir, exist_ok=True)
    record = _golden_record(out_dir, counts, pages_text, image_stats)
    path = os.path.join(golden_dir, record["key"] + ".json")
    result.update({"status": "missing", "path": path, "key": record["key"], "issues": []})
    if update_golden:
        existed = os.path.exists(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        result["status"] = "updated" if existed else "created"
        return result
    if not os.path.exists(path):
        result["issues"] = ["baseline missing; rerun with --update-golden only after manual approval"]
        return result
    baseline = _load_json(path)
    issues: List[str] = []
    for key in ("pages", "page_width_pt", "page_height_pt", "text_pages", "all_page_images"):
        if baseline.get("counts", {}).get(key) != record.get("counts", {}).get(key):
            issues.append(f"{key}: {record.get('counts', {}).get(key)} != {baseline.get('counts', {}).get(key)}")
    if baseline.get("rendered_text_hash") != record.get("rendered_text_hash"):
        issues.append("rendered text hash changed")
    old_hashes = baseline.get("page_hashes") or []
    new_hashes = record.get("page_hashes") or []
    if len(old_hashes) != len(new_hashes):
        issues.append(f"page hash count: {len(new_hashes)} != {len(old_hashes)}")
    else:
        changed = [idx + 1 for idx, pair in enumerate(zip(old_hashes, new_hashes)) if _hamming_hex(pair[0], pair[1]) > 8]
        if changed:
            issues.append("page image hashes changed: " + ",".join(map(str, changed[:12])))
    result["status"] = "matched" if not issues else "mismatch"
    result["issues"] = issues
    return result

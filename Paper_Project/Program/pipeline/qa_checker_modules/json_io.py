"""JSON artifact helpers for structural QA."""
from __future__ import annotations

import json
import os
from typing import Any, Dict

def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_manifest_counts(out_dir: str) -> Dict[str, Any]:
    path = os.path.join(out_dir, "build_manifest.json")
    if not os.path.exists(path):
        return {}
    data = _load_json(path)
    return data.get("counts") or {}

"""Rendered-page image statistics for visual QA."""
from __future__ import annotations

from typing import Any, Dict, List


def _image_stats(image_paths: List[str]) -> Dict[str, Any]:
    stats: Dict[str, Any] = {"blank_pages": [], "page_hashes": []}
    try:
        from PIL import Image, ImageStat
    except Exception:
        stats["pil_available"] = False
        return stats
    stats["pil_available"] = True
    for idx, path in enumerate(image_paths, 1):
        try:
            with Image.open(path) as im:
                gray = im.convert("L")
                extrema = gray.getextrema()
                if extrema and extrema[1] - extrema[0] < 3:
                    stats["blank_pages"].append(idx)
                small = gray.resize((8, 8))
                values = list(small.getdata())
                avg = sum(values) / max(len(values), 1)
                bits = "".join("1" if v >= avg else "0" for v in values)
                stats["page_hashes"].append(hex(int(bits, 2))[2:].zfill(16))
                if idx == 1:
                    stat = ImageStat.Stat(gray)
                    stats["first_page_mean"] = round(float(stat.mean[0]), 2)
        except Exception:
            stats.setdefault("unreadable_pages", []).append(idx)
    return stats


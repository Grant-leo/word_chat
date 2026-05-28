"""Test harness helpers for regression_suite.py."""
from __future__ import annotations

import argparse
import base64
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any, Callable, Dict, List


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lc0ndwAAAABJRU5ErkJggg=="
)

CASES: List[Callable[[], None]] = []
TEMP_DIRS: List[Path] = []
KEEP_ARTIFACTS = False


def case(fn: Callable[[], None]) -> Callable[[], None]:
    CASES.append(fn)
    return fn


def fail(message: str) -> None:
    raise AssertionError(message)


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def new_workdir(name: str) -> Path:
    work = Path(tempfile.mkdtemp(prefix=f"wordchat_{name}_"))
    TEMP_DIRS.append(work)
    return work


def write_json(path: Path, value: Dict[str, Any]) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def write_sample_png(path: Path, width: int = 480, height: int = 270) -> None:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((8, 8, width - 8, height - 8), outline=(30, 80, 160), width=3)
    draw.line((24, height - 40, width - 24, 40), fill=(180, 50, 50), width=4)
    image.save(path)


def base_format(source: str = "synthetic.md") -> Dict[str, Any]:
    return {
        "_meta": {
            "source": source,
            "sha256": "synthetic",
            "paragraphs": 1,
            "tables": 0,
            "sections": 1,
        },
        "paragraphs": [
            {
                "text": "Synthetic body paragraph used for style inference.",
                "runs": [{"font": "Times New Roman", "size_pt": 12}],
                "align": "JUSTIFY",
            }
        ],
        "tables": [],
        "sections": [
            {
                "page_width_cm": 21.0,
                "page_height_cm": 29.7,
                "margin_top_cm": 2.54,
                "margin_bottom_cm": 2.54,
                "margin_left_cm": 3.17,
                "margin_right_cm": 3.17,
            }
        ],
        "cover": [],
        "style_profiles": {},
    }


def base_content(paragraphs: List[Any], meta_tables: int = 0) -> Dict[str, Any]:
    return {
        "_meta": {
            "source": "synthetic.docx",
            "sha256": "synthetic",
            "paragraphs": max(1, len(paragraphs)),
            "tables_count": meta_tables,
            "images_extracted": 0,
        },
        "title_info": {"title_cn": "Synthetic Thesis"},
        "sections": [
            {
                "heading": "1 Introduction",
                "level": 1,
                "role": "body",
                "paragraphs": paragraphs,
                "images": [],
            }
        ],
        "references": ["[1] Synthetic reference."],
    }


def run_cases(selected: str | None = None) -> int:
    passed = 0
    failed = 0
    matched = 0
    for fn in CASES:
        if selected and selected not in fn.__name__:
            continue
        matched += 1
        try:
            fn()
            print(f"PASS {fn.__name__}")
            passed += 1
        except Exception as exc:
            failed += 1
            print(f"FAIL {fn.__name__}: {exc}")
    if selected and matched == 0:
        print("RESULT passed=0 failed=1")
        print(f"FAIL no tests matched filter: {selected}")
        return 1
    print(f"RESULT passed={passed} failed={failed}")
    return 1 if failed else 0


def main() -> None:
    global KEEP_ARTIFACTS
    parser = argparse.ArgumentParser(description="Run synthetic regression checks for the Word pipeline.")
    parser.add_argument("--keep", action="store_true", help="Keep temporary test artifacts.")
    parser.add_argument("--filter", default=None, help="Run cases whose function name contains this text.")
    args = parser.parse_args()
    KEEP_ARTIFACTS = bool(args.keep)
    try:
        code = run_cases(args.filter)
    finally:
        if not KEEP_ARTIFACTS and "code" in locals() and code == 0:
            for path in TEMP_DIRS:
                shutil.rmtree(path, ignore_errors=True)
        elif TEMP_DIRS:
            print("ARTIFACTS")
            for path in TEMP_DIRS:
                print(path)
    raise SystemExit(code)

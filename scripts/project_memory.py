"""Validate and append events to the repository memory bank."""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
MEMORY_DIR = ROOT / "memory"
STREAM_FILES = {
    "decisions": MEMORY_DIR / "decisions.jsonl",
    "progress": MEMORY_DIR / "progress.jsonl",
    "risks": MEMORY_DIR / "risks.jsonl",
}
REQUIRED_FILES = [
    MEMORY_DIR / "PROJECT_MEMORY.md",
    MEMORY_DIR / "active_context.md",
    MEMORY_DIR / "index.json",
    *STREAM_FILES.values(),
]


def _now() -> str:
    return datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")


def _parse_tags(raw: str) -> List[str]:
    return [part.strip() for part in (raw or "").split(",") if part.strip()]


def _read_jsonl(path: Path) -> Iterable[Dict[str, Any]]:
    if not path.exists():
        return []
    entries = []
    for line_no, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{path}:{line_no}: invalid JSONL: {exc}") from exc
        if not isinstance(value, dict):
            raise ValueError(f"{path}:{line_no}: entry must be an object")
        entries.append(value)
    return entries


def validate() -> None:
    missing = [str(path.relative_to(ROOT)) for path in REQUIRED_FILES if not path.exists()]
    if missing:
        raise SystemExit("Missing memory files: " + ", ".join(missing))

    index = json.loads((MEMORY_DIR / "index.json").read_text(encoding="utf-8"))
    if index.get("version") != 1:
        raise SystemExit("memory/index.json must declare version 1")

    for stream, path in STREAM_FILES.items():
        for entry in _read_jsonl(path):
            for key in ("timestamp", "stream", "title", "body", "tags", "source"):
                if key not in entry:
                    raise SystemExit(f"{path.relative_to(ROOT)} entry missing {key}: {entry}")
            if entry["stream"] != stream:
                raise SystemExit(f"{path.relative_to(ROOT)} has stream={entry['stream']!r}, expected {stream!r}")
            if not isinstance(entry["tags"], list):
                raise SystemExit(f"{path.relative_to(ROOT)} entry tags must be a list: {entry}")
            datetime.fromisoformat(str(entry["timestamp"]))

    print("memory validation OK")


def add_event(args: argparse.Namespace) -> None:
    path = STREAM_FILES[args.stream]
    entry = {
        "timestamp": args.timestamp or _now(),
        "stream": args.stream,
        "title": args.title.strip(),
        "body": args.body.strip(),
        "tags": _parse_tags(args.tags),
        "source": args.source.strip(),
    }
    if not entry["title"] or not entry["body"]:
        raise SystemExit("--title and --body are required")
    if not entry["source"]:
        raise SystemExit("--source is required")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False, separators=(",", ":")) + "\n")
    validate()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("validate", help="validate the memory bank")

    add = sub.add_parser("add", help="append a normalized memory event")
    add.add_argument("--stream", choices=sorted(STREAM_FILES), required=True)
    add.add_argument("--title", required=True)
    add.add_argument("--body", required=True)
    add.add_argument("--tags", default="")
    add.add_argument("--source", required=True)
    add.add_argument("--timestamp", default="")

    args = parser.parse_args()
    if args.cmd == "validate":
        validate()
    elif args.cmd == "add":
        add_event(args)


if __name__ == "__main__":
    main()

"""
template_profiler.py - stable entry point for template capability profiles.

The implementation lives in template_profiler_modules/ so callers can keep
using profile_format(), report_to_markdown(), and write_profile().
"""
from __future__ import annotations

import json
import os

try:
    from template_profiler_modules.profiles import STYLE_ROLES, profile_format
    from template_profiler_modules.reports import report_to_markdown, write_profile
except ImportError:  # pragma: no cover - package-style imports
    from .template_profiler_modules.profiles import STYLE_ROLES, profile_format
    from .template_profiler_modules.reports import report_to_markdown, write_profile

__all__ = ["STYLE_ROLES", "profile_format", "report_to_markdown", "write_profile"]


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Build a template profile from format.json.")
    parser.add_argument("format_json")
    parser.add_argument("--out-dir", default=None)
    args = parser.parse_args()

    with open(args.format_json, "r", encoding="utf-8") as f:
        fmt_obj = json.load(f)
    out = args.out_dir or os.path.dirname(os.path.abspath(args.format_json)) or "."
    result = write_profile(fmt_obj, out)
    print(report_to_markdown(result))

"""Shared paths for the public template compatibility suite."""
from __future__ import annotations

from pathlib import Path


PIPELINE_DIR = Path(__file__).resolve().parents[1]
ROOT = PIPELINE_DIR.parents[2]
TEST_ROOT = ROOT / "TestData" / "PublicTemplates"
FILES_DIR = TEST_ROOT / "files"
RUNS_DIR = TEST_ROOT / "runs"
MANIFEST_PATH = TEST_ROOT / "manifest.json"
SELECTION_SUMMARY_PATH = RUNS_DIR / "last_selection_summary.json"

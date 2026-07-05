"""Subprocess execution helpers for generated pipeline scripts."""
from __future__ import annotations

from dataclasses import dataclass
import os
import subprocess


@dataclass(frozen=True)
class ScriptExecutionResult:
    returncode: int
    stdout: str
    stderr: str


def _unsafe_unicode_decode_calls(gen_py_path):
    try:
        from qa_checker_modules.unicode_decode_guard import unsafe_unicode_decode_calls
    except Exception:
        return []
    try:
        return unsafe_unicode_decode_calls(gen_py_path)
    except Exception:
        return []


def run_generated_script(gen_py_path, out_dir, python_executable):
    unsafe_decode_calls = _unsafe_unicode_decode_calls(gen_py_path)
    if unsafe_decode_calls:
        return ScriptExecutionResult(
            returncode=2,
            stdout="",
            stderr="GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE: " + ", ".join(str(item) for item in unsafe_decode_calls),
        )
    result = subprocess.run(
        [python_executable, gen_py_path],
        capture_output=True,
        cwd=out_dir,
        env={**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONDONTWRITEBYTECODE": "1"},
    )
    return ScriptExecutionResult(
        returncode=result.returncode,
        stdout=(result.stdout or b"").decode("utf-8", errors="replace"),
        stderr=(result.stderr or b"").decode("utf-8", errors="replace"),
    )

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


def run_generated_script(gen_py_path, out_dir, python_executable):
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

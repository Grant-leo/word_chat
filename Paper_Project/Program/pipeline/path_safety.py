"""Safety checks for generated-output cleanup paths."""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Iterable


_SOURCE_DIR_NAMES = {"Inputs", "Templates"}


def _resolve(path: str | os.PathLike[str]) -> Path:
    return Path(path).expanduser().resolve()


def _is_relative_to(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


def ensure_safe_output_dir(output_dir: str | os.PathLike[str]) -> str:
    """Return an absolute output dir and reject known source-data trees."""
    resolved = _resolve(output_dir)
    if any(part in _SOURCE_DIR_NAMES for part in resolved.parts):
        raise ValueError(f"Unsafe output_dir points inside a source directory: {resolved}")
    return str(resolved)


def safe_rmtree_generated_child(
    target: str | os.PathLike[str],
    output_dir: str | os.PathLike[str],
    *,
    allowed_names: Iterable[str] | None = None,
    allowed_suffixes: Iterable[str] | None = None,
) -> None:
    """Remove a generated child directory only after boundary/name checks."""
    output_root = _resolve(ensure_safe_output_dir(output_dir))
    target_path = _resolve(target)
    if target_path == output_root or not _is_relative_to(target_path, output_root):
        raise ValueError(f"Unsafe generated cleanup target outside output_dir: {target_path}")

    names = set(allowed_names or [])
    suffixes = tuple(allowed_suffixes or ())
    if names or suffixes:
        name = target_path.name
        if name not in names and not any(name.endswith(suffix) for suffix in suffixes):
            raise ValueError(f"Unsafe generated cleanup target name: {name}")

    shutil.rmtree(target_path, ignore_errors=True)

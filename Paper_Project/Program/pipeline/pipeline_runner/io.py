"""CLI and filesystem helpers for the one-click pipeline runner."""
from __future__ import annotations

import os


def scan_inputs(folder, exts=(".docx", ".md")):
    """Return input files in a folder, excluding Word temp files."""
    if not os.path.isdir(folder):
        return []
    files = [
        f
        for f in os.listdir(folder)
        if any(f.endswith(e) for e in exts) and not f.startswith("~$")
    ]
    return sorted(files)


def choose_file(files, label):
    """Interactive file selection. Returns the chosen filename."""
    if len(files) == 0:
        return None
    if len(files) == 1:
        print(f"{label}: {files[0]} (自动选择)")
        return files[0]
    print(f"\n{label}:")
    for i, f in enumerate(files, 1):
        print(f"  [{i}] {f}")
    while True:
        try:
            choice = input(f"请选择 (1-{len(files)}): ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
            print(f"  输入无效，请输入 1-{len(files)}")
        except (ValueError, KeyboardInterrupt):
            print("\n  已取消")
            raise SystemExit(1)


def normalize_mode(mode):
    mode = (mode or "user").strip().lower()
    return mode if mode in ("user", "developer") else "user"


def choose_mode(default="user"):
    """Interactive workflow mode selection."""
    print("\n选择工作模式:")
    print("  [1] 普通用户：AI 只修改本次输出目录的 build_generated.py")
    print("  [2] 开发者：AI 只修改 Paper_Project/Program/pipeline/ 核心引擎脚本")
    prompt = f"请选择 (1-2，默认 {1 if default == 'user' else 2}): "
    try:
        choice = input(prompt).strip()
    except KeyboardInterrupt:
        print("\n  已取消")
        raise SystemExit(1)
    if not choice:
        return normalize_mode(default)
    if choice in ("2", "developer", "dev", "开发者"):
        return "developer"
    return "user"


def exit_from_result(result):
    raise SystemExit(0 if result else 1)

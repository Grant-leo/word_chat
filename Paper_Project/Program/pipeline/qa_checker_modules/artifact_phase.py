"""Artifact and manifest checks for structural QA."""
from __future__ import annotations

import ast
import os
from typing import Any, Callable, Dict, List

try:
    from qa_checker_modules.metrics import _load_json, _load_manifest_counts
    from qa_checker_modules.registry import VALID_MODES
except ImportError:  # pragma: no cover - package-style imports
    from .metrics import _load_json, _load_manifest_counts
    from .registry import VALID_MODES

AddIssue = Callable[..., None]
DANGEROUS_UNICODE_CODECS = {"unicode_escape", "raw_unicode_escape"}


def build_output_paths(out_dir: str, output_docx_name: str) -> Dict[str, str]:
    return {
        "docx": os.path.join(out_dir, output_docx_name),
        "build": os.path.join(out_dir, "build_generated.py"),
        "format": os.path.join(out_dir, "format.json"),
        "content": os.path.join(out_dir, "content.json"),
        "workflow": os.path.join(out_dir, "workflow_mode.json"),
        "manifest": os.path.join(out_dir, "build_manifest.json"),
    }


def _string_constant(node: ast.AST) -> str:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    return ""


def _call_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        parent = _call_name(node.value)
        return f"{parent}.{node.attr}" if parent else node.attr
    return ""


def _keyword_value(node: ast.Call, *names: str) -> str:
    wanted = set(names)
    for keyword in node.keywords or []:
        if keyword.arg in wanted:
            return _string_constant(keyword.value)
    return ""


def _unsafe_unicode_decode_calls(path: str) -> List[str]:
    try:
        text = open(path, "r", encoding="utf-8", errors="ignore").read()
        tree = ast.parse(text, filename=path)
    except Exception:
        return []
    hits: List[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = _call_name(node.func)
        if name == "codecs.decode":
            hits.append("codecs.decode")
            continue
        if name.endswith(".decode"):
            encoding = ""
            if node.args:
                encoding = _string_constant(node.args[0])
            encoding = encoding or _keyword_value(node, "encoding")
            if encoding in DANGEROUS_UNICODE_CODECS:
                hits.append(f"{name}({encoding})")
            continue
        if name.endswith("escape_decode"):
            hits.append(name)
    return sorted(set(hits))


def run_artifact_checks(out_dir: str, paths: Dict[str, str], counts: Dict[str, Any], add: AddIssue) -> Dict[str, Any]:
    if not os.path.exists(paths["workflow"]):
        add("WORKFLOW_MODE_INVALID", "warning", "未找到 workflow_mode.json，无法确认本轮应按用户模式还是开发者模式修复。")
    else:
        try:
            workflow = _load_json(paths["workflow"])
            if workflow.get("mode") not in VALID_MODES:
                add("WORKFLOW_MODE_INVALID", "warning", "workflow_mode.json 中的 mode 无效。", str(workflow.get("mode")))
        except Exception as exc:
            add("WORKFLOW_MODE_INVALID", "warning", "workflow_mode.json 无法读取。", str(exc))

    for key, code, message in [
        ("docx", "MISSING_DOCX", "缺少最终 docx。"),
        ("build", "MISSING_BUILD_SCRIPT", "缺少 build_generated.py。"),
        ("format", "MISSING_FORMAT_JSON", "缺少 format.json。"),
        ("content", "MISSING_CONTENT_JSON", "缺少 content.json。"),
    ]:
        if not os.path.exists(paths[key]):
            add(code, "error", message, paths[key])

    if os.path.exists(paths["build"]):
        unsafe_decode_calls = _unsafe_unicode_decode_calls(paths["build"])
        if unsafe_decode_calls:
            counts["unsafe_unicode_decode_calls"] = len(unsafe_decode_calls)
            add(
                "GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE",
                "error",
                "生成脚本包含可能破坏中文字符的 unicode-escape 解码调用。",
                ", ".join(unsafe_decode_calls),
            )

    manifest_counts: Dict[str, Any] = {}
    if os.path.exists(paths["manifest"]):
        try:
            manifest_counts = _load_manifest_counts(out_dir)
            for key, value in manifest_counts.items():
                counts[f"manifest_{key}"] = value
        except Exception as exc:
            add("BUILD_MANIFEST_MISSING", "warning", "build_manifest.json 无法读取，正文元素数量只能退回 XML 总量检测。", str(exc))
    elif os.path.exists(paths["docx"]):
        add("BUILD_MANIFEST_MISSING", "warning", "未找到 build_manifest.json，图片/表格数量检测可能被封面或页眉元素干扰。")
    return manifest_counts


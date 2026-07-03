"""AST helpers for detecting unsafe unicode-escape decoding."""
from __future__ import annotations

import ast
import codecs
from typing import List, Set, Tuple

DANGEROUS_UNICODE_CODECS = {"unicode-escape", "raw-unicode-escape"}


def _normalize_codec_name(name: str) -> str:
    if not name:
        return ""
    try:
        return codecs.lookup(name).name
    except LookupError:
        return name.strip().lower().replace("_", "-").replace(" ", "-")


def _is_dangerous_unicode_codec(name: str) -> bool:
    return _normalize_codec_name(name) in DANGEROUS_UNICODE_CODECS


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


def _codec_import_aliases(tree: ast.AST) -> Tuple[Set[str], Set[str], Set[str], Set[str], Set[str]]:
    module_aliases = {"codecs"}
    decode_aliases: Set[str] = set()
    escape_decode_aliases: Set[str] = set()
    getdecoder_aliases: Set[str] = set()
    lookup_aliases: Set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "codecs":
                    module_aliases.add(alias.asname or alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module == "codecs":
            for alias in node.names:
                imported_name = alias.asname or alias.name
                if alias.name == "decode":
                    decode_aliases.add(imported_name)
                elif alias.name == "escape_decode":
                    escape_decode_aliases.add(imported_name)
                elif alias.name == "getdecoder":
                    getdecoder_aliases.add(imported_name)
                elif alias.name == "lookup":
                    lookup_aliases.add(imported_name)
    return module_aliases, decode_aliases, escape_decode_aliases, getdecoder_aliases, lookup_aliases


def _attribute_on_module_alias(node: ast.AST, module_aliases: Set[str], attr: str) -> bool:
    return (
        isinstance(node, ast.Attribute)
        and node.attr == attr
        and _call_name(node.value) in module_aliases
    )


def _add_simple_factory_assignment_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
) -> None:
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        if value is None:
            continue
        value_name = _call_name(value)
        is_getdecoder = value_name in getdecoder_aliases or _attribute_on_module_alias(value, module_aliases, "getdecoder")
        is_lookup = value_name in lookup_aliases or _attribute_on_module_alias(value, module_aliases, "lookup")
        if not is_getdecoder and not is_lookup:
            continue
        for target in targets:
            if isinstance(target, ast.Name):
                if is_getdecoder:
                    getdecoder_aliases.add(target.id)
                if is_lookup:
                    lookup_aliases.add(target.id)


def _method_decode_encoding(node: ast.Call) -> str:
    if node.args:
        return _string_constant(node.args[0])
    return _keyword_value(node, "encoding")


def _codecs_decode_encoding(node: ast.Call) -> str:
    if len(node.args) >= 2:
        return _string_constant(node.args[1])
    return _keyword_value(node, "encoding")


def _codec_factory_encoding(node: ast.Call) -> str:
    if node.args:
        return _string_constant(node.args[0])
    return _keyword_value(node, "encoding")


def unsafe_unicode_decode_calls_from_text(text: str, filename: str = "<generated>") -> List[str]:
    tree = ast.parse(text, filename=filename)
    module_aliases, decode_aliases, escape_decode_aliases, getdecoder_aliases, lookup_aliases = _codec_import_aliases(tree)
    _add_simple_factory_assignment_aliases(tree, module_aliases, getdecoder_aliases, lookup_aliases)
    hits: List[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = _call_name(node.func)
        if isinstance(node.func, ast.Attribute) and node.func.attr == "decode":
            base = _call_name(node.func.value)
            if base in module_aliases:
                encoding = _codecs_decode_encoding(node)
                if _is_dangerous_unicode_codec(encoding):
                    hits.append(f"{name}({encoding})")
                continue
            encoding = _method_decode_encoding(node)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
            continue
        if name in decode_aliases:
            encoding = _codecs_decode_encoding(node)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
            continue
        if name in escape_decode_aliases or name.endswith("escape_decode"):
            hits.append(name)
            continue
        if (
            name in getdecoder_aliases
            or name in lookup_aliases
            or _attribute_on_module_alias(node.func, module_aliases, "getdecoder")
            or _attribute_on_module_alias(node.func, module_aliases, "lookup")
        ):
            encoding = _codec_factory_encoding(node)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
    return sorted(set(hits))


def unsafe_unicode_decode_calls(path: str) -> List[str]:
    try:
        text = open(path, "r", encoding="utf-8", errors="ignore").read()
        return unsafe_unicode_decode_calls_from_text(text, filename=path)
    except Exception:
        return []

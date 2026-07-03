"""AST helpers for detecting unsafe generated-script text decoding."""
from __future__ import annotations

import ast
import codecs
from typing import Dict, List, Optional, Set, Tuple

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


def _codec_label(name: str) -> str:
    return name or "<dynamic>"


def _static_string_value(node: ast.AST, constants: Dict[str, str]) -> Optional[str]:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.Name):
        return constants.get(node.id)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        left = _static_string_value(node.left, constants)
        right = _static_string_value(node.right, constants)
        if left is not None and right is not None:
            return left + right
    if isinstance(node, ast.JoinedStr):
        parts: List[str] = []
        for value in node.values:
            if isinstance(value, ast.Constant) and isinstance(value.value, str):
                parts.append(value.value)
            elif isinstance(value, ast.FormattedValue):
                item = _static_string_value(value.value, constants)
                if item is None:
                    return None
                parts.append(item)
            else:
                return None
        return "".join(parts)
    return None


def _string_constant(node: ast.AST, constants: Dict[str, str] | None = None) -> str:
    value = _static_string_value(node, constants or {})
    if value is not None:
        return value
    return ""


def _constant_string_aliases(tree: ast.AST) -> Dict[str, str]:
    constants: Dict[str, str] = {}
    for _ in range(4):
        changed = False
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
            item = _static_string_value(value, constants)
            if item is None:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and constants.get(target.id) != item:
                    constants[target.id] = item
                    changed = True
        if not changed:
            break
    return constants


def _call_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        parent = _call_name(node.value)
        return f"{parent}.{node.attr}" if parent else node.attr
    return ""


def _keyword_value(node: ast.Call, constants: Dict[str, str], *names: str) -> str:
    wanted = set(names)
    for keyword in node.keywords or []:
        if keyword.arg in wanted:
            return _string_constant(keyword.value, constants)
    return ""


def _keyword_node(node: ast.Call, *names: str) -> Optional[ast.AST]:
    wanted = set(names)
    for keyword in node.keywords or []:
        if keyword.arg in wanted:
            return keyword.value
    return None


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


def _add_simple_decode_assignment_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
) -> None:
    for _ in range(4):
        changed = False
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
            is_decode = value_name in decode_aliases or _attribute_on_module_alias(value, module_aliases, "decode")
            if not is_decode:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and target.id not in decode_aliases:
                    decode_aliases.add(target.id)
                    changed = True
        if not changed:
            break


def _method_decode_encoding(node: ast.Call, constants: Dict[str, str]) -> str:
    if node.args:
        return _string_constant(node.args[0], constants)
    return _keyword_value(node, constants, "encoding")


def _codecs_decode_encoding(node: ast.Call, constants: Dict[str, str]) -> str:
    if len(node.args) >= 2:
        return _string_constant(node.args[1], constants)
    return _keyword_value(node, constants, "encoding")


def _codec_factory_encoding(node: ast.Call, constants: Dict[str, str]) -> str:
    if node.args:
        return _string_constant(node.args[0], constants)
    return _keyword_value(node, constants, "encoding")


def _method_decode_encoding_node(node: ast.Call) -> Optional[ast.AST]:
    if node.args:
        return node.args[0]
    return _keyword_node(node, "encoding")


def _codecs_decode_encoding_node(node: ast.Call) -> Optional[ast.AST]:
    if len(node.args) >= 2:
        return node.args[1]
    return _keyword_node(node, "encoding")


def _codec_factory_encoding_node(node: ast.Call) -> Optional[ast.AST]:
    if node.args:
        return node.args[0]
    return _keyword_node(node, "encoding")


def _function_signature(node: ast.FunctionDef, constants: Dict[str, str]) -> Dict[str, object]:
    positional = [arg.arg for arg in [*node.args.posonlyargs, *node.args.args]]
    kwonly = [arg.arg for arg in node.args.kwonlyargs]
    defaults: Dict[str, str] = {}
    if node.args.defaults:
        default_names = positional[-len(node.args.defaults):]
        for name, default in zip(default_names, node.args.defaults):
            value = _static_string_value(default, constants)
            if value is not None:
                defaults[name] = value
    for name, default in zip(kwonly, node.args.kw_defaults):
        if default is None:
            continue
        value = _static_string_value(default, constants)
        if value is not None:
            defaults[name] = value
    return {
        "positions": {name: index for index, name in enumerate(positional)},
        "params": set(positional + kwonly),
        "defaults": defaults,
    }


def _codec_arg_param_name(codec_node: Optional[ast.AST], params: Set[str]) -> Optional[str]:
    if isinstance(codec_node, ast.Name) and codec_node.id in params:
        return codec_node.id
    return None


def _codec_wrapper_functions(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
) -> Dict[str, Dict[str, object]]:
    wrappers: Dict[str, Dict[str, object]] = {}
    for func in [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)]:
        signature = _function_signature(func, constants)
        params = signature["params"]
        if not isinstance(params, set):
            continue
        codec_params: Set[str] = set()
        for node in ast.walk(func):
            if not isinstance(node, ast.Call):
                continue
            name = _call_name(node.func)
            codec_node: Optional[ast.AST] = None
            if isinstance(node.func, ast.Attribute) and node.func.attr == "decode":
                base = _call_name(node.func.value)
                codec_node = _codecs_decode_encoding_node(node) if base in module_aliases else _method_decode_encoding_node(node)
            elif name in decode_aliases:
                codec_node = _codecs_decode_encoding_node(node)
            elif (
                name in getdecoder_aliases
                or name in lookup_aliases
                or _attribute_on_module_alias(node.func, module_aliases, "getdecoder")
                or _attribute_on_module_alias(node.func, module_aliases, "lookup")
            ):
                codec_node = _codec_factory_encoding_node(node)
            param = _codec_arg_param_name(codec_node, params)
            if param:
                codec_params.add(param)
        if codec_params:
            wrappers[func.name] = {
                "codec_params": codec_params,
                "positions": signature["positions"],
                "defaults": signature["defaults"],
            }
    return wrappers


def _add_simple_wrapper_assignment_aliases(tree: ast.AST, wrappers: Dict[str, Dict[str, object]]) -> None:
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        value_name = _call_name(value)
        if value_name not in wrappers:
            continue
        for target in targets:
            if isinstance(target, ast.Name):
                wrappers[target.id] = wrappers[value_name]


def _higher_order_decode_wrappers(tree: ast.AST, constants: Dict[str, str]) -> Dict[str, Dict[str, object]]:
    wrappers: Dict[str, Dict[str, object]] = {}
    for func in [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)]:
        signature = _function_signature(func, constants)
        params = signature["params"]
        if not isinstance(params, set):
            continue
        decode_arg_pairs: Set[Tuple[str, str]] = set()
        for node in ast.walk(func):
            if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name):
                continue
            decoder_param = node.func.id
            if decoder_param not in params:
                continue
            codec_param = _codec_arg_param_name(_codecs_decode_encoding_node(node), params)
            if codec_param:
                decode_arg_pairs.add((decoder_param, codec_param))
        if decode_arg_pairs:
            wrappers[func.name] = {
                "decode_arg_pairs": decode_arg_pairs,
                "positions": signature["positions"],
                "defaults": signature["defaults"],
            }
    return wrappers


def _wrapper_call_arg_node(node: ast.Call, wrapper: Dict[str, object], param_name: str) -> Optional[ast.AST]:
    positions = wrapper.get("positions") or {}
    if isinstance(positions, dict):
        position = positions.get(param_name)
        if isinstance(position, int) and position < len(node.args):
            return node.args[position]
    return _keyword_node(node, param_name)


def _wrapper_call_encoding(node: ast.Call, wrapper: Dict[str, object], param_name: str, constants: Dict[str, str]) -> str:
    defaults = wrapper.get("defaults") or {}
    value = _wrapper_call_arg_node(node, wrapper, param_name)
    if value is not None:
        return _string_constant(value, constants)
    if isinstance(defaults, dict):
        default = defaults.get(param_name)
        if isinstance(default, str):
            return default
    return ""


def _is_decode_function_ref(node: Optional[ast.AST], module_aliases: Set[str], decode_aliases: Set[str]) -> bool:
    if node is None:
        return False
    name = _call_name(node)
    return name in decode_aliases or _attribute_on_module_alias(node, module_aliases, "decode")


def unsafe_unicode_decode_calls_from_text(text: str, filename: str = "<generated>") -> List[str]:
    tree = ast.parse(text, filename=filename)
    constants = _constant_string_aliases(tree)
    module_aliases, decode_aliases, escape_decode_aliases, getdecoder_aliases, lookup_aliases = _codec_import_aliases(tree)
    _add_simple_decode_assignment_aliases(tree, module_aliases, decode_aliases)
    _add_simple_factory_assignment_aliases(tree, module_aliases, getdecoder_aliases, lookup_aliases)
    wrappers = _codec_wrapper_functions(tree, module_aliases, decode_aliases, getdecoder_aliases, lookup_aliases, constants)
    _add_simple_wrapper_assignment_aliases(tree, wrappers)
    higher_order_wrappers = _higher_order_decode_wrappers(tree, constants)
    _add_simple_wrapper_assignment_aliases(tree, higher_order_wrappers)
    hits: List[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = _call_name(node.func)
        if isinstance(node.func, ast.Attribute) and node.func.attr == "decode":
            base = _call_name(node.func.value)
            if base in module_aliases:
                encoding = _codecs_decode_encoding(node, constants)
                hits.append(f"{name}({_codec_label(encoding)})")
                continue
            encoding = _method_decode_encoding(node, constants)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
            continue
        if name in decode_aliases:
            encoding = _codecs_decode_encoding(node, constants)
            hits.append(f"{name}({_codec_label(encoding)})")
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
            encoding = _codec_factory_encoding(node, constants)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
            continue
        wrapper = wrappers.get(name)
        if wrapper:
            codec_params = wrapper.get("codec_params") or set()
            if isinstance(codec_params, set):
                for param_name in codec_params:
                    encoding = _wrapper_call_encoding(node, wrapper, param_name, constants)
                    if _is_dangerous_unicode_codec(encoding):
                        hits.append(f"{name}({encoding})")
            continue
        higher_order_wrapper = higher_order_wrappers.get(name)
        if higher_order_wrapper:
            decode_arg_pairs = higher_order_wrapper.get("decode_arg_pairs") or set()
            if isinstance(decode_arg_pairs, set):
                for decoder_param, codec_param in decode_arg_pairs:
                    decoder_node = _wrapper_call_arg_node(node, higher_order_wrapper, decoder_param)
                    if not _is_decode_function_ref(decoder_node, module_aliases, decode_aliases):
                        continue
                    encoding = _wrapper_call_encoding(node, higher_order_wrapper, codec_param, constants)
                    hits.append(f"{name}({_codec_label(encoding)})")
    return sorted(set(hits))


def unsafe_unicode_decode_calls(path: str) -> List[str]:
    try:
        text = open(path, "r", encoding="utf-8", errors="ignore").read()
        return unsafe_unicode_decode_calls_from_text(text, filename=path)
    except Exception:
        return []

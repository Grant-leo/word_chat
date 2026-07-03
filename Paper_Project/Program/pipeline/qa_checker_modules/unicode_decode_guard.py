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


def _functools_import_aliases(tree: ast.AST) -> Tuple[Set[str], Set[str]]:
    module_aliases = {"functools"}
    partial_aliases: Set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "functools":
                    module_aliases.add(alias.asname or alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module == "functools":
            for alias in node.names:
                if alias.name == "partial":
                    partial_aliases.add(alias.asname or alias.name)
    return module_aliases, partial_aliases


def _importlib_import_aliases(tree: ast.AST) -> Tuple[Set[str], Set[str]]:
    module_aliases = {"importlib"}
    import_module_aliases: Set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "importlib":
                    module_aliases.add(alias.asname or alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module == "importlib":
            for alias in node.names:
                if alias.name == "import_module":
                    import_module_aliases.add(alias.asname or alias.name)
    return module_aliases, import_module_aliases


def _operator_import_aliases(tree: ast.AST) -> Tuple[Set[str], Set[str], Set[str]]:
    module_aliases = {"operator"}
    methodcaller_aliases: Set[str] = set()
    attrgetter_aliases: Set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "operator":
                    module_aliases.add(alias.asname or alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module == "operator":
            for alias in node.names:
                imported_name = alias.asname or alias.name
                if alias.name == "methodcaller":
                    methodcaller_aliases.add(imported_name)
                elif alias.name == "attrgetter":
                    attrgetter_aliases.add(imported_name)
    return module_aliases, methodcaller_aliases, attrgetter_aliases


def _builtins_constructor_aliases(tree: ast.AST) -> Tuple[Set[str], Set[str]]:
    module_aliases = {"builtins", "__builtins__"}
    constructor_aliases: Set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "builtins":
                    module_aliases.add(alias.asname or alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module == "builtins":
            for alias in node.names:
                if alias.name in {"bytes", "bytearray"}:
                    constructor_aliases.add(alias.asname or alias.name)
    return module_aliases, constructor_aliases


def _attribute_on_module_alias(node: ast.AST, module_aliases: Set[str], attr: str) -> bool:
    return (
        isinstance(node, ast.Attribute)
        and node.attr == attr
        and _call_name(node.value) in module_aliases
    )


def _getattr_on_module_alias(
    node: ast.AST,
    module_aliases: Set[str],
    attr: str,
    constants: Dict[str, str],
) -> bool:
    return (
        isinstance(node, ast.Call)
        and _call_name(node.func) == "getattr"
        and len(node.args) >= 2
        and _call_name(node.args[0]) in module_aliases
        and _string_constant(node.args[1], constants) == attr
    )


def _importlib_codecs_module_call(
    node: ast.AST,
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    constants: Dict[str, str],
) -> bool:
    return (
        isinstance(node, ast.Call)
        and len(node.args) >= 1
        and _string_constant(node.args[0], constants) == "codecs"
        and (
            _call_name(node.func) in import_module_aliases
            or _attribute_on_module_alias(node.func, importlib_module_aliases, "import_module")
        )
    )


def _globals_or_locals_module_subscript(
    node: ast.AST,
    module_aliases: Set[str],
    constants: Dict[str, str],
) -> bool:
    return (
        isinstance(node, ast.Subscript)
        and isinstance(node.value, ast.Call)
        and _call_name(node.value.func) in {"globals", "locals"}
        and _string_constant(node.slice, constants) in module_aliases
    )


def _literal_dict_subscript(
    node: ast.AST,
    aliases: Dict[str, Set[str]],
    constants: Dict[str, str],
) -> bool:
    return (
        isinstance(node, ast.Subscript)
        and isinstance(node.value, ast.Name)
        and _string_constant(node.slice, constants) in aliases.get(node.value.id, set())
    )


def _static_codecs_module_ref(
    node: ast.AST,
    module_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
) -> bool:
    if _call_name(node) in module_aliases:
        return True
    if (
        isinstance(node, ast.Call)
        and _call_name(node.func) == "__import__"
        and len(node.args) >= 1
        and _string_constant(node.args[0], constants) == "codecs"
    ):
        return True
    if _importlib_codecs_module_call(node, importlib_module_aliases, import_module_aliases, constants):
        return True
    if _globals_or_locals_module_subscript(node, module_aliases, constants):
        return True
    if _literal_dict_subscript(node, module_dict_aliases, constants):
        return True
    return False


def _getattr_on_static_codecs_module(
    node: ast.AST,
    module_aliases: Set[str],
    attr: str,
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
) -> bool:
    return (
        isinstance(node, ast.Call)
        and _call_name(node.func) == "getattr"
        and len(node.args) >= 2
        and _static_codecs_module_ref(
            node.args[0],
            module_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
        )
        and _string_constant(node.args[1], constants) == attr
    )


def _module_dict_decode_subscript(
    node: ast.AST,
    module_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
) -> bool:
    if not isinstance(node, ast.Subscript) or _string_constant(node.slice, constants) != "decode":
        return False
    value = node.value
    if isinstance(value, ast.Attribute) and value.attr == "__dict__":
        return _static_codecs_module_ref(
            value.value,
            module_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
        )
    if isinstance(value, ast.Call) and _call_name(value.func) == "vars" and value.args:
        return _static_codecs_module_ref(
            value.args[0],
            module_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
        )
    return False


def _static_codecs_decode_ref(
    node: Optional[ast.AST],
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
) -> bool:
    if node is None:
        return False
    name = _call_name(node)
    if name in decode_aliases:
        return True
    if isinstance(node, ast.Attribute) and node.attr == "decode":
        return _static_codecs_module_ref(
            node.value,
            module_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
        )
    if _getattr_on_static_codecs_module(
        node,
        module_aliases,
        "decode",
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
    ):
        return True
    if _module_dict_decode_subscript(
        node,
        module_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
    ):
        return True
    return _literal_dict_subscript(node, decode_dict_aliases, constants)


def _literal_dict_codec_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
) -> Tuple[Dict[str, Set[str]], Dict[str, Set[str]]]:
    module_dict_aliases: Dict[str, Set[str]] = {}
    decode_dict_aliases: Dict[str, Set[str]] = {}
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
            if not isinstance(value, ast.Dict):
                continue
            module_keys: Set[str] = set()
            decode_keys: Set[str] = set()
            for key_node, value_node in zip(value.keys, value.values):
                if key_node is None:
                    continue
                key = _string_constant(key_node, constants)
                if not key:
                    continue
                if _static_codecs_module_ref(
                    value_node,
                    module_aliases,
                    constants,
                    importlib_module_aliases,
                    import_module_aliases,
                    module_dict_aliases,
                ):
                    module_keys.add(key)
                if _static_codecs_decode_ref(
                    value_node,
                    module_aliases,
                    decode_aliases,
                    constants,
                    importlib_module_aliases,
                    import_module_aliases,
                    module_dict_aliases,
                    decode_dict_aliases,
                ):
                    decode_keys.add(key)
            if not module_keys and not decode_keys:
                continue
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                before_module = set(module_dict_aliases.get(target.id, set()))
                before_decode = set(decode_dict_aliases.get(target.id, set()))
                module_dict_aliases.setdefault(target.id, set()).update(module_keys)
                decode_dict_aliases.setdefault(target.id, set()).update(decode_keys)
                changed = changed or before_module != module_dict_aliases[target.id]
                changed = changed or before_decode != decode_dict_aliases[target.id]
        if not changed:
            break
    return module_dict_aliases, decode_dict_aliases


def _builtins_subscript_constructor(node: ast.AST, constants: Dict[str, str]) -> bool:
    return (
        isinstance(node, ast.Subscript)
        and _call_name(node.value) == "__builtins__"
        and _string_constant(node.slice, constants) in {"bytes", "bytearray"}
    )


def _bytearray_zero_arg_call(node: ast.AST) -> bool:
    return isinstance(node, ast.Call) and _call_name(node.func) == "bytearray" and not node.args and not node.keywords


def _static_byte_constructor_factory(node: ast.AST) -> bool:
    if (
        isinstance(node, ast.Call)
        and _call_name(node.func) == "type"
        and len(node.args) == 1
        and not node.keywords
    ):
        value = node.args[0]
        return (
            isinstance(value, ast.Constant)
            and isinstance(value.value, (bytes, bytearray))
        ) or _bytearray_zero_arg_call(value)
    if isinstance(node, ast.Attribute) and node.attr == "__class__":
        value = node.value
        return (
            isinstance(value, ast.Constant)
            and isinstance(value.value, (bytes, bytearray))
        ) or _bytearray_zero_arg_call(value)
    return False


def _is_byte_constructor_ref(
    node: ast.AST,
    aliases: Set[str],
    constants: Dict[str, str],
) -> bool:
    return (
        _call_name(node) in aliases
        or _builtins_subscript_constructor(node, constants)
        or _static_byte_constructor_factory(node)
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


def _add_simple_module_assignment_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
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
            if not _static_codecs_module_ref(
                value,
                module_aliases,
                constants,
                importlib_module_aliases,
                import_module_aliases,
                module_dict_aliases,
            ):
                continue
            for target in targets:
                if isinstance(target, ast.Name) and target.id not in module_aliases:
                    module_aliases.add(target.id)
                    changed = True
        if not changed:
            break


def _add_simple_decode_assignment_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
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
            is_decode = (
                _static_codecs_decode_ref(
                    value,
                    module_aliases,
                    decode_aliases,
                    constants,
                    importlib_module_aliases,
                    import_module_aliases,
                    module_dict_aliases,
                    decode_dict_aliases,
                )
                or _attribute_on_module_alias(value, module_aliases, "decode")
                or _getattr_on_module_alias(value, module_aliases, "decode", constants)
            )
            if not is_decode:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and target.id not in decode_aliases:
                    decode_aliases.add(target.id)
                    changed = True
        if not changed:
            break


def _factory_call_encoding(
    node: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
) -> Tuple[str, str]:
    if not isinstance(node, ast.Call):
        return "", ""
    name = _call_name(node.func)
    if name in getdecoder_aliases or _attribute_on_module_alias(node.func, module_aliases, "getdecoder"):
        return "getdecoder", _codec_factory_encoding(node, constants)
    if name in lookup_aliases or _attribute_on_module_alias(node.func, module_aliases, "lookup"):
        return "lookup", _codec_factory_encoding(node, constants)
    return "", ""


def _decoder_factory_result_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
) -> Tuple[Dict[str, str], Dict[str, str]]:
    decoder_results: Dict[str, str] = {}
    lookup_results: Dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        kind, encoding = _factory_call_encoding(value, module_aliases, getdecoder_aliases, lookup_aliases, constants)
        if not kind:
            continue
        for target in targets:
            if not isinstance(target, ast.Name):
                continue
            if kind == "getdecoder":
                decoder_results[target.id] = encoding
            elif kind == "lookup":
                lookup_results[target.id] = encoding
    return decoder_results, lookup_results


def _decoder_factory_container_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
    decoder_results: Dict[str, str],
    lookup_results: Dict[str, str],
) -> Tuple[Dict[str, Dict[Tuple[str, str], str]], Dict[str, Dict[Tuple[str, str], str]]]:
    decoder_containers: Dict[str, Dict[Tuple[str, str], str]] = {}
    lookup_containers: Dict[str, Dict[Tuple[str, str], str]] = {}
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
            if isinstance(value, ast.Name) and value.id in decoder_containers:
                decoder_items = dict(decoder_containers[value.id])
            else:
                decoder_items: Dict[Tuple[str, str], str] = {}
            if isinstance(value, ast.Name) and value.id in lookup_containers:
                lookup_items = dict(lookup_containers[value.id])
            else:
                lookup_items: Dict[Tuple[str, str], str] = {}
            if not decoder_items and not lookup_items:
                for key, item in _literal_container_items(value, constants):
                    if isinstance(item, ast.Name) and item.id in decoder_results:
                        decoder_items[key] = decoder_results[item.id]
                        continue
                    if isinstance(item, ast.Name) and item.id in lookup_results:
                        lookup_items[key] = lookup_results[item.id]
                        continue
                    kind, encoding = _factory_call_encoding(item, module_aliases, getdecoder_aliases, lookup_aliases, constants)
                    if kind == "getdecoder":
                        decoder_items[key] = encoding
                    elif kind == "lookup":
                        lookup_items[key] = encoding
            if not decoder_items and not lookup_items:
                continue
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                if decoder_items and decoder_containers.get(target.id) != decoder_items:
                    decoder_containers[target.id] = dict(decoder_items)
                    changed = True
                if lookup_items and lookup_containers.get(target.id) != lookup_items:
                    lookup_containers[target.id] = dict(lookup_items)
                    changed = True
        if not changed:
            break
    return decoder_containers, lookup_containers


def _decoder_factory_value_alias(
    value: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
    decoder_results: Dict[str, str],
    lookup_results: Dict[str, str],
    decoder_containers: Dict[str, Dict[Tuple[str, str], str]],
    lookup_containers: Dict[str, Dict[Tuple[str, str], str]],
    decoder_attributes: Dict[str, str],
    lookup_attributes: Dict[str, str],
) -> Tuple[str, str]:
    if isinstance(value, ast.Name):
        if value.id in decoder_results:
            return "getdecoder", decoder_results[value.id]
        if value.id in lookup_results:
            return "lookup", lookup_results[value.id]
    container_decoder = _container_subscript_lookup(value, decoder_containers, constants)
    if container_decoder is not None:
        return "getdecoder", container_decoder
    container_lookup = _container_subscript_lookup(value, lookup_containers, constants)
    if container_lookup is not None:
        return "lookup", container_lookup
    attribute_decoder = _attribute_alias_lookup(value, decoder_attributes)
    if attribute_decoder is not None:
        return "getdecoder", attribute_decoder
    attribute_lookup = _attribute_alias_lookup(value, lookup_attributes)
    if attribute_lookup is not None:
        return "lookup", attribute_lookup
    return _factory_call_encoding(value, module_aliases, getdecoder_aliases, lookup_aliases, constants)


def _simple_namespace_keyword_items(value: ast.AST) -> List[Tuple[str, ast.AST]]:
    if not isinstance(value, ast.Call):
        return []
    if _call_name(value.func) not in {"SimpleNamespace", "types.SimpleNamespace"}:
        return []
    return [(keyword.arg, keyword.value) for keyword in value.keywords or [] if keyword.arg]


def _decoder_factory_attribute_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
    decoder_results: Dict[str, str],
    lookup_results: Dict[str, str],
    decoder_containers: Dict[str, Dict[Tuple[str, str], str]],
    lookup_containers: Dict[str, Dict[Tuple[str, str], str]],
) -> Tuple[Dict[str, str], Dict[str, str]]:
    decoder_attributes: Dict[str, str] = {}
    lookup_attributes: Dict[str, str] = {}
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

            for target in targets:
                if isinstance(target, ast.Attribute):
                    kind, encoding = _decoder_factory_value_alias(
                        value,
                        module_aliases,
                        getdecoder_aliases,
                        lookup_aliases,
                        constants,
                        decoder_results,
                        lookup_results,
                        decoder_containers,
                        lookup_containers,
                        decoder_attributes,
                        lookup_attributes,
                    )
                    attr_name = _call_name(target)
                    if kind == "getdecoder" and decoder_attributes.get(attr_name) != encoding:
                        decoder_attributes[attr_name] = encoding
                        changed = True
                    elif kind == "lookup" and lookup_attributes.get(attr_name) != encoding:
                        lookup_attributes[attr_name] = encoding
                        changed = True
                    continue

                if not isinstance(target, ast.Name):
                    continue
                for attr, item in _simple_namespace_keyword_items(value):
                    kind, encoding = _decoder_factory_value_alias(
                        item,
                        module_aliases,
                        getdecoder_aliases,
                        lookup_aliases,
                        constants,
                        decoder_results,
                        lookup_results,
                        decoder_containers,
                        lookup_containers,
                        decoder_attributes,
                        lookup_attributes,
                    )
                    attr_name = f"{target.id}.{attr}"
                    if kind == "getdecoder" and decoder_attributes.get(attr_name) != encoding:
                        decoder_attributes[attr_name] = encoding
                        changed = True
                    elif kind == "lookup" and lookup_attributes.get(attr_name) != encoding:
                        lookup_attributes[attr_name] = encoding
                        changed = True
        if not changed:
            break
    return decoder_attributes, lookup_attributes


def _method_decode_encoding(node: ast.Call, constants: Dict[str, str]) -> str:
    if node.args:
        return _string_constant(node.args[0], constants)
    return _keyword_value(node, constants, "encoding")


def _method_decode_encoding_or_default(node: ast.Call, constants: Dict[str, str]) -> str:
    if node.args:
        return _string_constant(node.args[0], constants)
    keyword = _keyword_node(node, "encoding")
    if keyword is not None:
        return _string_constant(keyword, constants)
    return "utf-8"


def _method_encode_encoding_or_default(node: ast.Call, constants: Dict[str, str]) -> str:
    if node.args:
        return _string_constant(node.args[0], constants)
    keyword = _keyword_node(node, "encoding")
    if keyword is not None:
        return _string_constant(keyword, constants)
    return "utf-8"


def _byte_constructor_aliases(tree: ast.AST, constants: Dict[str, str]) -> Set[str]:
    builtins_module_aliases, builtins_import_aliases = _builtins_constructor_aliases(tree)
    aliases = {"bytes", "bytearray"} | builtins_import_aliases
    for module_alias in builtins_module_aliases:
        aliases.add(f"{module_alias}.bytes")
        aliases.add(f"{module_alias}.bytearray")
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
            is_constructor = _is_byte_constructor_ref(value, aliases, constants) or any(
                _getattr_on_module_alias(value, builtins_module_aliases, constructor, constants)
                for constructor in ("bytes", "bytearray")
            )
            if not is_constructor:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and target.id not in aliases:
                    aliases.add(target.id)
                    changed = True
        if not changed:
            break
    return aliases


def _text_encode_call_encoding(
    node: ast.AST,
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
) -> Optional[str]:
    if not isinstance(node, ast.Call):
        return None
    if isinstance(node.func, ast.Attribute) and node.func.attr == "encode":
        return _method_encode_encoding_or_default(node, constants)
    if isinstance(node.func, ast.Attribute) and node.func.attr == "__new__":
        if not _is_byte_constructor_ref(node.func.value, byte_constructor_aliases, constants):
            return None
        if len(node.args) >= 3:
            return _string_constant(node.args[2], constants)
        keyword = _keyword_node(node, "encoding")
        if keyword is not None:
            return _string_constant(keyword, constants)
        return None
    if not _is_byte_constructor_ref(node.func, byte_constructor_aliases, constants):
        return None
    if len(node.args) >= 2:
        return _string_constant(node.args[1], constants)
    keyword = _keyword_node(node, "encoding")
    if keyword is not None:
        return _string_constant(keyword, constants)
    return None


def _encoded_text_bytes_encoding(
    node: ast.AST,
    aliases: Dict[str, str],
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
) -> Optional[str]:
    direct = _text_encode_call_encoding(node, constants, byte_constructor_aliases)
    if direct is not None:
        return direct
    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Attribute) and node.func.attr == "tobytes":
            return _encoded_text_bytes_encoding(node.func.value, aliases, constants, byte_constructor_aliases)
        if _call_name(node.func) == "memoryview" and node.args:
            return _encoded_text_bytes_encoding(node.args[0], aliases, constants, byte_constructor_aliases)
    if isinstance(node, ast.Name):
        return aliases.get(node.id)
    return None


def _encoded_text_byte_aliases(
    tree: ast.AST,
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
) -> Dict[str, str]:
    aliases: Dict[str, str] = {}
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
            encoding = _encoded_text_bytes_encoding(value, aliases, constants, byte_constructor_aliases)
            if encoding is None:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and aliases.get(target.id) != encoding:
                    aliases[target.id] = encoding
                    changed = True
        if not changed:
            break
    return aliases


def _operator_methodcaller_decode_encoding(
    node: ast.AST,
    operator_module_aliases: Set[str],
    methodcaller_aliases: Set[str],
    constants: Dict[str, str],
) -> Optional[str]:
    if not isinstance(node, ast.Call):
        return None
    is_methodcaller = (
        _call_name(node.func) in methodcaller_aliases
        or _attribute_on_module_alias(node.func, operator_module_aliases, "methodcaller")
    )
    if not is_methodcaller or not node.args:
        return None
    if _string_constant(node.args[0], constants) != "decode":
        return None
    if len(node.args) >= 2:
        return _string_constant(node.args[1], constants)
    keyword = _keyword_node(node, "encoding")
    if keyword is not None:
        return _string_constant(keyword, constants)
    return "utf-8"


def _operator_attrgetter_decode_ref(
    node: ast.AST,
    operator_module_aliases: Set[str],
    attrgetter_aliases: Set[str],
    constants: Dict[str, str],
) -> bool:
    return (
        isinstance(node, ast.Call)
        and node.args
        and _string_constant(node.args[0], constants) == "decode"
        and (
            _call_name(node.func) in attrgetter_aliases
            or _attribute_on_module_alias(node.func, operator_module_aliases, "attrgetter")
        )
    )


def _decode_method_source_encoding(
    node: ast.AST,
    encoded_text_aliases: Dict[str, str],
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
    operator_module_aliases: Set[str],
    attrgetter_aliases: Set[str],
) -> Optional[str]:
    if isinstance(node, ast.Attribute) and node.attr == "decode":
        return _encoded_text_bytes_encoding(
            node.value,
            encoded_text_aliases,
            constants,
            byte_constructor_aliases,
        )
    if (
        isinstance(node, ast.Call)
        and _call_name(node.func) == "getattr"
        and len(node.args) >= 2
        and _string_constant(node.args[1], constants) == "decode"
    ):
        return _encoded_text_bytes_encoding(
            node.args[0],
            encoded_text_aliases,
            constants,
            byte_constructor_aliases,
        )
    if (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "__getattribute__"
        and node.args
        and _string_constant(node.args[0], constants) == "decode"
    ):
        return _encoded_text_bytes_encoding(
            node.func.value,
            encoded_text_aliases,
            constants,
            byte_constructor_aliases,
        )
    if (
        isinstance(node, ast.Call)
        and _operator_attrgetter_decode_ref(
            node.func,
            operator_module_aliases,
            attrgetter_aliases,
            constants,
        )
        and node.args
    ):
        return _encoded_text_bytes_encoding(
            node.args[0],
            encoded_text_aliases,
            constants,
            byte_constructor_aliases,
        )
    return None


def _decode_method_result_aliases(
    tree: ast.AST,
    encoded_text_aliases: Dict[str, str],
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
    operator_module_aliases: Set[str],
    attrgetter_aliases: Set[str],
) -> Dict[str, str]:
    aliases: Dict[str, str] = {}
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
            if isinstance(value, ast.Name) and value.id in aliases:
                encoded_as = aliases[value.id]
            else:
                encoded_as = _decode_method_source_encoding(
                    value,
                    encoded_text_aliases,
                    constants,
                    byte_constructor_aliases,
                    operator_module_aliases,
                    attrgetter_aliases,
                )
            if encoded_as is None:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and aliases.get(target.id) != encoded_as:
                    aliases[target.id] = encoded_as
                    changed = True
        if not changed:
            break
    return aliases


def _container_key_from_literal(node: ast.AST, constants: Dict[str, str]) -> Optional[Tuple[str, str]]:
    if isinstance(node, ast.Constant) and isinstance(node.value, int) and node.value >= 0:
        return ("index", str(node.value))
    string_value = _static_string_value(node, constants)
    if string_value is not None:
        return ("key", string_value)
    return None


def _literal_container_items(
    node: ast.AST,
    constants: Dict[str, str],
) -> List[Tuple[Tuple[str, str], ast.AST]]:
    if isinstance(node, (ast.List, ast.Tuple)):
        return [(("index", str(index)), item) for index, item in enumerate(node.elts)]
    if isinstance(node, ast.Dict):
        items: List[Tuple[Tuple[str, str], ast.AST]] = []
        for key_node, value_node in zip(node.keys, node.values):
            if key_node is None:
                continue
            key = _container_key_from_literal(key_node, constants)
            if key is not None:
                items.append((key, value_node))
        return items
    return []


def _codec_literal_container_aliases(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
) -> Tuple[Dict[str, Dict[Tuple[str, str], str]], Dict[str, Dict[Tuple[str, str], str]]]:
    module_containers: Dict[str, Dict[Tuple[str, str], str]] = {}
    decode_containers: Dict[str, Dict[Tuple[str, str], str]] = {}
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
            if isinstance(value, ast.Name) and value.id in module_containers:
                module_items = dict(module_containers[value.id])
            else:
                module_items: Dict[Tuple[str, str], str] = {}
            if isinstance(value, ast.Name) and value.id in decode_containers:
                decode_items = dict(decode_containers[value.id])
            else:
                decode_items: Dict[Tuple[str, str], str] = {}
            if not module_items and not decode_items:
                for key, item in _literal_container_items(value, constants):
                    if _static_codecs_module_ref(
                        item,
                        module_aliases,
                        constants,
                        importlib_module_aliases,
                        import_module_aliases,
                        module_dict_aliases,
                    ):
                        module_items[key] = "codecs"
                    if _static_codecs_decode_ref(
                        item,
                        module_aliases,
                        decode_aliases,
                        constants,
                        importlib_module_aliases,
                        import_module_aliases,
                        module_dict_aliases,
                        decode_dict_aliases,
                    ):
                        decode_items[key] = "codecs.decode"
            if not module_items and not decode_items:
                continue
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                if module_items and module_containers.get(target.id) != module_items:
                    module_containers[target.id] = dict(module_items)
                    changed = True
                if decode_items and decode_containers.get(target.id) != decode_items:
                    decode_containers[target.id] = dict(decode_items)
                    changed = True
        if not changed:
            break
    return module_containers, decode_containers


def _container_subscript_lookup(
    node: ast.AST,
    aliases: Dict[str, Dict[Tuple[str, str], str]],
    constants: Dict[str, str],
) -> Optional[str]:
    if not isinstance(node, ast.Subscript) or not isinstance(node.value, ast.Name):
        return None
    key = _container_key_from_literal(node.slice, constants)
    if key is None:
        return None
    return aliases.get(node.value.id, {}).get(key)


def _attribute_alias_lookup(node: ast.AST, aliases: Dict[str, str]) -> Optional[str]:
    if not isinstance(node, ast.Attribute):
        return None
    return aliases.get(_call_name(node))


def _subscript_call_label(node: ast.AST) -> str:
    if not isinstance(node, ast.Subscript) or not isinstance(node.value, ast.Name):
        return "container"
    if isinstance(node.slice, ast.Constant):
        return f"{node.value.id}[{node.slice.value!r}]"
    return f"{node.value.id}[...]"


def _attribute_call_label(node: ast.AST) -> str:
    return _call_name(node) or "attribute"


def _decode_method_container_aliases(
    tree: ast.AST,
    encoded_text_aliases: Dict[str, str],
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
    operator_module_aliases: Set[str],
    attrgetter_aliases: Set[str],
    method_aliases: Dict[str, str],
) -> Dict[str, Dict[Tuple[str, str], str]]:
    aliases: Dict[str, Dict[Tuple[str, str], str]] = {}
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
            if isinstance(value, ast.Name) and value.id in aliases:
                container_items = dict(aliases[value.id])
            else:
                container_items: Dict[Tuple[str, str], str] = {}
                for key, item in _literal_container_items(value, constants):
                    if isinstance(item, ast.Name) and item.id in method_aliases:
                        encoded_as = method_aliases[item.id]
                    else:
                        encoded_as = _decode_method_source_encoding(
                            item,
                            encoded_text_aliases,
                            constants,
                            byte_constructor_aliases,
                            operator_module_aliases,
                            attrgetter_aliases,
                        )
                    if encoded_as is not None:
                        container_items[key] = encoded_as
            if not container_items:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and aliases.get(target.id) != container_items:
                    aliases[target.id] = dict(container_items)
                    changed = True
        if not changed:
            break
    return aliases


def _function_required_arg_count(node: ast.FunctionDef) -> int:
    positional = [*node.args.posonlyargs, *node.args.args]
    positional_defaults = len(node.args.defaults or [])
    required_positional = max(0, len(positional) - positional_defaults)
    required_kwonly = sum(1 for default in node.args.kw_defaults if default is None)
    return required_positional + required_kwonly


def _single_return_value(node: ast.FunctionDef) -> Optional[ast.AST]:
    returns = [item.value for item in ast.walk(node) if isinstance(item, ast.Return) and item.value is not None]
    if len(returns) != 1:
        return None
    return returns[0]


def _decode_method_function_results(
    tree: ast.AST,
    encoded_text_aliases: Dict[str, str],
    constants: Dict[str, str],
    byte_constructor_aliases: Set[str],
    operator_module_aliases: Set[str],
    attrgetter_aliases: Set[str],
    method_aliases: Dict[str, str],
) -> Dict[str, str]:
    functions: Dict[str, str] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or _function_required_arg_count(node) != 0:
            continue
        value = _single_return_value(node)
        if value is None:
            continue
        if isinstance(value, ast.Name) and value.id in method_aliases:
            encoded_as = method_aliases[value.id]
        else:
            encoded_as = _decode_method_source_encoding(
                value,
                encoded_text_aliases,
                constants,
                byte_constructor_aliases,
                operator_module_aliases,
                attrgetter_aliases,
            )
        if encoded_as is not None:
            functions[node.name] = encoded_as
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
            if not isinstance(value, ast.Name) or value.id not in functions:
                continue
            encoded_as = functions[value.id]
            for target in targets:
                if isinstance(target, ast.Name) and functions.get(target.id) != encoded_as:
                    functions[target.id] = encoded_as
                    changed = True
        if not changed:
            break
    return functions


def _operator_methodcaller_result_aliases(
    tree: ast.AST,
    operator_module_aliases: Set[str],
    methodcaller_aliases: Set[str],
    constants: Dict[str, str],
) -> Dict[str, str]:
    aliases: Dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        encoding = _operator_methodcaller_decode_encoding(
            value,
            operator_module_aliases,
            methodcaller_aliases,
            constants,
        )
        if encoding is None:
            continue
        for target in targets:
            if isinstance(target, ast.Name):
                aliases[target.id] = encoding
    return aliases


def _operator_methodcaller_container_aliases(
    tree: ast.AST,
    operator_module_aliases: Set[str],
    methodcaller_aliases: Set[str],
    constants: Dict[str, str],
    methodcaller_result_aliases: Dict[str, str],
) -> Dict[str, Dict[Tuple[str, str], str]]:
    aliases: Dict[str, Dict[Tuple[str, str], str]] = {}
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
            if isinstance(value, ast.Name) and value.id in aliases:
                container_items = dict(aliases[value.id])
            else:
                container_items: Dict[Tuple[str, str], str] = {}
                for key, item in _literal_container_items(value, constants):
                    if isinstance(item, ast.Name) and item.id in methodcaller_result_aliases:
                        encoding = methodcaller_result_aliases[item.id]
                    else:
                        encoding = _operator_methodcaller_decode_encoding(
                            item,
                            operator_module_aliases,
                            methodcaller_aliases,
                            constants,
                        )
                    if encoding is not None:
                        container_items[key] = encoding
            if not container_items:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and aliases.get(target.id) != container_items:
                    aliases[target.id] = dict(container_items)
                    changed = True
        if not changed:
            break
    return aliases


def _operator_methodcaller_function_results(
    tree: ast.AST,
    operator_module_aliases: Set[str],
    methodcaller_aliases: Set[str],
    constants: Dict[str, str],
    methodcaller_result_aliases: Dict[str, str],
) -> Dict[str, str]:
    functions: Dict[str, str] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or _function_required_arg_count(node) != 0:
            continue
        value = _single_return_value(node)
        if value is None:
            continue
        if isinstance(value, ast.Name) and value.id in methodcaller_result_aliases:
            encoding = methodcaller_result_aliases[value.id]
        else:
            encoding = _operator_methodcaller_decode_encoding(
                value,
                operator_module_aliases,
                methodcaller_aliases,
                constants,
            )
        if encoding is not None:
            functions[node.name] = encoding
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
            if not isinstance(value, ast.Name) or value.id not in functions:
                continue
            encoding = functions[value.id]
            for target in targets:
                if isinstance(target, ast.Name) and functions.get(target.id) != encoding:
                    functions[target.id] = encoding
                    changed = True
        if not changed:
            break
    return functions


def _zero_arg_function_call_name(node: ast.AST) -> str:
    if not isinstance(node, ast.Call) or node.args or node.keywords:
        return ""
    return _call_name(node.func)


def _codecs_decode_function_results(
    tree: ast.AST,
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
    decode_containers: Dict[str, Dict[Tuple[str, str], str]],
) -> Set[str]:
    functions: Set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or _function_required_arg_count(node) != 0:
            continue
        value = _single_return_value(node)
        if value is None:
            continue
        if _static_codecs_decode_ref(
            value,
            module_aliases,
            decode_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
            decode_dict_aliases,
        ) or _container_subscript_lookup(value, decode_containers, constants) is not None:
            functions.add(node.name)
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
            if not isinstance(value, ast.Name) or value.id not in functions:
                continue
            for target in targets:
                if isinstance(target, ast.Name) and target.id not in functions:
                    functions.add(target.id)
                    changed = True
        if not changed:
            break
    return functions


def _decoder_factory_function_results(
    tree: ast.AST,
    module_aliases: Set[str],
    getdecoder_aliases: Set[str],
    lookup_aliases: Set[str],
    constants: Dict[str, str],
    decoder_results: Dict[str, str],
    lookup_results: Dict[str, str],
    decoder_containers: Dict[str, Dict[Tuple[str, str], str]],
    lookup_containers: Dict[str, Dict[Tuple[str, str], str]],
    decoder_attributes: Dict[str, str],
    lookup_attributes: Dict[str, str],
) -> Tuple[Dict[str, str], Dict[str, str]]:
    decoder_functions: Dict[str, str] = {}
    lookup_functions: Dict[str, str] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or _function_required_arg_count(node) != 0:
            continue
        value = _single_return_value(node)
        if value is None:
            continue
        if isinstance(value, ast.Name) and value.id in decoder_results:
            decoder_functions[node.name] = decoder_results[value.id]
            continue
        if isinstance(value, ast.Name) and value.id in lookup_results:
            lookup_functions[node.name] = lookup_results[value.id]
            continue
        container_decoder = _container_subscript_lookup(value, decoder_containers, constants)
        if container_decoder is not None:
            decoder_functions[node.name] = container_decoder
            continue
        container_lookup = _container_subscript_lookup(value, lookup_containers, constants)
        if container_lookup is not None:
            lookup_functions[node.name] = container_lookup
            continue
        attribute_decoder = _attribute_alias_lookup(value, decoder_attributes)
        if attribute_decoder is not None:
            decoder_functions[node.name] = attribute_decoder
            continue
        attribute_lookup = _attribute_alias_lookup(value, lookup_attributes)
        if attribute_lookup is not None:
            lookup_functions[node.name] = attribute_lookup
            continue
        kind, encoding = _factory_call_encoding(value, module_aliases, getdecoder_aliases, lookup_aliases, constants)
        if kind == "getdecoder":
            decoder_functions[node.name] = encoding
        elif kind == "lookup":
            lookup_functions[node.name] = encoding
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
            if not isinstance(value, ast.Name):
                continue
            if value.id in decoder_functions:
                encoding = decoder_functions[value.id]
                for target in targets:
                    if isinstance(target, ast.Name) and decoder_functions.get(target.id) != encoding:
                        decoder_functions[target.id] = encoding
                        changed = True
            if value.id in lookup_functions:
                encoding = lookup_functions[value.id]
                for target in targets:
                    if isinstance(target, ast.Name) and lookup_functions.get(target.id) != encoding:
                        lookup_functions[target.id] = encoding
                        changed = True
        if not changed:
            break
    return decoder_functions, lookup_functions


def _is_mismatched_text_redecode(encode_encoding: str, decode_encoding: str) -> bool:
    if not encode_encoding or not decode_encoding:
        return True
    return _normalize_codec_name(encode_encoding) != _normalize_codec_name(decode_encoding)


def _str_constructor_decode_encoding(node: ast.Call, constants: Dict[str, str]) -> str:
    if len(node.args) >= 2:
        return _string_constant(node.args[1], constants)
    keyword = _keyword_node(node, "encoding")
    if keyword is not None:
        return _string_constant(keyword, constants)
    return ""


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
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
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
                if _static_codecs_module_ref(
                    node.func.value,
                    module_aliases,
                    constants,
                    importlib_module_aliases,
                    import_module_aliases,
                    module_dict_aliases,
                ):
                    codec_node = _codecs_decode_encoding_node(node)
                else:
                    codec_node = _method_decode_encoding_node(node)
            elif _getattr_on_module_alias(node.func, module_aliases, "decode", constants):
                codec_node = _codecs_decode_encoding_node(node)
            elif _static_codecs_decode_ref(
                node.func,
                module_aliases,
                decode_aliases,
                constants,
                importlib_module_aliases,
                import_module_aliases,
                module_dict_aliases,
                decode_dict_aliases,
            ):
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


def _is_decode_function_ref(
    node: Optional[ast.AST],
    module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
) -> bool:
    return _static_codecs_decode_ref(
        node,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )


def _is_functools_partial_call(node: ast.AST, module_aliases: Set[str], partial_aliases: Set[str]) -> bool:
    return (
        isinstance(node, ast.Call)
        and (
            _call_name(node.func) in partial_aliases
            or _attribute_on_module_alias(node.func, module_aliases, "partial")
        )
    )


def _partial_decode_encoding(
    node: ast.AST,
    functools_module_aliases: Set[str],
    partial_aliases: Set[str],
    codec_module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
) -> Optional[str]:
    if not _is_functools_partial_call(node, functools_module_aliases, partial_aliases):
        return None
    if not isinstance(node, ast.Call) or not node.args:
        return None
    if not _is_decode_function_ref(
        node.args[0],
        codec_module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    ):
        return None
    if len(node.args) >= 3:
        return _string_constant(node.args[2], constants)
    return _keyword_value(node, constants, "encoding")


def _partial_decode_result_aliases(
    tree: ast.AST,
    functools_module_aliases: Set[str],
    partial_aliases: Set[str],
    codec_module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
) -> Dict[str, str]:
    results: Dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        encoding = _partial_decode_encoding(
            value,
            functools_module_aliases,
            partial_aliases,
            codec_module_aliases,
            decode_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
            decode_dict_aliases,
        )
        if encoding is None:
            continue
        for target in targets:
            if isinstance(target, ast.Name):
                results[target.id] = encoding
    return results


def _partial_decode_function_results(
    tree: ast.AST,
    functools_module_aliases: Set[str],
    partial_aliases: Set[str],
    codec_module_aliases: Set[str],
    decode_aliases: Set[str],
    constants: Dict[str, str],
    importlib_module_aliases: Set[str],
    import_module_aliases: Set[str],
    module_dict_aliases: Dict[str, Set[str]],
    decode_dict_aliases: Dict[str, Set[str]],
    partial_alias_results: Dict[str, str],
) -> Dict[str, str]:
    functions: Dict[str, str] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or _function_required_arg_count(node) != 0:
            continue
        value = _single_return_value(node)
        if value is None:
            continue
        if isinstance(value, ast.Name) and value.id in partial_alias_results:
            encoding = partial_alias_results[value.id]
        else:
            encoding = _partial_decode_encoding(
                value,
                functools_module_aliases,
                partial_aliases,
                codec_module_aliases,
                decode_aliases,
                constants,
                importlib_module_aliases,
                import_module_aliases,
                module_dict_aliases,
                decode_dict_aliases,
            )
        if encoding is not None:
            functions[node.name] = encoding
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
            if not isinstance(value, ast.Name) or value.id not in functions:
                continue
            encoding = functions[value.id]
            for target in targets:
                if isinstance(target, ast.Name) and functions.get(target.id) != encoding:
                    functions[target.id] = encoding
                    changed = True
        if not changed:
            break
    return functions


def unsafe_unicode_decode_calls_from_text(text: str, filename: str = "<generated>") -> List[str]:
    tree = ast.parse(text, filename=filename)
    constants = _constant_string_aliases(tree)
    module_aliases, decode_aliases, escape_decode_aliases, getdecoder_aliases, lookup_aliases = _codec_import_aliases(tree)
    functools_module_aliases, partial_aliases = _functools_import_aliases(tree)
    importlib_module_aliases, import_module_aliases = _importlib_import_aliases(tree)
    operator_module_aliases, methodcaller_aliases, attrgetter_aliases = _operator_import_aliases(tree)
    module_dict_aliases: Dict[str, Set[str]] = {}
    decode_dict_aliases: Dict[str, Set[str]] = {}
    _add_simple_module_assignment_aliases(
        tree,
        module_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
    )
    _add_simple_decode_assignment_aliases(
        tree,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )
    module_dict_aliases, decode_dict_aliases = _literal_dict_codec_aliases(
        tree,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
    )
    _add_simple_module_assignment_aliases(
        tree,
        module_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
    )
    _add_simple_decode_assignment_aliases(
        tree,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )
    codec_module_containers, codec_decode_containers = _codec_literal_container_aliases(
        tree,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )
    _add_simple_factory_assignment_aliases(tree, module_aliases, getdecoder_aliases, lookup_aliases)
    wrappers = _codec_wrapper_functions(
        tree,
        module_aliases,
        decode_aliases,
        getdecoder_aliases,
        lookup_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )
    _add_simple_wrapper_assignment_aliases(tree, wrappers)
    higher_order_wrappers = _higher_order_decode_wrappers(tree, constants)
    _add_simple_wrapper_assignment_aliases(tree, higher_order_wrappers)
    decoder_results, lookup_results = _decoder_factory_result_aliases(
        tree,
        module_aliases,
        getdecoder_aliases,
        lookup_aliases,
        constants,
    )
    decoder_containers, lookup_containers = _decoder_factory_container_aliases(
        tree,
        module_aliases,
        getdecoder_aliases,
        lookup_aliases,
        constants,
        decoder_results,
        lookup_results,
    )
    decoder_attributes, lookup_attributes = _decoder_factory_attribute_aliases(
        tree,
        module_aliases,
        getdecoder_aliases,
        lookup_aliases,
        constants,
        decoder_results,
        lookup_results,
        decoder_containers,
        lookup_containers,
    )
    byte_constructor_aliases = _byte_constructor_aliases(tree, constants)
    encoded_text_bytes = _encoded_text_byte_aliases(tree, constants, byte_constructor_aliases)
    decode_method_results = _decode_method_result_aliases(
        tree,
        encoded_text_bytes,
        constants,
        byte_constructor_aliases,
        operator_module_aliases,
        attrgetter_aliases,
    )
    methodcaller_results = _operator_methodcaller_result_aliases(
        tree,
        operator_module_aliases,
        methodcaller_aliases,
        constants,
    )
    decode_method_containers = _decode_method_container_aliases(
        tree,
        encoded_text_bytes,
        constants,
        byte_constructor_aliases,
        operator_module_aliases,
        attrgetter_aliases,
        decode_method_results,
    )
    methodcaller_containers = _operator_methodcaller_container_aliases(
        tree,
        operator_module_aliases,
        methodcaller_aliases,
        constants,
        methodcaller_results,
    )
    codec_decode_functions = _codecs_decode_function_results(
        tree,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
        codec_decode_containers,
    )
    decoder_functions, lookup_functions = _decoder_factory_function_results(
        tree,
        module_aliases,
        getdecoder_aliases,
        lookup_aliases,
        constants,
        decoder_results,
        lookup_results,
        decoder_containers,
        lookup_containers,
        decoder_attributes,
        lookup_attributes,
    )
    decode_method_functions = _decode_method_function_results(
        tree,
        encoded_text_bytes,
        constants,
        byte_constructor_aliases,
        operator_module_aliases,
        attrgetter_aliases,
        decode_method_results,
    )
    methodcaller_functions = _operator_methodcaller_function_results(
        tree,
        operator_module_aliases,
        methodcaller_aliases,
        constants,
        methodcaller_results,
    )
    partial_results = _partial_decode_result_aliases(
        tree,
        functools_module_aliases,
        partial_aliases,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
    )
    partial_functions = _partial_decode_function_results(
        tree,
        functools_module_aliases,
        partial_aliases,
        module_aliases,
        decode_aliases,
        constants,
        importlib_module_aliases,
        import_module_aliases,
        module_dict_aliases,
        decode_dict_aliases,
        partial_results,
    )
    hits: List[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = _call_name(node.func)
        if name == "str" and node.args:
            encoded_as = _encoded_text_bytes_encoding(node.args[0], encoded_text_bytes, constants, byte_constructor_aliases)
            if encoded_as is not None:
                encoding = _str_constructor_decode_encoding(node, constants)
                if _is_mismatched_text_redecode(encoded_as, encoding):
                    hits.append(f"str({_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                    continue
        methodcaller_encoding = _operator_methodcaller_decode_encoding(
            node.func,
            operator_module_aliases,
            methodcaller_aliases,
            constants,
        )
        if methodcaller_encoding is not None and node.args:
            encoded_as = _encoded_text_bytes_encoding(
                node.args[0],
                encoded_text_bytes,
                constants,
                byte_constructor_aliases,
            )
            if encoded_as is not None and _is_mismatched_text_redecode(encoded_as, methodcaller_encoding):
                hits.append(f"operator.methodcaller(decode {_codec_label(methodcaller_encoding)} after encode {_codec_label(encoded_as)})")
                continue
        if name in methodcaller_results and node.args:
            encoded_as = _encoded_text_bytes_encoding(
                node.args[0],
                encoded_text_bytes,
                constants,
                byte_constructor_aliases,
            )
            encoding = methodcaller_results[name]
            if encoded_as is not None and _is_mismatched_text_redecode(encoded_as, encoding):
                hits.append(f"{name}(decode {_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                continue
        container_methodcaller_encoding = _container_subscript_lookup(node.func, methodcaller_containers, constants)
        if container_methodcaller_encoding is not None and node.args:
            encoded_as = _encoded_text_bytes_encoding(
                node.args[0],
                encoded_text_bytes,
                constants,
                byte_constructor_aliases,
            )
            if encoded_as is not None and _is_mismatched_text_redecode(encoded_as, container_methodcaller_encoding):
                label = _subscript_call_label(node.func)
                hits.append(f"{label}(decode {_codec_label(container_methodcaller_encoding)} after encode {_codec_label(encoded_as)})")
                continue
        function_methodcaller_name = _zero_arg_function_call_name(node.func)
        if function_methodcaller_name in methodcaller_functions and node.args:
            encoded_as = _encoded_text_bytes_encoding(
                node.args[0],
                encoded_text_bytes,
                constants,
                byte_constructor_aliases,
            )
            encoding = methodcaller_functions[function_methodcaller_name]
            if encoded_as is not None and _is_mismatched_text_redecode(encoded_as, encoding):
                hits.append(f"{function_methodcaller_name}()(decode {_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                continue
        bound_encoded_as = _decode_method_source_encoding(
            node.func,
            encoded_text_bytes,
            constants,
            byte_constructor_aliases,
            operator_module_aliases,
            attrgetter_aliases,
        )
        if bound_encoded_as is not None:
            encoding = _method_decode_encoding_or_default(node, constants)
            if _is_mismatched_text_redecode(bound_encoded_as, encoding):
                hits.append(f"bound decode({_codec_label(encoding)} after encode {_codec_label(bound_encoded_as)})")
                continue
        container_bound_encoded_as = _container_subscript_lookup(node.func, decode_method_containers, constants)
        if container_bound_encoded_as is not None:
            encoding = _method_decode_encoding_or_default(node, constants)
            if _is_mismatched_text_redecode(container_bound_encoded_as, encoding):
                label = _subscript_call_label(node.func)
                hits.append(f"{label}({_codec_label(encoding)} after encode {_codec_label(container_bound_encoded_as)})")
                continue
        function_bound_name = _zero_arg_function_call_name(node.func)
        if function_bound_name in decode_method_functions:
            encoding = _method_decode_encoding_or_default(node, constants)
            encoded_as = decode_method_functions[function_bound_name]
            if _is_mismatched_text_redecode(encoded_as, encoding):
                hits.append(f"{function_bound_name}()({_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                continue
        if name in decode_method_results:
            encoding = _method_decode_encoding_or_default(node, constants)
            encoded_as = decode_method_results[name]
            if _is_mismatched_text_redecode(encoded_as, encoding):
                hits.append(f"{name}({_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                continue
        partial_encoding = _partial_decode_encoding(
            node.func,
            functools_module_aliases,
            partial_aliases,
            module_aliases,
            decode_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
            decode_dict_aliases,
        )
        if partial_encoding is not None:
            hits.append(f"functools.partial(codecs.decode)({_codec_label(partial_encoding)})")
            continue
        container_codec_decode = _container_subscript_lookup(node.func, codec_decode_containers, constants)
        if container_codec_decode is not None:
            encoding = _codecs_decode_encoding(node, constants)
            label = _subscript_call_label(node.func)
            hits.append(f"{label}({_codec_label(encoding)})")
            continue
        function_codec_decode_name = _zero_arg_function_call_name(node.func)
        if function_codec_decode_name in codec_decode_functions:
            encoding = _codecs_decode_encoding(node, constants)
            hits.append(f"{function_codec_decode_name}()({_codec_label(encoding)})")
            continue
        if function_codec_decode_name in partial_functions:
            hits.append(f"{function_codec_decode_name}()({_codec_label(partial_functions[function_codec_decode_name])})")
            continue
        kind, encoding = _factory_call_encoding(node.func, module_aliases, getdecoder_aliases, lookup_aliases, constants)
        if kind == "getdecoder":
            hits.append(f"{name or kind}({_codec_label(encoding)})")
            continue
        container_decoder_encoding = _container_subscript_lookup(node.func, decoder_containers, constants)
        if container_decoder_encoding is not None:
            hits.append(f"{_subscript_call_label(node.func)}({_codec_label(container_decoder_encoding)})")
            continue
        attribute_decoder_encoding = _attribute_alias_lookup(node.func, decoder_attributes)
        if attribute_decoder_encoding is not None:
            hits.append(f"{_attribute_call_label(node.func)}({_codec_label(attribute_decoder_encoding)})")
            continue
        function_decoder_name = _zero_arg_function_call_name(node.func)
        if function_decoder_name in decoder_functions:
            hits.append(f"{function_decoder_name}()({_codec_label(decoder_functions[function_decoder_name])})")
            continue
        if _getattr_on_static_codecs_module(
            node.func,
            module_aliases,
            "decode",
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
        ):
            encoding = _codecs_decode_encoding(node, constants)
            hits.append(f"getattr(codecs, decode)({_codec_label(encoding)})")
            continue
        if isinstance(node.func, ast.Attribute) and node.func.attr == "decode":
            base = _call_name(node.func.value)
            if base in lookup_results:
                hits.append(f"{name}({_codec_label(lookup_results[base])})")
                continue
            container_lookup_encoding = _container_subscript_lookup(node.func.value, lookup_containers, constants)
            if container_lookup_encoding is not None:
                hits.append(f"{_subscript_call_label(node.func.value)}.decode({_codec_label(container_lookup_encoding)})")
                continue
            attribute_lookup_encoding = _attribute_alias_lookup(node.func.value, lookup_attributes)
            if attribute_lookup_encoding is not None:
                hits.append(f"{_attribute_call_label(node.func.value)}.decode({_codec_label(attribute_lookup_encoding)})")
                continue
            function_lookup_name = _zero_arg_function_call_name(node.func.value)
            if function_lookup_name in lookup_functions:
                hits.append(f"{function_lookup_name}().decode({_codec_label(lookup_functions[function_lookup_name])})")
                continue
            kind, encoding = _factory_call_encoding(
                node.func.value,
                module_aliases,
                getdecoder_aliases,
                lookup_aliases,
                constants,
            )
            if kind == "lookup":
                hits.append(f"{name}({_codec_label(encoding)})")
                continue
            if _container_subscript_lookup(node.func.value, codec_module_containers, constants) is not None:
                encoding = _codecs_decode_encoding(node, constants)
                hits.append(f"{_subscript_call_label(node.func.value)}.decode({_codec_label(encoding)})")
                continue
            if _static_codecs_module_ref(
                node.func.value,
                module_aliases,
                constants,
                importlib_module_aliases,
                import_module_aliases,
                module_dict_aliases,
            ):
                encoding = _codecs_decode_encoding(node, constants)
                hits.append(f"{name}({_codec_label(encoding)})")
                continue
            encoded_as = _encoded_text_bytes_encoding(node.func.value, encoded_text_bytes, constants, byte_constructor_aliases)
            if encoded_as is not None:
                encoding = _method_decode_encoding_or_default(node, constants)
                if _is_mismatched_text_redecode(encoded_as, encoding):
                    hits.append(f"{name}({_codec_label(encoding)} after encode {_codec_label(encoded_as)})")
                    continue
            encoding = _method_decode_encoding(node, constants)
            if _is_dangerous_unicode_codec(encoding):
                hits.append(f"{name}({encoding})")
            continue
        if _static_codecs_decode_ref(
            node.func,
            module_aliases,
            decode_aliases,
            constants,
            importlib_module_aliases,
            import_module_aliases,
            module_dict_aliases,
            decode_dict_aliases,
        ):
            encoding = _codecs_decode_encoding(node, constants)
            hits.append(f"{name or 'codecs.decode'}({_codec_label(encoding)})")
            continue
        if name in decode_aliases:
            encoding = _codecs_decode_encoding(node, constants)
            hits.append(f"{name}({_codec_label(encoding)})")
            continue
        if name in decoder_results:
            hits.append(f"{name}({_codec_label(decoder_results[name])})")
            continue
        if name in partial_results:
            hits.append(f"{name}({_codec_label(partial_results[name])})")
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
                    if not _is_decode_function_ref(
                        decoder_node,
                        module_aliases,
                        decode_aliases,
                        constants,
                        importlib_module_aliases,
                        import_module_aliases,
                        module_dict_aliases,
                        decode_dict_aliases,
                    ):
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

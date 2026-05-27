"""Extractor verification helpers for pipeline phases."""
from __future__ import annotations

import copy
import hashlib
import json
import os
import shutil


class VerificationError(RuntimeError):
    """Raised when repeated extraction runs cannot be reconciled."""


def _payload(result):
    return result[0] if isinstance(result, tuple) else result


def _signature(result):
    data = _payload(result)
    if isinstance(result, tuple):
        paragraphs = data.get("paragraphs") or []
        tables = data.get("tables") or []
        sections = data.get("sections") or []
        meta = data.get("_meta") or {}
        runs = sum(len(p.get("runs", [])) for p in paragraphs if isinstance(p, dict))
        return ("format", len(paragraphs), len(tables), len(sections), runs, meta.get("paragraphs"))

    sections = data.get("sections") or []
    meta = data.get("_meta") or {}
    return (
        "content",
        len(sections),
        _digest(
            [
                (
                    section.get("heading") or "",
                    section.get("level") or 0,
                    len(section.get("paragraphs") or []),
                    section.get("images") or [],
                )
                for section in sections
            ]
        ),
        len(data.get("references") or []),
        _digest(data.get("references") or []),
        meta.get("images_extracted"),
        len(meta.get("image_extract_failures") or []),
        _digest(
            {
                "images": _collect_image_names(data),
                "missing": meta.get("missing_images") or [],
                "failures": meta.get("image_extract_failures") or [],
                "non_body": meta.get("non_body_images") or [],
            }
        ),
    )


def _signature_diff(sig1, sig2):
    labels = (
        ("kind", "paragraphs", "tables", "sections", "runs", "meta_paragraphs")
        if sig1 and sig1[0] == "format"
        else ("kind", "sections", "section_shape", "references", "reference_digest", "images", "image_failures", "image_digest")
    )
    return [f"{label}: {left} vs {right}" for label, left, right in zip(labels, sig1, sig2) if left != right]


def _canonical(value):
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    except TypeError:
        return str(value)


def _digest(value):
    return hashlib.sha1(_canonical(value).encode("utf-8")).hexdigest()[:12]


def _convergence_key(content):
    clean = copy.deepcopy(content)
    meta = clean.get("_meta")
    if isinstance(meta, dict):
        meta.pop("converged_extraction", None)
    return _canonical(clean)


def _safe_int(value, default=0):
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


def _unique_items(items):
    out = []
    seen = set()
    for item in items or []:
        key = _canonical(item)
        if key in seen:
            continue
        seen.add(key)
        out.append(copy.deepcopy(item))
    return out


def _intersection_items(lists):
    if not lists:
        return []
    key_to_item = {}
    common = None
    for values in lists:
        keys = set()
        for item in values or []:
            key = _canonical(item)
            key_to_item.setdefault(key, item)
            keys.add(key)
        common = keys if common is None else common & keys
    return [copy.deepcopy(key_to_item[key]) for key in sorted(common or [])]


def _section_key(content):
    return tuple(
        (
            str(section.get("heading") or ""),
            _safe_int(section.get("level")),
            str(section.get("role") or ""),
        )
        for section in content.get("sections") or []
    )


def _image_name_from_item(item):
    if not isinstance(item, dict):
        return ""
    if item.get("role") not in ("image", "figure") and not item.get("image"):
        return ""
    return str(item.get("image") or item.get("filename") or item.get("asset") or "")


def _merge_item_key(item):
    image_name = _image_name_from_item(item)
    if image_name:
        return ("image", image_name)
    return ("item", _canonical(item))


def _collect_image_names(content):
    names = []
    for section in content.get("sections") or []:
        names.extend(str(item) for item in (section.get("images") or []) if item)
        for paragraph in section.get("paragraphs") or []:
            name = _image_name_from_item(paragraph)
            if name:
                names.append(name)
    return _unique_items(names)


def _paragraph_score(item):
    if isinstance(item, str):
        return (0, len(item))
    if not isinstance(item, dict):
        return (0, 0)
    role = str(item.get("role") or "")
    return (
        5 if role in ("figure", "image") or item.get("image") else 0,
        4 if role == "formula" or item.get("math") or item.get("latex") or item.get("xml") else 0,
        3 if item.get("table_rows") else 0,
        2 if role == "rich_text" else 0,
        len(str(item.get("text") or "")),
    )


def _section_score(section):
    paragraphs = section.get("paragraphs") or []
    images = section.get("images") or []
    return (
        len(paragraphs),
        len(images),
        sum(1 for item in paragraphs if isinstance(item, dict) and (item.get("image") or item.get("role") in ("figure", "image"))),
        sum(1 for item in paragraphs if isinstance(item, dict) and (item.get("math") or item.get("role") == "formula")),
        sum(1 for item in paragraphs if isinstance(item, dict) and item.get("table_rows")),
        sum(max(_paragraph_score(item)) for item in paragraphs),
    )


def _content_score(content):
    sections = content.get("sections") or []
    meta = content.get("_meta") or {}
    return (
        len(sections),
        sum(len(section.get("paragraphs") or []) for section in sections),
        len(_collect_image_names(content)),
        _safe_int(meta.get("images_extracted")),
        len(content.get("references") or []),
        -len(meta.get("image_extract_failures") or []),
        -len(meta.get("missing_images") or []),
    )


def _merge_paragraph_items(base_items, candidate_sections):
    out = list(copy.deepcopy(base_items or []))
    seen = {_merge_item_key(item) for item in out}

    def insert_recovered_item(candidate_items, item_index, item):
        positions = {_merge_item_key(existing): index for index, existing in enumerate(out)}
        for pos in range(item_index - 1, -1, -1):
            anchor_key = _merge_item_key(candidate_items[pos])
            if anchor_key in positions:
                out.insert(positions[anchor_key] + 1, copy.deepcopy(item))
                return
        for pos in range(item_index + 1, len(candidate_items)):
            anchor_key = _merge_item_key(candidate_items[pos])
            if anchor_key in positions:
                out.insert(positions[anchor_key], copy.deepcopy(item))
                return
        out.append(copy.deepcopy(item))

    for section in candidate_sections:
        candidate_items = section.get("paragraphs") or []
        for item_index, item in enumerate(candidate_items):
            image_name = _image_name_from_item(item)
            should_append = False
            item_key = _merge_item_key(item)
            if image_name and item_key not in seen:
                should_append = True
            elif isinstance(item, dict) and (item.get("table_rows") or item.get("math") or item.get("role") == "formula"):
                should_append = item_key not in seen
            if should_append:
                seen.add(item_key)
                insert_recovered_item(candidate_items, item_index, item)
    return out


def _merge_content_results(results):
    contents = [_payload(result) for result in results]
    section_keys = [_section_key(content) for content in contents]
    if not section_keys or any(key != section_keys[0] for key in section_keys):
        return None, "section structure was not stable"

    merged = copy.deepcopy(max(contents, key=_content_score))
    merged_sections = []
    for index in range(len(section_keys[0])):
        candidates = [content.get("sections", [])[index] for content in contents]
        best = copy.deepcopy(max(candidates, key=_section_score))
        best["images"] = _unique_items(
            image
            for section in candidates
            for image in (section.get("images") or [])
            if image
        )
        best["paragraphs"] = _merge_paragraph_items(best.get("paragraphs") or [], candidates)
        merged_sections.append(best)
    merged["sections"] = merged_sections

    merged["references"] = _unique_items(
        reference
        for content in contents
        for reference in (content.get("references") or [])
    )

    meta = copy.deepcopy(merged.get("_meta") or {})
    metas = [content.get("_meta") or {} for content in contents]
    meta["missing_images"] = _intersection_items([item.get("missing_images") or [] for item in metas])
    meta["image_extract_failures"] = _intersection_items([item.get("image_extract_failures") or [] for item in metas])
    meta["non_body_images"] = _unique_items(
        image
        for item in metas
        for image in (item.get("non_body_images") or [])
    )
    meta["converged_extraction"] = {
        "runs": len(results),
        "strategy": "stable_sections_union",
        "signatures": [repr(_signature(result)) for result in results],
    }
    merged["_meta"] = meta
    return merged, "converged by stable section structure and unioned recoverable content"


def _source_image_dirs(results):
    dirs = []
    for result in results:
        meta = (_payload(result).get("_meta") or {})
        images_dir = meta.get("images_dir")
        if images_dir and os.path.isdir(images_dir):
            dirs.append(images_dir)
    return dirs


def _materialize_content_result(result, all_results, path, output_dir):
    if not output_dir:
        return result
    content = copy.deepcopy(_payload(result))
    base = os.path.splitext(os.path.basename(path))[0]
    final_fig_dir = os.path.join(output_dir, base, "figures")
    shutil.rmtree(final_fig_dir, ignore_errors=True)
    os.makedirs(final_fig_dir, exist_ok=True)

    copied = 0
    for filename in _collect_image_names(content):
        for source_dir in _source_image_dirs(all_results):
            source_path = os.path.join(source_dir, filename)
            if os.path.exists(source_path):
                dest_path = os.path.join(final_fig_dir, filename)
                if not os.path.exists(dest_path):
                    shutil.copy2(source_path, dest_path)
                    copied += 1
                break

    meta = content.setdefault("_meta", {})
    meta["images_dir"] = os.path.abspath(final_fig_dir)
    meta["images_extracted"] = max(copied, len(_collect_image_names(content)))
    return content


def _majority_result(results):
    counts = {}
    first_by_signature = {}
    for result in results:
        sig = _signature(result)
        counts[sig] = counts.get(sig, 0) + 1
        first_by_signature.setdefault(sig, result)
    for sig, count in counts.items():
        if count >= 2:
            return first_by_signature[sig], sig
    return None, None


def _is_content_verification(label, kw):
    return str(label).strip().lower() == "content" and "output_dir" in kw


def double_verify(extractor_fn, path, label, **kw):
    """Run extraction repeatedly, then verify or safely converge structural content."""
    is_content = _is_content_verification(label, kw)
    output_dir = kw.get("output_dir") if is_content else None
    verify_root = os.path.join(output_dir, "_extract_verify_runs") if output_dir else None
    if verify_root:
        shutil.rmtree(verify_root, ignore_errors=True)

    results = []

    def run_once(index):
        run_kw = dict(kw)
        if verify_root:
            run_kw["output_dir"] = os.path.join(verify_root, f"run_{index}")
        return extractor_fn(path, **run_kw)

    try:
        max_runs = 5 if is_content else 3
        last_merge_key = None
        for index in range(1, max_runs + 1):
            results.append(run_once(index))
            if len(results) < 2:
                continue

            majority, sig = _majority_result(results)
            if majority is not None:
                status = "consistent" if len(results) == 2 else f"arbitrated: {sig}"
                print(f"  [OK] {label}: verified ({status})")
                return _materialize_content_result(majority, results, path, output_dir) if is_content else majority

            if is_content and len(results) >= 3:
                merged, reason = _merge_content_results(results)
                if merged is not None:
                    merge_key = _convergence_key(merged)
                    if merge_key == last_merge_key or index == max_runs:
                        merged = _materialize_content_result(merged, results, path, output_dir)
                        print(f"  [OK] {label}: converged ({reason}; runs={len(results)})")
                        return merged
                    last_merge_key = merge_key

        sig1 = _signature(results[0]) if results else None
        sig2 = _signature(results[1]) if len(results) > 1 else None
        checks = _signature_diff(sig1, sig2) if sig1 and sig2 else []
        message = f"{label}: verification failed ({'; '.join(checks)}; signatures={[repr(_signature(item)) for item in results]})"
        print(f"  [ERROR] {message}")
        raise VerificationError(message)
    finally:
        if verify_root:
            shutil.rmtree(verify_root, ignore_errors=True)

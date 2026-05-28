"""Artifact-writing helpers for pipeline JSON and markdown handoffs."""
from __future__ import annotations

import json
import os


def _math_count(item):
    count = len(item.get("math") or [])
    if count:
        return count
    count = 0
    for run in item.get("runs") or []:
        if isinstance(run, dict):
            count += len(run.get("math") or [])
    return count


def _table_shape(rows):
    rows = rows or []
    cols = max((len(row or []) for row in rows), default=0)
    return len(rows), cols


def _paragraph_summary(paragraph):
    if not isinstance(paragraph, dict):
        return str(paragraph)

    role = str(paragraph.get("role") or paragraph.get("type") or "")
    if role == "figure":
        image = paragraph.get("image") or paragraph.get("path") or ""
        caption = paragraph.get("caption") or ""
        text = f"[图片] {image}".strip()
        return f"{text} — {caption}" if caption else text
    if paragraph.get("table_rows") or role == "table":
        rows, cols = _table_shape(paragraph.get("table_rows") or [])
        return f"[表格] {rows}行 x {cols}列"
    if role in {"figure_caption", "table_caption"}:
        return paragraph.get("text") or "[题注]"
    if role == "formula_problem":
        return f"[公式问题] {paragraph.get('text') or paragraph.get('latex') or ''}".strip()
    if role == "formula" or paragraph.get("latex") or paragraph.get("xml"):
        return paragraph.get("text") or paragraph.get("latex") or "[公式]"
    if role == "code":
        return f"[代码] {paragraph.get('code') or paragraph.get('text') or ''}".strip()

    text = paragraph.get("text") or paragraph.get("code") or "[结构化内容]"
    count = _math_count(paragraph)
    if count:
        text += f" (+{count}公式)"
    return text


def write_format_artifacts(fmt, md_text, out_dir):
    fmt_json_path = os.path.join(out_dir, "format.json")
    fmt_md_path = os.path.join(out_dir, "格式提取.md")
    with open(fmt_json_path, "w", encoding="utf-8") as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    with open(fmt_md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    return fmt_json_path, fmt_md_path


def build_content_markdown(content, content_path):
    lines = [f"# 内容提取 — {os.path.basename(content_path)}\n"]
    for sec in content.get("sections", []):
        lines.append(f'## {sec.get("heading", "")}\n')
        inline_images = {
            str(paragraph.get("image") or paragraph.get("filename") or paragraph.get("asset") or "")
            for paragraph in sec.get("paragraphs", [])
            if isinstance(paragraph, dict) and (paragraph.get("role") in {"figure", "image"} or paragraph.get("image"))
        }
        for img in sec.get("images", []):
            if str(img) not in inline_images:
                lines.append(f"- [图片] {img}")
        for paragraph in sec.get("paragraphs", []):
            text = _paragraph_summary(paragraph)
            text = text[:120] + "..." if len(text) > 120 else text
            lines.append(f"- {text}")
        lines.append("")
    if content.get("references"):
        lines.append("## 参考文献\n")
        for ref in content["references"]:
            if isinstance(ref, dict):
                text = ref.get("text") or ref.get("code") or "[结构化内容]"
            else:
                text = str(ref)
            lines.append(f"- {text[:120]}")
    return "\n".join(lines)


def write_content_artifacts(content, out_dir, content_path):
    cnt_json_path = os.path.join(out_dir, "content.json")
    cnt_md_path = os.path.join(out_dir, "内容提取.md")
    with open(cnt_json_path, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    with open(cnt_md_path, "w", encoding="utf-8") as f:
        f.write(build_content_markdown(content, content_path))
    return cnt_json_path, cnt_md_path

"""Artifact-writing helpers for pipeline JSON and markdown handoffs."""
from __future__ import annotations

import json
import os


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
        for img in sec.get("images", []):
            lines.append(f"- [图片] {img}")
        for paragraph in sec.get("paragraphs", []):
            if isinstance(paragraph, dict):
                text = paragraph.get("text", "") or "[公式]"
                if paragraph.get("math"):
                    text += f' (+{len(paragraph["math"])}公式)'
            else:
                text = str(paragraph)
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

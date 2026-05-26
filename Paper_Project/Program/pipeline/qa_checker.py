"""
qa_checker.py - lightweight QA checks for generated Word pipeline outputs.

The checker does not fix files. It writes a structured report that tells the AI
which artifact should be edited in user mode and which core engine owns the same
class of issue in developer mode.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from datetime import datetime
from typing import Any, Dict, Iterable, List

try:
    from privacy import sanitize_value
except ImportError:  # pragma: no cover - standalone fallback
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value

try:
    from formula_semantics import (
        CATEGORY_CONTAMINATED,
        classify_formula_text,
        formula_text_looks_contaminated as semantic_formula_text_looks_contaminated,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .formula_semantics import (
        CATEGORY_CONTAMINATED,
        classify_formula_text,
        formula_text_looks_contaminated as semantic_formula_text_looks_contaminated,
    )


VALID_MODES = {"user", "developer"}

OWNER_BY_CODE = {
    "MISSING_DOCX": "script_generator.py",
    "MISSING_BUILD_SCRIPT": "script_generator.py",
    "MISSING_FORMAT_JSON": "format_extractor.py / md_parser.py",
    "MISSING_CONTENT_JSON": "content_parser.py / md_parser.py",
    "FORMAT_EMPTY": "format_extractor.py / md_parser.py",
    "CONTENT_EMPTY": "content_parser.py / md_parser.py",
    "STYLE_PROFILE_MISSING": "format_extractor.py / script_generator.py",
    "COVER_NOT_EXTRACTED": "format_extractor.py / script_generator.py",
    "TITLE_MISSING": "content_parser.py / md_parser.py",
    "REFERENCES_MISSING": "content_parser.py / md_parser.py",
    "DOCX_XML_UNREADABLE": "script_generator.py",
    "LATEX_ERROR_TEXT": "latex_omath.py / script_generator.py",
    "LATEX_DELIMITER_TEXT": "content_parser.py / script_generator.py / latex_omath.py",
    "FORMULA_PIPE_ARTIFACT": "latex_omath.py",
    "FORMULA_COUNT_MISMATCH": "content_parser.py / md_parser.py / script_generator.py / latex_omath.py",
    "FORMULA_NOT_NATIVE": "content_parser.py / script_generator.py / latex_omath.py",
    "IMAGE_NOT_RENDERED": "content_parser.py / script_generator.py",
    "IMAGE_COUNT_MISMATCH": "content_parser.py / script_generator.py",
    "CONTENT_IMAGE_MISSING": "md_parser.py / content_parser.py",
    "IMAGE_EXTRACT_FAILED": "content_parser.py",
    "NON_BODY_IMAGE_UNSUPPORTED": "content_parser.py / script_generator.py",
    "LOW_RES_IMAGE_FRAGMENT": "content_parser.py / script_generator.py / qa_checker.py",
    "TABLE_COUNT_MISMATCH": "content_parser.py / md_parser.py / script_generator.py",
    "BUILD_MANIFEST_MISSING": "script_generator.py",
    "DOCX_TEXT_TOO_SHORT": "content_parser.py / md_parser.py / script_generator.py",
    "CONTENT_HEADING_MISSING": "content_parser.py / md_parser.py / script_generator.py",
    "CONTENT_TOC_POLLUTION": "content_parser.py / script_generator.py / qa_checker.py",
    "DUPLICATE_FRONT_MATTER_HEADING": "format_extractor.py / script_generator.py / qa_checker.py",
    "UNFILLED_PLACEHOLDER_TEXT": "content_parser.py / script_generator.py / qa_checker.py",
    "FORMULA_NUMBER_CONFLICT": "content_parser.py / script_generator.py / qa_checker.py",
    "FORMULA_TEXT_FRAGMENTED": "content_parser.py / script_generator.py / latex_omath.py",
    "PLACEHOLDER_TEXT_LEFT": "script_generator.py / qa_checker.py",
    "WORD_FIELD_ERROR": "script_generator.py",
    "TOC_MISSING": "script_generator.py",
    "WORKFLOW_MODE_INVALID": "run_pipeline.py",
}


REPAIR_GUIDES = {
    "IMAGE_COUNT_MISMATCH": {
        "title": "图片没有全部进入最终论文",
        "why": "内容中识别到的图片数量大于最终 DOCX 里实际渲染的正文图片数量。",
        "user_action": "先打开 `内容提取.md` 核对图片列表，再打开最终 DOCX 对照缺少哪几张图。若原文图片在文本框、组合图、浮动图层、页眉页脚或嵌入对象中，请在原文里把它们改成普通的“嵌入型图片”后重新运行流水线。",
        "developer_action": "检查 `content_parser.py` 是否完整提取图片关系，以及 `script_generator.py` 是否跳过了 figure/image 段落。",
        "auto_level": "needs_user_confirmation",
    },
    "IMAGE_NOT_RENDERED": {
        "title": "最终论文没有渲染图片",
        "why": "内容中有图片，但最终 DOCX XML 没有检测到 drawing。",
        "user_action": "确认 `figures/` 目录里是否有图片文件；如果没有，回到原内容文件重新插入普通图片后再跑。",
        "developer_action": "检查图片复制路径和 `script_generator.py` 的图片插入分支。",
        "auto_level": "developer_fix",
    },
    "CONTENT_IMAGE_MISSING": {
        "title": "内容里引用的图片文件缺失",
        "why": "Markdown 或 DOCX 内容引用了图片，但对应文件没有找到或没有复制成功。",
        "user_action": "把缺失图片放回原 Markdown 同目录或 Inputs 相关资源目录，保持文件名一致后重新运行流水线。",
        "developer_action": "检查 `md_parser.py` / `content_parser.py` 的图片路径解析。",
        "auto_level": "needs_user_file",
    },
    "IMAGE_EXTRACT_FAILED": {
        "title": "DOCX 图片关系读取失败",
        "why": "源 DOCX 中有图片关系无法读取，通常和损坏图片、特殊嵌入对象或兼容模式有关。",
        "user_action": "在 Word/WPS 打开原内容文件，将报错附近图片另存后重新插入为普通图片，再保存并重跑。",
        "developer_action": "检查 `content_parser.py` 的 relationship 读取和异常记录。",
        "auto_level": "needs_user_file",
    },
    "NON_BODY_IMAGE_UNSUPPORTED": {
        "title": "页眉页脚图片不属于正文内容流",
        "why": "源 DOCX 的页眉或页脚里存在图片。当前流水线以模板控制页眉页脚，不会把内容文件的页眉页脚图片自动迁移到正文。",
        "user_action": "如果这张图是论文正文内容，请在原文中把它移动到正文普通段落或表格外的普通嵌入图片位置后重跑；如果它只是源文件页眉/页脚装饰，可忽略或删除源文件中的该图片。",
        "developer_action": "检查 `content_parser.py` 的 non_body_images 统计以及是否需要为特定产品形态支持页眉/页脚内容迁移。",
        "auto_level": "needs_user_confirmation",
    },
    "LOW_RES_IMAGE_FRAGMENT": {
        "title": "低分辨率图片碎片被当成正文图",
        "why": "内容源里存在极小的图片对象，通常是图表坐标轴文字、公式字形或粘贴对象碎片。如果按正文图片放大，会在最终 DOCX 里变成巨大的模糊文字图片。",
        "user_action": "回到原始未排版内容，替换为完整图表图片或 Word 原生图/公式后重跑；不要把已经碎裂的最终稿再次作为内容输入。",
        "developer_action": "检查 `content_parser.py` 的图片抽取来源、`script_generator.py` 的图片缩放策略，以及是否需要把该类碎片从正文图片流中剔除。",
        "auto_level": "needs_user_file",
    },
    "FORMULA_COUNT_MISMATCH": {
        "title": "公式没有全部转换为 Word 原生公式",
        "why": "内容中识别到的公式数量大于最终 DOCX 里渲染出的公式数量。",
        "user_action": "检查 `内容提取.md` 中的公式。若原文里是普通文本公式，优先改成 Word 公式或 Markdown LaTeX：`$...$` / `$$...$$`，再重跑。",
        "developer_action": "检查 `content_parser.py` 的公式识别、`script_generator.py` 的公式渲染分支，以及 `latex_omath.py` 转换失败情况。",
        "auto_level": "mixed",
    },
    "FORMULA_NOT_NATIVE": {
        "title": "公式不是 Word 原生公式",
        "why": "内容中有公式，但最终 DOCX 没有 OOXML Math。",
        "user_action": "把普通文本公式改成 Word 公式或 Markdown LaTeX 后重新运行。",
        "developer_action": "检查公式 item 是否进入 `body_with_formula()` / OMML 渲染路径。",
        "auto_level": "developer_fix",
    },
    "TABLE_COUNT_MISMATCH": {
        "title": "表格没有全部进入最终论文",
        "why": "内容中识别到的表格数量大于最终 DOCX 里实际渲染的正文表格数量。",
        "user_action": "检查原文是否有嵌套表格、文本框里的表格、截图表格或过于复杂的合并单元格。小白用户最稳妥的做法是把这些表格改成普通 Word 表格后重跑。",
        "developer_action": "检查 `content_parser.py` 的表格抽取和 `script_generator.py` 的 table_rows 渲染。",
        "auto_level": "mixed",
    },
    "REFERENCES_MISSING": {
        "title": "没有识别到参考文献",
        "why": "内容解析结果里没有 references。",
        "user_action": "如果论文需要参考文献，请在原文末尾添加 `参考文献` 或 `References` 标题，并用编号列表写参考文献；如果本来就没有参考文献，可以先忽略这个 warning。",
        "developer_action": "检查 `content_parser.py` / `md_parser.py` 的参考文献标题识别。",
        "auto_level": "optional_user_input",
    },
    "CONTENT_HEADING_MISSING": {
        "title": "部分标题在最终论文中未按原文出现",
        "why": "QA 按原始标题文本搜索最终 DOCX 时没有找到完全匹配。",
        "user_action": "打开最终 DOCX 看这些标题是否被模板翻译或改写。如果只是 `致谢` 变成 `ACKNOWLEDGEMENTS`，通常可忽略；如果确实缺标题，请回到原文补齐标题后重跑。",
        "developer_action": "检查标题本地化、front matter/back matter 映射和 `script_generator.py` 的标题输出。",
        "auto_level": "manual_review",
    },
    "DUPLICATE_FRONT_MATTER_HEADING": {
        "title": "前置章节标题重复",
        "why": "最终 DOCX 里同一类前置标题（如 摘要 / ABSTRACT）出现了多次，通常是模板示例页或封面抽取残留又和正文生成结果叠加。",
        "user_action": "打开最终 DOCX 对照重复标题所在页；普通用户可先换一个干净模板或删除模板中的示例摘要页后重跑。",
        "developer_action": "检查 `format_extractor.py` 的封面/声明停止边界，以及 `script_generator.py` 是否重复渲染 front matter。",
        "auto_level": "developer_fix",
    },
    "CONTENT_TOC_POLLUTION": {
        "title": "源文档目录被当成正文",
        "why": "内容提取结果里出现了目录页残留，正文长句、公式或数值碎片被识别成章节标题。",
        "user_action": "如果源 DOCX 本身已经排版过，请删除源文件里的目录页后重跑；也可以直接使用未排版正文作为输入。",
        "developer_action": "检查 `content_parser.py` 的源目录跳过逻辑和标题识别规则。",
        "auto_level": "developer_fix",
    },
    "UNFILLED_PLACEHOLDER_TEXT": {
        "title": "未填写的模板占位符",
        "why": "输入或最终 DOCX 中仍有 `[报名序号]`、待填写、TODO、XXXX 等占位文本。",
        "user_action": "在源文档或模板中补齐对应字段，或删除不需要的占位行后重跑。",
        "developer_action": "检查 `content_parser.py` 的占位符识别，以及 `script_generator.py` 的封面字段替换/过滤。",
        "auto_level": "needs_user_input",
    },
    "FORMULA_NUMBER_CONFLICT": {
        "title": "公式编号冲突",
        "why": "公式文本里出现 `(1)(1)`、`(6)(3)` 这类重复编号，通常是源文档已有编号又被生成器重新编号。",
        "user_action": "源文档里的公式编号可以保留，但不要把已经由旧流水线生成过且编号异常的 DOCX 当成最终交付件。",
        "developer_action": "检查 `content_parser.py` 是否清理源公式编号，以及 `script_generator.py` 是否只编号一次。",
        "auto_level": "developer_fix",
    },
    "FORMULA_TEXT_FRAGMENTED": {
        "title": "公式文本碎裂",
        "why": "检测到公式变量、分母、求和上下限等被拆成多个普通短段落，或公式对象中混入了正文叙述，视觉上会像散落/错位的公式碎片。",
        "user_action": "优先使用 Word 原生公式或 Markdown LaTeX 输入；如果源 DOCX 已经碎裂，请回到原始未排版内容重跑。",
        "developer_action": "检查 `content_parser.py` 的公式段合并/OMML 提取，以及 `latex_omath.py` 的转换能力。",
        "auto_level": "mixed",
    },
    "DOCX_TEXT_TOO_SHORT": {
        "title": "最终论文正文明显变少",
        "why": "最终 DOCX 的文本量显著少于内容提取结果。",
        "user_action": "不要直接交付当前 DOCX。先检查 `内容提取.md` 是否完整，再看最终 DOCX 从哪一章开始缺失。",
        "developer_action": "检查内容分段、生成循环和异常提前退出。",
        "auto_level": "developer_fix",
    },
    "LATEX_ERROR_TEXT": {
        "title": "公式转换报错文本残留",
        "why": "最终 DOCX 中还残留 `[LaTeX error ...]`。",
        "user_action": "把对应公式改成更简单、标准的 LaTeX 写法后重跑。",
        "developer_action": "补充 `latex_omath.py` 对该 LaTeX 语法的支持。",
        "auto_level": "mixed",
    },
    "LATEX_DELIMITER_TEXT": {
        "title": "LaTeX 公式分隔符残留",
        "why": "最终 DOCX 里仍存在 `$...$` 或 `$$...$$` 原始公式文本，说明内容解析或公式渲染漏掉了该公式。",
        "user_action": "检查 `内容提取.md` 中对应段落；普通用户可把该公式改成 Word 原生公式或标准 Markdown LaTeX 后重跑。",
        "developer_action": "检查 `content_parser.py` 是否把该段识别为 formula item，以及 `script_generator.py` / `latex_omath.py` 是否完成渲染。",
        "auto_level": "developer_fix",
    },
    "PLACEHOLDER_TEXT_LEFT": {
        "title": "模板占位符残留",
        "why": "最终 DOCX 里可能还保留了 TODO、XXXX、待填写等模板文字。",
        "user_action": "打开最终 DOCX 搜索 TODO、XXXX、待填写，确认是否需要删除或替换。",
        "developer_action": "扩展 `script_generator.py` 的模板占位符过滤规则。",
        "auto_level": "manual_review",
    },
}


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_manifest_counts(out_dir: str) -> Dict[str, Any]:
    path = os.path.join(out_dir, "build_manifest.json")
    if not os.path.exists(path):
        return {}
    data = _load_json(path)
    return data.get("counts") or {}


def _read_docx_xml(docx_path: str) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        return zf.read("word/document.xml").decode("utf-8", errors="replace")


def _xml_plain_text(xml: str) -> str:
    texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, flags=re.S)
    return "".join(re.sub(r"<[^>]+>", "", t) for t in texts)


def _xml_paragraph_texts(xml: str) -> List[str]:
    paragraphs = re.findall(r"<w:p\b[^>]*>.*?</w:p>", xml, flags=re.S)
    return [_xml_plain_text(p).strip() for p in paragraphs]


def _front_matter_heading_role(text: str) -> str:
    compact = re.sub(r"[\s\u3000:：]+", "", str(text or "")).upper()
    if compact in {"摘要", "中文摘要"}:
        return "cn_abstract"
    if compact == "ABSTRACT":
        return "en_abstract"
    if compact in {"关键词", "关键字"}:
        return "cn_keywords"
    if compact in {"KEYWORDS", "KEYWORD"}:
        return "en_keywords"
    return ""


def _content_front_roles(content: Dict[str, Any]) -> set[str]:
    roles: set[str] = set()
    for sec in content.get("sections") or []:
        role = str(sec.get("role") or "")
        if role in {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}:
            roles.add(role)
            continue
        heading_role = _front_matter_heading_role(str(sec.get("heading") or ""))
        if heading_role:
            roles.add(heading_role)
    return roles


def _duplicate_front_matter_headings(content: Dict[str, Any], xml: str) -> List[str]:
    expected_roles = _content_front_roles(content)
    if not expected_roles:
        return []
    role_labels = {
        "cn_abstract": "摘要",
        "en_abstract": "ABSTRACT",
        "cn_keywords": "关键词",
        "en_keywords": "KEY WORDS",
    }
    counts: Dict[str, int] = {}
    samples: Dict[str, List[int]] = {}
    for idx, text in enumerate(_xml_paragraph_texts(xml), 1):
        role = _front_matter_heading_role(text)
        if not role or role not in expected_roles:
            continue
        counts[role] = counts.get(role, 0) + 1
        samples.setdefault(role, []).append(idx)
    return [
        f"{role_labels.get(role, role)} x{count} at paragraphs {','.join(map(str, samples.get(role, [])[:6]))}"
        for role, count in sorted(counts.items())
        if count > 1
    ]


_PLACEHOLDER_RE = re.compile(
    r"(\[[^\]\n]*(?:报名|序号|姓名|学号|学院|专业|班级|题目|指导|教师|日期|编码|待填|请输入|XX|XXX)[^\]\n]*\])"
    r"|(\{\{[^}]+\}\}|TODO|FIXME|待填写|待补全|XXXX)",
    re.I,
)


def _placeholder_samples_from_texts(texts: Iterable[str], limit: int = 8) -> List[str]:
    out: List[str] = []
    for text in texts:
        t = str(text or "").strip()
        if t and _PLACEHOLDER_RE.search(t):
            out.append(t[:120])
            if len(out) >= limit:
                break
    return out


def _heading_looks_like_body_or_formula(text: str) -> bool:
    t = str(text or "").strip()
    if not t:
        return False
    if len(t) > 90:
        return True
    if re.search(r"(?:MWh|MW|Etotal|Esell|Ebuy|ERE|Pbuy|Psell|PRE|rself|rgreen|rup|max\s*\(|∑)", t):
        return True
    if re.search(r"\d+\s*\.\s*\d+$", t):
        return True
    if re.search(r"[=+\-*/×·]\s*\d", t) and len(t) > 20:
        return True
    if re.search(r"[。；;]\s*$", t) and len(t) > 35:
        return True
    return False


def _content_toc_pollution_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    sections = content.get("sections") or []
    samples: List[str] = []
    seen_empty: Dict[str, int] = {}
    for idx, sec in enumerate(sections):
        heading = str(sec.get("heading") or "").strip()
        role = str(sec.get("role") or "")
        paras = sec.get("paragraphs") or []
        norm = re.sub(r"\s+", "", heading)
        if _heading_looks_like_body_or_formula(heading):
            samples.append(f"heading#{idx + 1}: {heading[:120]}")
        if role == "heading" and heading and not paras:
            seen_empty.setdefault(norm, idx + 1)
        elif norm in seen_empty:
            samples.append(f"empty duplicate before body: {heading[:80]}")
        if len(samples) >= limit:
            break
    return samples


def _formula_number_conflict_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    out: List[str] = []
    for item in _iter_paragraph_items(content):
        if not isinstance(item, dict):
            continue
        if item.get("role") == "rich_text":
            continue
        if not (item.get("role") == "formula" or item.get("latex") or item.get("xml") or item.get("math")):
            continue
        text = str(item.get("text") or "")
        labels = re.findall(r"[\(\uff08]\s*\d+(?:\s*[-.]\s*\d+)?\s*[\)\uff09]", text)
        if len(labels) >= 2:
            out.append(text[:160])
            if len(out) >= limit:
                break
    return out


def _formula_text_looks_contaminated(text: str) -> bool:
    return semantic_formula_text_looks_contaminated(text)


def _fragmented_formula_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    out: List[str] = []
    token_re = re.compile(r"^(?:E|RE|total|rself|rgreen|rup|max|t=\d+|\d+\s+\d+|\d+(?:\.\d+)?|[A-Za-z]{1,8})$")
    for sec in content.get("sections") or []:
        streak: List[str] = []
        for item in sec.get("paragraphs") or []:
            if isinstance(item, dict) and item.get("role") == "rich_text":
                streak = []
                continue
            if isinstance(item, dict) and item.get("role") == "formula_problem":
                formula_text = str(item.get("text") or "")
                semantic = item.get("formula_semantics") or classify_formula_text(formula_text).to_dict()
                out.append(
                    f"{sec.get('heading')}: formula semantic problem "
                    f"`{semantic.get('category', CATEGORY_CONTAMINATED)}` `{formula_text[:100]}`"
                )
                if len(out) >= limit:
                    return out
                streak = []
                continue
            if isinstance(item, dict) and (item.get("role") == "formula" or item.get("math") or item.get("latex") or item.get("xml")):
                formula_text = str(item.get("text") or "")
                semantic = item.get("formula_semantics") or classify_formula_text(formula_text).to_dict()
                if semantic.get("category") == CATEGORY_CONTAMINATED or _formula_text_looks_contaminated(formula_text):
                    out.append(f"{sec.get('heading')}: contaminated formula text `{formula_text[:100]}`")
                    if len(out) >= limit:
                        return out
                    streak = []
                    continue
            text = item if isinstance(item, str) else (item.get("text") if isinstance(item, dict) else "")
            t = str(text or "").strip()
            if token_re.match(t):
                streak.append(t)
                if len(streak) >= 3:
                    out.append(f"{sec.get('heading')}: {' / '.join(streak[-6:])}")
                    if len(out) >= limit:
                        return out
            else:
                streak = []
    return out


def _iter_paragraph_items(content: Dict[str, Any]) -> Iterable[Any]:
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            yield item


def _count_content_formulas(content: Dict[str, Any]) -> int:
    total = 0
    for item in _iter_paragraph_items(content):
        if not isinstance(item, dict):
            continue
        math_items = item.get("math") or []
        if math_items:
            total += len(math_items)
        elif item.get("role") == "formula" or item.get("latex"):
            total += 1
    return total


def _count_content_tables(content: Dict[str, Any]) -> int:
    total = 0
    saw_table_rows = False
    for item in _iter_paragraph_items(content):
        if isinstance(item, dict) and item.get("table_rows"):
            saw_table_rows = True
            if item.get("role") != "code":
                total += 1
    if saw_table_rows:
        return total
    return int((content.get("_meta") or {}).get("tables_count") or 0)


def _count_content_images(content: Dict[str, Any]) -> int:
    inline_total = 0
    inline_names: List[str] = []
    section_total = 0
    section_names: List[str] = []
    for sec in content.get("sections") or []:
        section_images = [str(x or "") for x in (sec.get("images") or [])]
        section_total += len(section_images)
        section_names.extend(section_images)
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in ("image", "figure") and (item.get("image") or item.get("filename") or item.get("asset")):
                inline_total += 1
                inline_names.append(str(item.get("image") or item.get("filename") or item.get("asset") or ""))
    if inline_total:
        extra_section_only = [name for name in section_names if name and name not in inline_names]
        return inline_total + len(extra_section_only)
    if section_total:
        return section_total
    return int((content.get("_meta") or {}).get("images_extracted") or 0)


def _iter_content_image_refs(content: Dict[str, Any]) -> Iterable[Dict[str, str]]:
    seen: set[str] = set()
    for sec in content.get("sections") or []:
        heading = str(sec.get("heading") or "")
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in ("image", "figure") and (item.get("image") or item.get("filename") or item.get("asset")):
                name = str(item.get("image") or item.get("filename") or item.get("asset") or "")
                if name and name not in seen:
                    seen.add(name)
                    yield {"name": name, "heading": heading, "caption": str(item.get("caption") or "")}
        for name in sec.get("images") or []:
            name = str(name or "")
            if name and name not in seen:
                seen.add(name)
                yield {"name": name, "heading": heading, "caption": ""}


def _looks_like_low_res_text_fragment(width: int, height: int, context: str = "") -> bool:
    if width <= 0 or height <= 0:
        return False
    context = str(context or "")
    context_hint = bool(re.search(
        r"(fragment|shard|broken|label|formula|equation|text|ocr|"
        r"碎片|残片|公式|方程|标签|标注|文字|截图|图表)",
        context,
        re.I,
    ))
    wide_ratio = width / max(height, 1)
    tall_ratio = height / max(width, 1)
    if width < 160 and height < 45 and wide_ratio >= 2.4:
        return True
    if height < 80 and wide_ratio >= 4.0:
        return True
    if width < 120 and tall_ratio >= 4.0:
        return True
    if context_hint and width < 240 and height < 120 and max(wide_ratio, tall_ratio) >= 1.8:
        return True
    return False


def _low_res_image_fragment_samples(content: Dict[str, Any], out_dir: str, limit: int = 8) -> List[str]:
    try:
        from PIL import Image
    except Exception:
        return []
    meta = content.get("_meta") or {}
    images_dir = str(meta.get("images_dir") or "")
    candidates = []
    if images_dir:
        candidates.append(images_dir)
    candidates.append(os.path.join(out_dir, "figures"))
    samples: List[str] = []
    for ref in _iter_content_image_refs(content):
        name = ref["name"]
        path = next((os.path.join(base, name) for base in candidates if base and os.path.exists(os.path.join(base, name))), "")
        if not path:
            continue
        try:
            with Image.open(path) as im:
                width, height = im.size
        except Exception:
            continue
        context = ref.get("caption") or ref.get("heading") or ""
        if _looks_like_low_res_text_fragment(width, height, context):
            detail = f"{name} {width}x{height}"
            if context:
                detail += f" ({context[:60]})"
            samples.append(detail)
            if len(samples) >= limit:
                break
    return samples


def _content_text_chars(content: Dict[str, Any]) -> int:
    parts: List[str] = []
    title_info = content.get("title_info") or {}
    parts.extend(str(v or "") for v in title_info.values())
    for sec in content.get("sections") or []:
        parts.append(str(sec.get("heading") or ""))
        for item in sec.get("paragraphs") or []:
            if isinstance(item, dict):
                parts.append(str(item.get("text") or item.get("code") or ""))
                for row in item.get("table_rows") or []:
                    parts.extend(str(cell or "") for cell in row)
            else:
                parts.append(str(item or ""))
    for ref in content.get("references") or []:
        if isinstance(ref, dict):
            parts.append(str(ref.get("text") or ref.get("code") or ""))
        else:
            parts.append(str(ref or ""))
    return sum(len(p.strip()) for p in parts if p and p.strip())


def _compact_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


def _missing_heading_samples(content: Dict[str, Any], plain_text: str, limit: int = 6) -> List[str]:
    compact = _compact_text(plain_text)
    missing: List[str] = []
    front_roles = {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}
    for sec in content.get("sections") or []:
        if sec.get("role") in front_roles:
            continue
        heading = str(sec.get("heading") or "").strip()
        if not heading or heading == "正文" or len(heading) > 80:
            continue
        sample = _compact_text(heading)
        if sample and sample not in compact:
            missing.append(heading)
        if len(missing) >= limit:
            break
    return missing


def _issue(code: str, severity: str, message: str, mode: str, detail: str = "") -> Dict[str, Any]:
    owner_dev = OWNER_BY_CODE.get(code, "script_generator.py")
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "detail": detail,
        "owner_user": "Outputs/<run>/build_generated.py",
        "owner_developer": owner_dev,
        "active_owner": "Outputs/<run>/build_generated.py" if mode == "user" else owner_dev,
    }


def _safe_rel(path: str, root: str | None = None) -> str:
    try:
        base = os.path.abspath(root or os.getcwd())
        return os.path.relpath(os.path.abspath(path), base).replace("\\", "/")
    except Exception:
        return path.replace("\\", "/")


def _workflow_commands(out_dir: str, mode: str) -> Dict[str, str]:
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    commands = {
        "rebuild_current_docx": f"python {os.path.join(out_dir, 'build_generated.py')}",
        "rerun_current_pipeline": "",
    }
    try:
        workflow = _load_json(workflow_path)
        template = workflow.get("template")
        content = workflow.get("content")
        if template and content:
            commands["rerun_current_pipeline"] = (
                f"python run_pipeline.py --mode {mode} --template {template} --content {content}"
            )
    except Exception:
        pass
    return commands


def _parse_count_detail(detail: str) -> Dict[str, int]:
    found: Dict[str, int] = {}
    for key, value in re.findall(r"(content|rendered|docx)=([0-9]+)", str(detail or "")):
        try:
            found[key] = int(value)
        except ValueError:
            pass
    return found


def _repair_step(issue: Dict[str, Any], counts: Dict[str, Any], mode: str) -> Dict[str, Any]:
    code = str(issue.get("code") or "")
    guide = REPAIR_GUIDES.get(code, {})
    count_detail = _parse_count_detail(str(issue.get("detail") or ""))
    owner = issue.get("active_owner") or ("Outputs/<run>/build_generated.py" if mode == "user" else OWNER_BY_CODE.get(code, "script_generator.py"))
    return {
        "code": code,
        "severity": issue.get("severity"),
        "title": guide.get("title") or code,
        "why": guide.get("why") or str(issue.get("message") or ""),
        "detail": issue.get("detail") or "",
        "counts": count_detail,
        "auto_level": guide.get("auto_level") or "manual_review",
        "target": owner,
        "user_action": guide.get("user_action") or "打开 `qa_report.md` 和最终 DOCX，对照问题详情核查。",
        "developer_action": guide.get("developer_action") or f"检查 `{OWNER_BY_CODE.get(code, 'script_generator.py')}`。",
    }


def build_repair_plan(report: Dict[str, Any], out_dir: str) -> Dict[str, Any]:
    issues = report.get("issues") or []
    counts = report.get("counts") or {}
    mode = str(report.get("mode") or "user")
    errors = [i for i in issues if i.get("severity") == "error"]
    warnings = [i for i in issues if i.get("severity") == "warning"]
    ordered_issues = errors + warnings + [i for i in issues if i not in errors and i not in warnings]
    steps = [_repair_step(item, counts, mode) for item in ordered_issues]
    commands = _workflow_commands(out_dir, mode)
    summary = (
        "QA 已通过，仍建议用 WPS/Word 做最终视觉核对。"
        if not errors else
        f"QA 发现 {len(errors)} 个阻断错误和 {len(warnings)} 个警告。最终 DOCX 已保留，但交付前需要按修复计划处理。"
    )
    user_prompt_lines = [
        "请继续修复本次 Word 论文流水线输出。",
        f"输出目录：{_safe_rel(out_dir)}",
        "先阅读 `qa_repair_plan.md`、`qa_report.md`、`内容提取.md`、`build_manifest.json`。",
        "目标：优先处理 error，再处理 warning；修复后重新生成最终 DOCX 并重新运行 QA。",
    ]
    if mode == "user":
        user_prompt_lines.append("当前是 user 模式，优先只修改本次输出目录里的 `build_generated.py` 或指导用户修正输入文件。")
    else:
        user_prompt_lines.append("当前是 developer 模式，可修改 `Paper_Project/Program/pipeline/` 下的核心引擎脚本并重跑完整流水线。")
    for idx, step in enumerate(steps[:5], 1):
        user_prompt_lines.append(f"{idx}. {step['code']}: {step['user_action']}")
    return {
        "schema_version": 1,
        "passed": bool(report.get("passed")),
        "summary": summary,
        "mode": mode,
        "blocking_errors": len(errors),
        "warnings": len(warnings),
        "output_dir": _safe_rel(out_dir),
        "open_first": [
            "qa_repair_plan.md",
            "qa_report.md",
            "内容提取.md",
            "build_manifest.json",
            "最终论文.docx",
        ],
        "commands": commands,
        "steps": steps,
        "copy_to_ai_prompt": "\n".join(user_prompt_lines),
    }


def check_output(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    mode = mode if mode in VALID_MODES else "user"
    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}

    def add(code: str, severity: str, message: str, detail: str = "") -> None:
        issues.append(_issue(code, severity, message, mode, detail))

    paths = {
        "docx": os.path.join(out_dir, output_docx_name),
        "build": os.path.join(out_dir, "build_generated.py"),
        "format": os.path.join(out_dir, "format.json"),
        "content": os.path.join(out_dir, "content.json"),
        "workflow": os.path.join(out_dir, "workflow_mode.json"),
        "manifest": os.path.join(out_dir, "build_manifest.json"),
    }

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

    fmt: Dict[str, Any] = {}
    content: Dict[str, Any] = {}
    if os.path.exists(paths["format"]):
        try:
            fmt = _load_json(paths["format"])
            counts["format_paragraphs"] = len(fmt.get("paragraphs") or [])
            counts["format_tables"] = len(fmt.get("tables") or [])
            counts["format_sections"] = len(fmt.get("sections") or [])
            counts["cover_elements"] = len(fmt.get("cover") or [])
            counts["style_profiles"] = len(fmt.get("style_profiles") or {})
            source = str((fmt.get("_meta") or {}).get("source") or "")
            is_md_format = source.lower().endswith(".md")
            if not fmt.get("sections"):
                add("FORMAT_EMPTY", "error", "格式提取结果没有 section。")
            if not fmt.get("paragraphs"):
                add("FORMAT_EMPTY", "warning", "格式提取结果没有 paragraph，可能大量使用默认格式。")
            expected = {"body", "h1", "h2", "h3"}
            missing = sorted(expected - set((fmt.get("style_profiles") or {}).keys()))
            if missing and not is_md_format:
                add("STYLE_PROFILE_MISSING", "warning", "关键样式 profile 不完整。", ", ".join(missing))
            if not fmt.get("cover") and not is_md_format:
                add("COVER_NOT_EXTRACTED", "warning", "没有提取到封面结构；纯 MD 或无封面模板时可以忽略。")
        except Exception as exc:
            add("MISSING_FORMAT_JSON", "error", "format.json 无法读取。", str(exc))

    if os.path.exists(paths["content"]):
        try:
            content = _load_json(paths["content"])
            counts["content_sections"] = len(content.get("sections") or [])
            counts["references"] = len(content.get("references") or [])
            counts["content_images"] = _count_content_images(content)
            counts["content_formulas"] = _count_content_formulas(content)
            counts["content_tables"] = _count_content_tables(content)
            counts["content_text_chars"] = _content_text_chars(content)
            meta = content.get("_meta") or {}
            missing_images = meta.get("missing_images") or []
            image_failures = meta.get("image_extract_failures") or []
            non_body_images = meta.get("non_body_images") or []
            counts["content_missing_images"] = len(missing_images)
            counts["content_image_extract_failures"] = len(image_failures)
            counts["content_non_body_images"] = len(non_body_images)
            if missing_images:
                add(
                    "CONTENT_IMAGE_MISSING",
                    "error",
                    "内容中存在未能解析或复制的图片引用，最终文档会丢图。",
                    json.dumps(sanitize_value(missing_images[:5], os.getcwd()), ensure_ascii=False),
                )
            if image_failures:
                add(
                    "IMAGE_EXTRACT_FAILED",
                    "error",
                    "DOCX 图片关系读取失败，最终文档可能缺图。",
                    json.dumps(sanitize_value(image_failures[:5], os.getcwd()), ensure_ascii=False),
                )
            if non_body_images:
                add(
                    "NON_BODY_IMAGE_UNSUPPORTED",
                    "error",
                    "源 DOCX 页眉/页脚中存在图片，当前不会作为正文图片渲染。",
                    json.dumps(sanitize_value(non_body_images[:5], os.getcwd()), ensure_ascii=False),
                )
            low_res_fragments = _low_res_image_fragment_samples(content, out_dir)
            if low_res_fragments:
                add(
                    "LOW_RES_IMAGE_FRAGMENT",
                    "error",
                    "内容中存在疑似图表/公式碎片的低分辨率图片，放大后会形成巨大模糊文字图。",
                    " / ".join(low_res_fragments),
                )
            if not content.get("sections"):
                add("CONTENT_EMPTY", "error", "内容提取结果没有正文 section。")
            toc_pollution = _content_toc_pollution_samples(content)
            if toc_pollution:
                add(
                    "CONTENT_TOC_POLLUTION",
                    "error",
                    "内容提取结果疑似混入源文档目录页，正文长句/公式碎片被当成标题。",
                    " / ".join(toc_pollution),
                )
            source_placeholders = []
            meta = content.get("_meta") or {}
            for item in meta.get("source_placeholders") or []:
                if isinstance(item, dict):
                    source_placeholders.append(str(item.get("text") or ""))
                else:
                    source_placeholders.append(str(item))
            content_texts = []
            title_info = content.get("title_info") or {}
            content_texts.extend(str(v or "") for v in title_info.values())
            for sec in content.get("sections") or []:
                content_texts.append(str(sec.get("heading") or ""))
                for item in sec.get("paragraphs") or []:
                    if isinstance(item, str):
                        content_texts.append(item)
                    elif isinstance(item, dict):
                        content_texts.append(str(item.get("text") or item.get("code") or ""))
            placeholder_samples = source_placeholders[:8] or _placeholder_samples_from_texts(content_texts)
            if placeholder_samples:
                add(
                    "UNFILLED_PLACEHOLDER_TEXT",
                    "error",
                    "输入内容或提取结果中存在未填写的模板占位符。",
                    " / ".join(x[:120] for x in placeholder_samples[:8] if x),
                )
            formula_conflicts = _formula_number_conflict_samples(content)
            if formula_conflicts:
                add(
                    "FORMULA_NUMBER_CONFLICT",
                    "error",
                    "内容公式中存在重复或冲突编号，可能导致最终公式编号叠加。",
                    " / ".join(formula_conflicts),
                )
            fragmented_formulas = _fragmented_formula_samples(content)
            if fragmented_formulas:
                add(
                    "FORMULA_TEXT_FRAGMENTED",
                    "error",
                    "内容中存在疑似碎裂公式文本，需要人工核对或改用原生公式/LaTeX。",
                    " / ".join(fragmented_formulas),
                )
            title_info = content.get("title_info") or {}
            if not any(str(v or "").strip() for v in title_info.values()):
                add("TITLE_MISSING", "warning", "未识别到论文标题信息，封面和标题页可能需要人工核对。")
            if not content.get("references"):
                add("REFERENCES_MISSING", "warning", "未识别到参考文献。")
        except Exception as exc:
            add("MISSING_CONTENT_JSON", "error", "content.json 无法读取。", str(exc))

    if os.path.exists(paths["docx"]):
        try:
            xml = _read_docx_xml(paths["docx"])
            plain = _xml_plain_text(xml)
            counts["docx_oMathPara"] = len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMathPara\b", xml))
            counts["docx_oMath"] = len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMath\b", xml))
            counts["docx_drawings"] = len(re.findall(r"<wp:(?:inline|anchor)\b", xml))
            counts["docx_tables"] = len(re.findall(r"<w:tbl\b", xml))
            counts["docx_text_chars"] = len(plain)

            if "[LaTeX error" in plain or "[LaTeX error" in xml:
                add("LATEX_ERROR_TEXT", "error", "最终文档中仍包含 LaTeX 转换错误占位。")
            if re.search(r"\$\$[^$]{2,}\$\$|\$[^\n$]{2,}\$", plain):
                add("LATEX_DELIMITER_TEXT", "error", "最终文档中仍残留 LaTeX 公式分隔符，可能有公式未转换。")
            if "M|b|p|s" in plain or "M|b|p|s" in xml:
                add("FORMULA_PIPE_ARTIFACT", "error", "公式出现 run 分隔伪影，例如 M|b|p|s。")
            rendered_formulas = int(manifest_counts["content_formulas_rendered"]) if "content_formulas_rendered" in manifest_counts else int(counts.get("docx_oMath", 0) or 0)
            rendered_images = int(manifest_counts["content_images_rendered"]) if "content_images_rendered" in manifest_counts else int(counts.get("docx_drawings", 0) or 0)
            rendered_tables = int(manifest_counts["content_tables_rendered"]) if "content_tables_rendered" in manifest_counts else int(counts.get("docx_tables", 0) or 0)
            if counts.get("content_formulas", 0) and counts.get("docx_oMath", 0) == 0:
                add("FORMULA_NOT_NATIVE", "error", "内容中有公式，但最终 docx 未检测到原生 OOXML Math。")
            if counts.get("content_formulas", 0) and rendered_formulas < counts.get("content_formulas", 0):
                add(
                    "FORMULA_COUNT_MISMATCH",
                    "warning",
                    "最终 docx 中的原生公式数量少于内容提取数量，可能有公式被丢失或转成普通文本。",
                    f"content={counts.get('content_formulas')} rendered={rendered_formulas} docx={counts.get('docx_oMath')}",
                )
            if counts.get("content_images", 0) and counts.get("docx_drawings", 0) == 0:
                add("IMAGE_NOT_RENDERED", "error", "内容中有图片，但最终 docx 未检测到 drawing。")
            if counts.get("content_images", 0) and rendered_images < counts.get("content_images", 0):
                add(
                    "IMAGE_COUNT_MISMATCH",
                    "error",
                    "最终 docx 中的图片数量少于内容提取数量，可能有图片未插入。",
                    f"content={counts.get('content_images')} rendered={rendered_images} docx={counts.get('docx_drawings')}",
                )
            if counts.get("content_tables", 0) and rendered_tables < counts.get("content_tables", 0):
                add(
                    "TABLE_COUNT_MISMATCH",
                    "warning",
                    "最终 docx 中的表格数量少于内容提取数量，可能有表格未渲染。",
                    f"content={counts.get('content_tables')} rendered={rendered_tables} docx={counts.get('docx_tables')}",
                )
            if counts.get("content_text_chars", 0) > 200 and counts.get("docx_text_chars", 0) < counts.get("content_text_chars", 0) * 0.6:
                add(
                    "DOCX_TEXT_TOO_SHORT",
                    "error",
                    "最终 docx 文本量明显少于提取内容，可能发生正文丢失。",
                    f"content={counts.get('content_text_chars')} docx={counts.get('docx_text_chars')}",
                )
            missing_headings = _missing_heading_samples(content, plain)
            if missing_headings:
                add("CONTENT_HEADING_MISSING", "warning", "部分内容标题没有出现在最终 docx 中。", " / ".join(missing_headings))
            duplicate_front = _duplicate_front_matter_headings(content, xml)
            if duplicate_front:
                add(
                    "DUPLICATE_FRONT_MATTER_HEADING",
                    "error",
                    "最终 docx 中检测到重复的摘要/关键词等前置章节标题。",
                    " / ".join(duplicate_front),
                )
            final_placeholders = _placeholder_samples_from_texts(_xml_paragraph_texts(xml))
            if final_placeholders:
                add("PLACEHOLDER_TEXT_LEFT", "error", "最终 docx 中残留模板占位符或待补全文本。", " / ".join(final_placeholders[:8]))
            if re.search(r"Error!\s*(Reference source not found|Bookmark not defined)|错误！未找到", plain, re.I):
                add("WORD_FIELD_ERROR", "warning", "最终 docx 中可能存在 Word 域错误文本。")
            plain_compact = re.sub(r"\s+", "", plain)
            has_toc_text = "目录" in plain_compact or "Contents" in plain
            has_toc_field = r"TOC \o" in xml or r"TOC\\o" in xml
            if len(content.get("sections") or []) >= 3 and not (has_toc_text or has_toc_field):
                add("TOC_MISSING", "warning", "最终文档中未检测到目录文本。")
        except Exception as exc:
            add("DOCX_XML_UNREADABLE", "error", "最终 docx 无法读取 document.xml。", str(exc))

    passed = not any(i["severity"] == "error" for i in issues)
    next_action = (
        "通过 QA。仍建议用 WPS/Word 做最终视觉核对。"
        if passed else
        ("用户模式：根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
         if mode == "user" else
         "开发者模式：根据 active_owner 修改核心引擎脚本后重跑完整流水线。")
    )

    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": passed,
        "counts": counts,
        "issues": issues,
        "next_action": next_action,
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    return report


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# QA 检测报告",
        "",
        f"- 模式：`{report.get('mode')}`",
        f"- 结果：{'通过' if report.get('passed') else '未通过'}",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action')}",
        "",
        "## 统计",
        "",
    ]
    counts = report.get("counts") or {}
    if counts:
        for key in sorted(counts):
            lines.append(f"- `{key}`: {counts[key]}")
    else:
        lines.append("- 无统计信息")

    lines.extend(["", "## 问题", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- 未发现结构性问题。")
    else:
        for item in issues:
            lines.append(
                f"- **{item.get('severity')}** `{item.get('code')}`：{item.get('message')} "
                f"修复目标：`{item.get('active_owner')}`"
            )
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    repair_plan = report.get("repair_plan") or {}
    if repair_plan:
        lines.extend(["", "## 修复计划", ""])
        lines.append(f"- 摘要：{repair_plan.get('summary')}")
        commands = repair_plan.get("commands") or {}
        if commands.get("rerun_current_pipeline"):
            lines.append(f"- 重新跑完整流水线：`{commands.get('rerun_current_pipeline')}`")
        if commands.get("rebuild_current_docx"):
            lines.append(f"- 只重建当前 DOCX：`{commands.get('rebuild_current_docx')}`")
        steps = repair_plan.get("steps") or []
        if steps:
            lines.extend(["", "### 建议步骤", ""])
            for idx, step in enumerate(steps, 1):
                lines.append(f"{idx}. **{step.get('code')}**：{step.get('title')}")
                lines.append(f"   - 原因：{step.get('why')}")
                lines.append(f"   - 小白用户下一步：{step.get('user_action')}")
                lines.append(f"   - 开发者检查：{step.get('developer_action')}")
                if step.get("detail"):
                    lines.append(f"   - 细节：`{step.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def repair_plan_to_markdown(plan: Dict[str, Any]) -> str:
    lines = [
        "# QA 修复向导",
        "",
        f"- 结果：{'已通过' if plan.get('passed') else '需要修复'}",
        f"- 摘要：{plan.get('summary') or ''}",
        f"- 输出目录：`{plan.get('output_dir') or ''}`",
        "",
        "## 先打开这些文件",
        "",
    ]
    for item in plan.get("open_first") or []:
        lines.append(f"- `{item}`")
    commands = plan.get("commands") or {}
    lines.extend(["", "## 可执行命令", ""])
    if commands.get("rerun_current_pipeline"):
        lines.append(f"- 重新跑完整流水线：`{commands.get('rerun_current_pipeline')}`")
    if commands.get("rebuild_current_docx"):
        lines.append(f"- 修改 `build_generated.py` 后只重建当前 DOCX：`{commands.get('rebuild_current_docx')}`")
    if not commands.get("rerun_current_pipeline") and not commands.get("rebuild_current_docx"):
        lines.append("- 暂无可自动推断的命令。")
    steps = plan.get("steps") or []
    lines.extend(["", "## 修复步骤", ""])
    if not steps:
        lines.append("- 当前没有 QA 问题。")
    else:
        for idx, step in enumerate(steps, 1):
            lines.append(f"### {idx}. {step.get('title') or step.get('code')}")
            lines.append("")
            lines.append(f"- 问题码：`{step.get('code')}`")
            lines.append(f"- 级别：`{step.get('severity')}`")
            lines.append(f"- 可能原因：{step.get('why')}")
            lines.append(f"- 小白用户下一步：{step.get('user_action')}")
            lines.append(f"- 开发者检查：{step.get('developer_action')}")
            lines.append(f"- 修复目标：`{step.get('target')}`")
            if step.get("detail"):
                lines.append(f"- 细节：`{step.get('detail')}`")
            counts = step.get("counts") or {}
            if counts:
                lines.append("- 数量： " + ", ".join(f"`{k}={v}`" for k, v in sorted(counts.items())))
            lines.append("")
    prompt = plan.get("copy_to_ai_prompt")
    if prompt:
        lines.extend(["## 可直接发给 AI 的修复请求", "", "```text", str(prompt), "```", ""])
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    json_path = os.path.join(out_dir, "qa_report.json")
    md_path = os.path.join(out_dir, "qa_report.md")
    repair_json_path = os.path.join(out_dir, "qa_repair_plan.json")
    repair_md_path = os.path.join(out_dir, "qa_repair_plan.md")
    prompt_path = os.path.join(out_dir, "qa_fix_prompt.txt")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
    repair_plan = report.get("repair_plan") or {}
    with open(repair_json_path, "w", encoding="utf-8") as f:
        json.dump(repair_plan, f, ensure_ascii=False, indent=2)
    with open(repair_md_path, "w", encoding="utf-8") as f:
        f.write(repair_plan_to_markdown(repair_plan))
    with open(prompt_path, "w", encoding="utf-8") as f:
        f.write(str(repair_plan.get("copy_to_ai_prompt") or ""))


def check_and_write(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    report = check_output(out_dir, mode=mode, output_docx_name=output_docx_name)
    write_reports(report, out_dir)
    return report


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Check generated pipeline output.")
    parser.add_argument("out_dir")
    parser.add_argument("--mode", choices=sorted(VALID_MODES), default="user")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()

    result = check_and_write(args.out_dir, mode=args.mode, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)

"""Visual QA orchestration for generated DOCX outputs."""
from __future__ import annotations

import os
import shutil
from datetime import datetime
from typing import Any, Dict, List

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value

try:
    from qa_visual_modules.exporters import _export_pdf, _export_wps_pdf
    from qa_visual_modules.golden import _compare_or_update_golden
    from qa_visual_modules.image_stats import _image_stats
    from qa_visual_modules.pdf_tools import _find_page, _pdf_pages_text, _pdfinfo, _render_all_pages, _render_samples, _sample_pages
except ImportError:  # pragma: no cover - package-style imports
    from .exporters import _export_pdf, _export_wps_pdf
    from .golden import _compare_or_update_golden
    from .image_stats import _image_stats
    from .pdf_tools import _find_page, _pdf_pages_text, _pdfinfo, _render_all_pages, _render_samples, _sample_pages

def _issue(code: str, severity: str, message: str, detail: str = "") -> Dict[str, Any]:
    return {"code": code, "severity": severity, "message": message, "detail": detail}


def _float_or_none(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


def _page_size_differs(a_width: Any, a_height: Any, b_width: Any, b_height: Any, *, tolerance_pt: float = 2.0) -> bool:
    aw = _float_or_none(a_width)
    ah = _float_or_none(a_height)
    bw = _float_or_none(b_width)
    bh = _float_or_none(b_height)
    if aw is None or ah is None or bw is None or bh is None:
        return False
    return abs(aw - bw) > tolerance_pt or abs(ah - bh) > tolerance_pt


def _nonblank_page_count(pages: List[str]) -> int:
    return len([page for page in pages or [] if str(page or "").strip()])


def _text_page_count_differs(word_text_pages: Any, wps_text_pages: Any) -> bool:
    try:
        word_count = int(word_text_pages or 0)
        wps_count = int(wps_text_pages or 0)
    except (TypeError, ValueError):
        return False
    if word_count <= 0:
        return False
    if wps_count <= 0:
        return True
    return (word_count - wps_count) > max(2, word_count // 8)


def _preserve_rendered_text_artifact(visual_dir: str, label: str) -> str | None:
    src = os.path.join(visual_dir, "rendered.txt")
    if not os.path.exists(src):
        return None
    dst = os.path.join(visual_dir, f"rendered_{label}.txt")
    shutil.copyfile(src, dst)
    return dst


def _restore_rendered_text_artifact(visual_dir: str, label: str) -> None:
    src = os.path.join(visual_dir, f"rendered_{label}.txt")
    if os.path.exists(src):
        shutil.copyfile(src, os.path.join(visual_dir, "rendered.txt"))


def _next_action(issues: List[Dict[str, Any]]) -> str:
    error_codes = {str(item.get("code") or "") for item in issues if item.get("severity") == "error"}
    warning_codes = {str(item.get("code") or "") for item in issues if item.get("severity") == "warning"}
    if not error_codes and not warning_codes:
        return "视觉 QA 的机器检查已通过；仍建议用 Word/WPS 打开最终 DOCX 做人工视觉核对。"
    if error_codes & {"PDF_EXPORT_FAILED"}:
        return "修复 Microsoft Word COM/PDF 导出环境后重跑 visual QA；若 DOCX 本身无法打开，先重新生成最终论文。"
    if error_codes & {"PDFINFO_UNAVAILABLE", "PDFTOTEXT_UNAVAILABLE", "SAMPLE_RENDER_FAILED", "ALL_PAGE_RENDER_FAILED"}:
        return "安装或修复 Poppler 命令行工具（pdfinfo、pdftotext、pdftoppm）后重跑 visual QA。"
    if error_codes & {"PDFINFO_FAILED"}:
        return "打开 visual_report.md 查看 pdfinfo 错误；修复 PDF 导出文件或 Poppler 环境后重跑 visual QA。"
    if error_codes & {"PDF_PAGE_COUNT_INVALID"}:
        return "PDF 导出后没有有效页面；先用 Word 打开最终 DOCX 检查文件，再重新导出并重跑 visual QA。"
    if error_codes & {"PDFTOTEXT_FAILED"}:
        return "打开 visual_report.md 查看 pdftotext 错误；修复 PDF 导出或 Poppler 环境后重跑 visual QA。"
    if error_codes & {"PAGE_IMAGE_UNREADABLE"}:
        return "打开 visual_report.md 查看不可读页面；修复 PDF 渲染或页面 PNG 生成后重跑 visual QA。"
    if error_codes & {"WPS_EXPORT_UNAVAILABLE"}:
        return "安装/配置 WPS COM，或取消 --require-wps 后重跑 visual QA。"
    if error_codes & {"WPS_PDFINFO_UNAVAILABLE", "WPS_PDFINFO_FAILED"}:
        return "WPS 已导出 PDF，但无法读取 WPS PDF 页面信息；先确认 WPS 导出的 PDF 能正常打开，再修复 PDF/Poppler 环境并重跑 visual QA。"
    if error_codes & {"WPS_PAGE_COUNT_INVALID"}:
        return "WPS 导出的 PDF 没有有效页面；先用 WPS 打开最终 DOCX 和导出的 PDF 检查是否为空白，修复后重跑 visual QA。"
    if error_codes & {"WPS_PAGE_COUNT_MISMATCH"}:
        return "分别打开 Word 与 WPS 导出的 PDF 比对分页差异；确认是兼容性差异还是排版脚本问题。修复后重跑 visual QA。"
    if error_codes & {"WPS_PAGE_SIZE_MISMATCH"}:
        return "分别打开 Word 与 WPS 导出的 PDF 比对纸张大小、页面尺寸和横竖方向；修复模板页面设置或 WPS 兼容性问题后重跑 visual QA。"
    if error_codes & {"WPS_TEXT_PAGE_MISMATCH"}:
        return "分别打开 Word 与 WPS 导出的 PDF 比对正文、目录、公式和图片内容；WPS 文本页明显缺失时先修复 WPS 导出、字体兼容或排版生成问题，再重跑 visual QA。"
    if error_codes & {"GOLDEN_BASELINE_MISMATCH"}:
        return "打开 visual_report.md 和 visual_qa/samples/ 对比页面；确认变化正确则用 --update-golden 更新基线，否则继续修复排版。"
    if error_codes & {"MISSING_DOCX"}:
        return "先修复构建阶段，确保最终论文 DOCX 生成后再运行 visual QA。"
    if warning_codes & {"MANY_BLANK_PAGES"}:
        return "visual QA 通过但发现较多文本空白页；打开导出的 PDF 核对空白页，异常时修复分页/分节逻辑后重跑 visual QA。"
    if warning_codes & {"TOC_TEXT_NOT_FOUND"}:
        return "visual QA 通过但未在 PDF 文本中找到目录；打开导出的 PDF 核对目录页，缺失时检查 TOC 生成或 Word 字段更新后重跑 visual QA。"
    if warning_codes & {"MANY_BLANK_PAGE_IMAGES"}:
        return "visual QA 通过但全页 PNG 样张中有较多疑似空白页；打开 visual_qa/samples/ 和 all_pages 核对，异常时修复分页或 PDF 渲染后重跑 visual QA。"
    if warning_codes & {"GOLDEN_BASELINE_MISSING"}:
        return "visual QA 通过但缺少黄金基线；首次建立视觉基线时可用 --update-golden 生成，若不需要基线则取消 golden 参数后重跑 visual QA。"
    if warning_codes & {"WPS_EXPORT_UNAVAILABLE"}:
        return "visual QA 通过但 WPS 交叉渲染不可用；需要 WPS 校验时安装/配置 WPS COM，否则可忽略该 warning 或取消 --require-wps 后重跑 visual QA。"
    if warning_codes & {"WPS_PDFINFO_UNAVAILABLE", "WPS_PDFINFO_FAILED"}:
        return "visual QA 通过但 WPS PDF 页面信息不可读；需要 WPS 校验时先确认 WPS 导出的 PDF 能打开，再修复 PDF/Poppler 环境并重跑 visual QA。"
    if warning_codes & {"WPS_PAGE_COUNT_INVALID"}:
        return "visual QA 通过但 WPS 导出的 PDF 没有有效页面；需要 WPS 校验时先用 WPS 打开 DOCX/PDF 检查是否为空白，修复后重跑 visual QA。"
    if warning_codes & {"WPS_PAGE_SIZE_MISMATCH"}:
        return "visual QA 通过但 WPS 与 Word 的页面尺寸不同；需要 WPS 校验时比对纸张大小和横竖方向，修复后重跑 visual QA。"
    if warning_codes & {"WPS_TEXT_PAGE_MISMATCH"}:
        return "visual QA 通过但 WPS 与 Word 的可提取文本页数不同；需要 WPS 校验时打开两份 PDF 比对内容，修复后重跑 visual QA。"
    if warning_codes:
        return "visual QA 通过但仍有 warning；打开 visual_report.md 按问题码确认是否影响交付，必要时修复后重跑 visual QA。"
    return "打开 visual_report.md 和 visual_qa/samples/，按页面样张定位排版问题后重跑流水线。"


def check_visual(
    out_dir: str,
    output_docx_name: str = "最终论文.docx",
    project_root: str | None = None,
    render_all_pages: bool = True,
    require_wps: bool = False,
    golden_dir: str | None = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    out_dir = os.path.abspath(out_dir)
    docx_path = os.path.join(out_dir, output_docx_name)
    visual_dir = os.path.join(out_dir, "visual_qa")
    os.makedirs(visual_dir, exist_ok=True)

    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}
    artifacts: Dict[str, Any] = {}

    if not os.path.exists(docx_path):
        issues.append(_issue("MISSING_DOCX", "error", "Cannot run visual QA because final DOCX is missing.", docx_path))
    else:
        try:
            pdf_path = _export_pdf(docx_path, visual_dir)
            artifacts["pdf"] = pdf_path
            info = _pdfinfo(pdf_path)
            counts.update({k: v for k, v in info.items() if k in {"pages", "page_width_pt", "page_height_pt"}})
            if not info.get("available"):
                issues.append(_issue("PDFINFO_UNAVAILABLE", "error", "pdfinfo is not available; visual QA cannot verify page count or paper size."))
            elif info.get("error"):
                issues.append(_issue("PDFINFO_FAILED", "error", "pdfinfo failed.", str(info.get("error"))))
            if int(info.get("pages") or 0) <= 0:
                issues.append(_issue("PDF_PAGE_COUNT_INVALID", "error", "Rendered PDF has no pages."))

            text_tool_available = shutil.which("pdftotext") is not None
            pages_text = _pdf_pages_text(pdf_path, visual_dir)
            word_text_artifact = _preserve_rendered_text_artifact(visual_dir, "word")
            if word_text_artifact:
                artifacts["rendered_text"] = os.path.join(visual_dir, "rendered.txt")
                artifacts["word_text"] = word_text_artifact
            counts["text_pages"] = _nonblank_page_count(pages_text)
            if not text_tool_available:
                issues.append(_issue("PDFTOTEXT_UNAVAILABLE", "error", "pdftotext is not available; visual QA cannot verify rendered text."))
            elif not pages_text:
                issues.append(_issue("PDFTOTEXT_FAILED", "error", "pdftotext did not produce readable page text."))
            blank_pages = [idx + 1 for idx, page in enumerate(pages_text) if not page.strip()]
            if len(blank_pages) > max(2, int(info.get("pages") or 0) // 8):
                issues.append(_issue("MANY_BLANK_PAGES", "warning", "Rendered PDF has many text-empty pages.", ",".join(map(str, blank_pages[:12]))))
            if pages_text and not _find_page(pages_text, [r"目录", r"contents"]) and int(info.get("pages") or 0) >= 6:
                issues.append(_issue("TOC_TEXT_NOT_FOUND", "warning", "Rendered PDF text does not expose a TOC page."))

            page_count = int(info.get("pages") or 0)
            samples = _sample_pages(page_count, pages_text)
            rendered = _render_samples(pdf_path, visual_dir, samples)
            artifacts["samples"] = rendered
            counts["sample_pages"] = samples
            counts["sample_images"] = len(rendered)
            if samples and len(rendered) < len(samples):
                issues.append(_issue("SAMPLE_RENDER_FAILED", "error", "Could not render all PDF sample PNGs; install pdftoppm/Poppler."))
            image_stats: Dict[str, Any] = {"page_hashes": [], "blank_pages": []}
            if render_all_pages and page_count > 0:
                all_pages = _render_all_pages(pdf_path, visual_dir, page_count)
                artifacts["all_pages"] = all_pages
                counts["all_page_images"] = len(all_pages)
                if len(all_pages) != page_count:
                    issues.append(_issue("ALL_PAGE_RENDER_FAILED", "error", "Could not render every PDF page to PNG.", f"pages={page_count} rendered={len(all_pages)}"))
                image_stats = _image_stats(all_pages)
                counts["blank_page_images"] = len(image_stats.get("blank_pages") or [])
                counts["image_hashes"] = len(image_stats.get("page_hashes") or [])
                if image_stats.get("unreadable_pages"):
                    issues.append(_issue("PAGE_IMAGE_UNREADABLE", "error", "Some rendered page PNGs could not be inspected.", ",".join(map(str, image_stats.get("unreadable_pages")[:12]))))
                if len(image_stats.get("blank_pages") or []) > max(2, page_count // 8):
                    issues.append(_issue("MANY_BLANK_PAGE_IMAGES", "warning", "Rendered all-page PNG set contains many blank-looking pages.", ",".join(map(str, image_stats.get("blank_pages")[:12]))))
            golden = _compare_or_update_golden(out_dir, counts, pages_text, image_stats, golden_dir, update_golden)
            artifacts["golden_baseline"] = golden
            if golden.get("status") == "mismatch":
                issues.append(_issue("GOLDEN_BASELINE_MISMATCH", "error", "Rendered output differs from the golden baseline.", " / ".join(golden.get("issues") or [])))
            elif golden.get("enabled") and golden.get("status") == "missing":
                issues.append(_issue("GOLDEN_BASELINE_MISSING", "warning", "Golden baseline was requested but no baseline exists."))
            try:
                wps_pdf = _export_wps_pdf(docx_path, visual_dir)
                artifacts["wps_pdf"] = wps_pdf
                wps_info = _pdfinfo(wps_pdf)
                counts["wps_pages"] = wps_info.get("pages")
                counts["wps_page_width_pt"] = wps_info.get("page_width_pt")
                counts["wps_page_height_pt"] = wps_info.get("page_height_pt")
                wps_severity = "error" if require_wps else "warning"
                if not wps_info.get("available"):
                    issues.append(_issue("WPS_PDFINFO_UNAVAILABLE", wps_severity, "pdfinfo is not available for the WPS-rendered PDF; visual QA cannot verify WPS pages."))
                elif wps_info.get("error"):
                    issues.append(_issue("WPS_PDFINFO_FAILED", wps_severity, "pdfinfo failed for the WPS-rendered PDF.", str(wps_info.get("error"))))
                elif int(wps_info.get("pages") or 0) <= 0:
                    issues.append(_issue("WPS_PAGE_COUNT_INVALID", wps_severity, "WPS-rendered PDF has no valid pages."))
                elif page_count and int(wps_info.get("pages")) != page_count:
                    issues.append(_issue("WPS_PAGE_COUNT_MISMATCH", "error", "WPS-rendered PDF page count differs from Word-rendered PDF.", f"word={page_count} wps={wps_info.get('pages')}"))
                elif _page_size_differs(
                    info.get("page_width_pt"),
                    info.get("page_height_pt"),
                    wps_info.get("page_width_pt"),
                    wps_info.get("page_height_pt"),
                ):
                    issues.append(
                        _issue(
                            "WPS_PAGE_SIZE_MISMATCH",
                            "error",
                            "WPS-rendered PDF page size differs from Word-rendered PDF.",
                            (
                                f"word={info.get('page_width_pt')}x{info.get('page_height_pt')} "
                                f"wps={wps_info.get('page_width_pt')}x{wps_info.get('page_height_pt')}"
                            ),
                        )
                    )
                else:
                    wps_pages_text = _pdf_pages_text(wps_pdf, visual_dir)
                    wps_text_artifact = _preserve_rendered_text_artifact(visual_dir, "wps")
                    if wps_text_artifact:
                        artifacts["wps_text"] = wps_text_artifact
                    _restore_rendered_text_artifact(visual_dir, "word")
                    counts["wps_text_pages"] = _nonblank_page_count(wps_pages_text)
                    if _text_page_count_differs(counts.get("text_pages"), counts.get("wps_text_pages")):
                        issues.append(
                            _issue(
                                "WPS_TEXT_PAGE_MISMATCH",
                                "error",
                                "WPS-rendered PDF has substantially fewer extractable text pages than Word-rendered PDF.",
                                f"word_text_pages={counts.get('text_pages')} wps_text_pages={counts.get('wps_text_pages')}",
                            )
                        )
            except Exception as exc:
                severity = "error" if require_wps else "warning"
                issues.append(_issue("WPS_EXPORT_UNAVAILABLE", severity, "WPS PDF export could not be completed.", str(exc)))
        except Exception as exc:
            issues.append(_issue("PDF_EXPORT_FAILED", "error", "DOCX could not be exported to PDF for visual QA.", str(exc)))

    passed = not any(i.get("severity") == "error" for i in issues)
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "passed": passed,
        "output_dir_name": os.path.basename(out_dir),
        "counts": counts,
        "issues": sanitize_value(issues, project_root),
        "artifacts": sanitize_value(artifacts, project_root),
        "next_action": _next_action(issues),
    }



"""Build-script generation orchestration."""
from __future__ import annotations

import json
import os
import shutil

from .runtime_template import RUNTIME_TEMPLATE
from .sections import _front_matter_sections, _normalize_numbered_section_order
from .style_profiles import _infer_style_profiles
from .template_rules import _extract_page_and_header, _infer_template_rules


def _copy_runtime_dependency(src: str, dst: str) -> str:
    if not src or not os.path.isdir(src):
        return src
    if os.path.abspath(src) != os.path.abspath(dst):
        shutil.rmtree(dst, ignore_errors=True)
        shutil.copytree(src, dst)
    return os.path.basename(dst)


def _copy_latex_runtime(output_dir: str) -> None:
    pipeline_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    latex_src = os.path.join(pipeline_dir, 'latex_omath.py')
    if os.path.exists(latex_src):
        shutil.copy2(latex_src, os.path.join(output_dir, 'latex_omath.py'))
    latex_modules_src = os.path.join(pipeline_dir, 'latex_omath_modules')
    if os.path.isdir(latex_modules_src):
        latex_modules_dst = os.path.join(output_dir, 'latex_omath_modules')
        shutil.rmtree(latex_modules_dst, ignore_errors=True)
        shutil.copytree(latex_modules_src, latex_modules_dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))


def generate(format_json_path: str, content_json_path: str, output_dir: str, output_docx_name: str = '最终论文.docx') -> int:
    os.makedirs(output_dir, exist_ok=True)
    with open(format_json_path, 'r', encoding='utf-8') as f:
        fmt = json.load(f)
    with open(content_json_path, 'r', encoding='utf-8') as f:
        cnt = json.load(f)

    profiles = _infer_style_profiles(fmt)
    page = _extract_page_and_header(fmt)
    front = _front_matter_sections(cnt)
    cover_info = cnt.get('cover_info', {}) or {}
    title_cn = cover_info.get('paper_title') or cnt.get('title_info', {}).get('title_cn') or ''
    rules = _infer_template_rules(fmt)

    src_assets = fmt.get('_meta', {}).get('assets_dir') or ''
    runtime_assets = _copy_runtime_dependency(src_assets, os.path.join(output_dir, 'assets'))

    src_images = cnt.get('_meta', {}).get('images_dir') or ''
    runtime_images = _copy_runtime_dependency(src_images, os.path.join(output_dir, 'figures'))

    _copy_latex_runtime(output_dir)

    normalized_sections = _normalize_numbered_section_order(cnt.get('sections', []))

    data_blob = {
        'fmt_meta': fmt.get('_meta', {}),
        'content_meta': cnt.get('_meta', {}),
        'page': page,
        'profiles': profiles,
        'cover': fmt.get('cover', []),
        'cover_info': cover_info,
        'title_cn': title_cn,
        'sections': normalized_sections,
        'references': cnt.get('references', []),
        'front_indices': sorted(front['front_indices']),
        'front': {k: v for k, v in front.items() if k != 'front_indices'},
        'images_dir': runtime_images,
        'assets_dir': runtime_assets,
        'rules': rules,
    }
    blob = json.dumps(data_blob, ensure_ascii=False, indent=2)
    code = RUNTIME_TEMPLATE.replace('__DATA_BLOB__', repr(blob)).replace('__OUT_DOCX__', repr(os.path.basename(output_docx_name)))
    out_py = os.path.join(output_dir, 'build_generated.py')
    with open(out_py, 'w', encoding='utf-8') as f:
        f.write(code)
    return len(code)

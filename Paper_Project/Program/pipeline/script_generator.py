"""
script_generator.py — role-driven DOCX build script generator.

输入: format.json + content.json
输出: build_generated.py；运行该脚本生成最终论文 docx。

修复点：
- 封面按模板元素重建，支持校徽/图片 assets，学位/学校编码行左对齐。
- 中文摘要、英文摘要、目录、正文使用独立 role；英文摘要另起一页，正文在目录后另起一节。
- 标题设置 outline level，编号和标题之间自动补一个空格。
- 正文 [N] / [1,2] / [1-3] 引用自动上标。
- 表格渲染为三线表，保留单元格换行。
- 图/表题走 caption role，居中并规范“图 1 标题”的空格。
- 代码块保留换行和空格，使用等宽字体。
- 参考文献单条一个段落，悬挂缩进，中文宋体/英文 Times New Roman 混排。
"""
from __future__ import annotations

import json
import os
import re
import shutil
from typing import Any, Dict, List, Optional

try:
    from script_generator_modules.sections import (
        _front_matter_sections,
        _normalize_numbered_section_order,
    )
    from script_generator_modules.template_rules import (
        _extract_page_and_header,
        _infer_template_rules,
    )
    from script_generator_modules.runtime_base import BASE_RUNTIME
    from script_generator_modules.runtime_media_tables import MEDIA_TABLE_RUNTIME
    from script_generator_modules.runtime_references import REFERENCES_RUNTIME
    from script_generator_modules.runtime_formula import FORMULA_RUNTIME
    from script_generator_modules.runtime_formula_text import FORMULA_TEXT_RUNTIME
    from script_generator_modules.runtime_content_helpers import CONTENT_HELPERS_RUNTIME
    from script_generator_modules.runtime_formula_render import FORMULA_RENDER_RUNTIME
    from script_generator_modules.runtime_toc import TOC_RUNTIME
    from script_generator_modules.runtime_build import BUILD_RUNTIME
    from script_generator_modules.runtime_cover import COVER_RUNTIME
    from script_generator_modules.runtime_front_matter import FRONT_MATTER_RUNTIME
    from script_generator_modules.runtime_body import BODY_RUNTIME
    from script_generator_modules.style_profiles import (
        _CN_SIZE_PATTERNS,
        _align_from_text,
        _apply_template_text_rules,
        _ascii_ratio,
        _find_instruction,
        _find_regex_instruction,
        _first_run,
        _first_text_run,
        _font_from_text,
        _has_format_instruction,
        _indent_from_text,
        _infer_style_profiles,
        _is_cjk,
        _line_spacing_from_text,
        _normalize_profile,
        _profile_from_instruction,
        _profile_from_para,
        _profile_from_para_first_text,
        _size_from_text,
        _spacing_before_after_from_text,
        _text_blob,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .script_generator_modules.sections import (
        _front_matter_sections,
        _normalize_numbered_section_order,
    )
    from .script_generator_modules.template_rules import (
        _extract_page_and_header,
        _infer_template_rules,
    )
    from .script_generator_modules.runtime_base import BASE_RUNTIME
    from .script_generator_modules.runtime_media_tables import MEDIA_TABLE_RUNTIME
    from .script_generator_modules.runtime_references import REFERENCES_RUNTIME
    from .script_generator_modules.runtime_formula import FORMULA_RUNTIME
    from .script_generator_modules.runtime_formula_text import FORMULA_TEXT_RUNTIME
    from .script_generator_modules.runtime_content_helpers import CONTENT_HELPERS_RUNTIME
    from .script_generator_modules.runtime_formula_render import FORMULA_RENDER_RUNTIME
    from .script_generator_modules.runtime_toc import TOC_RUNTIME
    from .script_generator_modules.runtime_build import BUILD_RUNTIME
    from .script_generator_modules.runtime_cover import COVER_RUNTIME
    from .script_generator_modules.runtime_front_matter import FRONT_MATTER_RUNTIME
    from .script_generator_modules.runtime_body import BODY_RUNTIME
    from .script_generator_modules.style_profiles import (
        _CN_SIZE_PATTERNS,
        _align_from_text,
        _apply_template_text_rules,
        _ascii_ratio,
        _find_instruction,
        _find_regex_instruction,
        _first_run,
        _first_text_run,
        _font_from_text,
        _has_format_instruction,
        _indent_from_text,
        _infer_style_profiles,
        _is_cjk,
        _line_spacing_from_text,
        _normalize_profile,
        _profile_from_instruction,
        _profile_from_para,
        _profile_from_para_first_text,
        _size_from_text,
        _spacing_before_after_from_text,
        _text_blob,
    )


RUNTIME_TEMPLATE = r'''
__BASE_RUNTIME__

__COVER_RUNTIME__

__FORMULA_RUNTIME__

__TOC_RUNTIME__

__FRONT_MATTER_RUNTIME__

__CONTENT_HELPERS_RUNTIME__

__FORMULA_TEXT_RUNTIME__

__FORMULA_RENDER_RUNTIME__

__MEDIA_TABLE_RUNTIME__

__REFERENCES_RUNTIME__

__BODY_RUNTIME__

__BUILD_RUNTIME__
'''

RUNTIME_TEMPLATE = (
    RUNTIME_TEMPLATE
    .replace('__BASE_RUNTIME__', BASE_RUNTIME)
    .replace('__COVER_RUNTIME__', COVER_RUNTIME)
    .replace('__FORMULA_RUNTIME__', FORMULA_RUNTIME)
    .replace('__CONTENT_HELPERS_RUNTIME__', CONTENT_HELPERS_RUNTIME)
    .replace('__FORMULA_TEXT_RUNTIME__', FORMULA_TEXT_RUNTIME)
    .replace('__FORMULA_RENDER_RUNTIME__', FORMULA_RENDER_RUNTIME)
    .replace('__MEDIA_TABLE_RUNTIME__', MEDIA_TABLE_RUNTIME)
    .replace('__REFERENCES_RUNTIME__', REFERENCES_RUNTIME)
    .replace('__TOC_RUNTIME__', TOC_RUNTIME)
    .replace('__FRONT_MATTER_RUNTIME__', FRONT_MATTER_RUNTIME)
    .replace('__BODY_RUNTIME__', BODY_RUNTIME)
    .replace('__BUILD_RUNTIME__', BUILD_RUNTIME)
)



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
    runtime_assets = src_assets
    if src_assets and os.path.isdir(src_assets):
        dst_assets = os.path.join(output_dir, 'assets')
        if os.path.abspath(src_assets) != os.path.abspath(dst_assets):
            shutil.rmtree(dst_assets, ignore_errors=True)
            shutil.copytree(src_assets, dst_assets)
        runtime_assets = 'assets'

    src_images = cnt.get('_meta', {}).get('images_dir') or ''
    runtime_images = src_images
    if src_images and os.path.isdir(src_images):
        dst_images = os.path.join(output_dir, 'figures')
        if os.path.abspath(src_images) != os.path.abspath(dst_images):
            shutil.rmtree(dst_images, ignore_errors=True)
            shutil.copytree(src_images, dst_images)
        runtime_images = 'figures'

    latex_src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'latex_omath.py')
    if os.path.exists(latex_src):
        shutil.copy2(latex_src, os.path.join(output_dir, 'latex_omath.py'))

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


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 4:
        print('Usage: python script_generator.py format.json content.json output_dir [output.docx]')
        raise SystemExit(2)
    out_name = sys.argv[4] if len(sys.argv) > 4 else '最终论文.docx'
    n = generate(sys.argv[1], sys.argv[2], sys.argv[3], out_name)
    print(f'Generated build script, {n} chars')

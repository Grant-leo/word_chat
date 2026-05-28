"""Stable role-driven DOCX build script generator entry point."""
from __future__ import annotations

try:
    from script_generator_modules.generator import generate
    from script_generator_modules.runtime_template import RUNTIME_TEMPLATE
    from script_generator_modules.sections import _front_matter_sections, _normalize_numbered_section_order
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
    from script_generator_modules.template_rules import _extract_page_and_header, _infer_template_rules
except ImportError:  # pragma: no cover - package-style imports
    from .script_generator_modules.generator import generate
    from .script_generator_modules.runtime_template import RUNTIME_TEMPLATE
    from .script_generator_modules.sections import _front_matter_sections, _normalize_numbered_section_order
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
    from .script_generator_modules.template_rules import _extract_page_and_header, _infer_template_rules


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 4:
        print('Usage: python script_generator.py format.json content.json output_dir [output.docx]')
        raise SystemExit(2)
    out_name = sys.argv[4] if len(sys.argv) > 4 else '最终论文.docx'
    n = generate(sys.argv[1], sys.argv[2], sys.argv[3], out_name)
    print(f'Generated build script, {n} chars')

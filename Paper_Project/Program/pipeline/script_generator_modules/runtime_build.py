"""Document build orchestration and manifest runtime template fragment for generated build scripts."""
from __future__ import annotations

BUILD_RUNTIME = r'''
FORMULA_COUNTERS = {}
TABLE_COUNTERS = {}


def reset_build_stats():
    global BUILD_STATS
    BUILD_STATS = {
        'content_images_rendered': 0,
        'content_tables_rendered': 0,
        'content_formulas_rendered': 0,
        'content_image_fragments_contained': 0,
        'inline_formulas_rendered': 0,
        'display_formulas_rendered': 0,
    }


def write_build_manifest():
    path = os.path.join(BASE, 'build_manifest.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({'schema_version': 1, 'counts': BUILD_STATS}, f, ensure_ascii=False, indent=2)


def build_document(toc_page_map=None, native_toc=True):
    """Build the whole DOCX once, optionally with resolved static TOC pages."""
    global doc, TOC_PAGE_MAP, USE_NATIVE_TOC, FORMULA_COUNTERS, TABLE_COUNTERS
    TOC_PAGE_MAP = dict(toc_page_map or {})
    USE_NATIVE_TOC = bool(native_toc)
    FORMULA_COUNTERS = {}
    TABLE_COUNTERS = {}
    reset_build_stats()
    doc = Document()
    configure_global_styles()
    setup_section(doc.sections[0])
    clear_header_footer(doc.sections[0])
    remove_initial_empty_paragraph()
    render_cover_and_declarations()
    render_front_matter()
    render_body()
    force_cover_headerless()
    doc.save(OUT)


def main():
    build_document({}, native_toc=False)
    page_map = _infer_heading_pages_from_word_com()
    if page_map:
        build_document(page_map, native_toc=False)
    write_build_manifest()
    suffix = '目录页码已由 Word COM 解析' if page_map else '已生成静态目录行，页码未由 Word COM 解析'
    print(f'已保存: {os.path.basename(OUT)}  ({suffix})')


if __name__ == '__main__':
    main()
'''

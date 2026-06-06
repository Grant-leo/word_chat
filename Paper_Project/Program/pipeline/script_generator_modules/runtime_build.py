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
    BUILD_STATS['footnote_references_rendered'] = NOTE_REF_COUNTS.get('footnote', 0)
    BUILD_STATS['endnote_references_rendered'] = NOTE_REF_COUNTS.get('endnote', 0)
    BUILD_STATS['footnote_definitions_rendered'] = len(NOTE_DEFS.get('footnote') or {})
    BUILD_STATS['endnote_definitions_rendered'] = len(NOTE_DEFS.get('endnote') or {})
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({'schema_version': 1, 'counts': BUILD_STATS}, f, ensure_ascii=False, indent=2)


def build_document(toc_page_map=None, caption_page_map=None, native_toc=True):
    """Build the whole DOCX once, optionally with resolved static TOC pages."""
    global doc, TOC_PAGE_MAP, CAPTION_PAGE_MAP, USE_NATIVE_TOC, FORMULA_COUNTERS, TABLE_COUNTERS
    TOC_PAGE_MAP = dict(toc_page_map or {})
    CAPTION_PAGE_MAP = dict(caption_page_map or {})
    USE_NATIVE_TOC = bool(native_toc)
    FORMULA_COUNTERS = {}
    TABLE_COUNTERS = {}
    reset_build_stats()
    reset_note_state()
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
    apply_note_parts_to_docx(OUT)


def main():
    build_document({}, {}, native_toc=False)
    page_map = _infer_heading_pages_from_word_com()
    cap_page_map = _infer_caption_pages_from_word_com()
    if page_map or cap_page_map:
        build_document(page_map, cap_page_map, native_toc=False)
    write_build_manifest()
    parts = []
    if page_map:
        parts.append('目录页码')
    if cap_page_map:
        parts.append('图表页码')
    detail = ('、'.join(parts) + ' 已由 Word COM 解析') if parts else '已生成静态目录行，页码未由 Word COM 解析'
    print(f'已保存: {os.path.basename(OUT)}  ({detail})')


if __name__ == '__main__':
    main()
'''

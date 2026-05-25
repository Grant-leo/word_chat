"""
content_parser.py — 从文本资料 docx 提取结构化内容
输出: content.json (章节结构 + 段落 + 图片路径 + 表格 + 参考文献)
"""
from docx import Document
from docx.shared import Pt, Inches
from lxml import etree
import json, os, re, shutil, hashlib

def emu_to_pt(emu):
    if emu is None: return None
    return round(emu / 12700, 1)


def _math_text(elem):
    """Extract plain text from a math OOXML element."""
    M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    parts = []
    for t in elem.iter(f'{{{M}}}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


def extract_math(para):
    """Extract OOXML math elements from a paragraph. Returns (text, math_list).
    math_list entries: {'type': 'inline'|'display', 'xml': escaped_xml_string, 'text': plain_text}
    Text is cleaned of formula garbling."""
    xml = para._element.xml
    if 'm:oMath' not in xml and 'oMathPara' not in xml:
        return para.text, []

    math_list = []
    root = para._element

    # Extract m:oMathPara (display formulas) — these ARE the paragraph
    for omp in root.findall('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara'):
        raw = etree.tounicode(omp, with_tail=False)
        math_list.append({'type': 'display', 'xml': raw, 'text': _math_text(omp)})

    # Extract m:oMath (inline formulas) — embedded in runs
    for om in root.iter('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath'):
        # Skip if already inside an oMathPara
        parent_tag = om.getparent().tag.split('}')[-1] if '}' in om.getparent().tag else om.getparent().tag
        if parent_tag == 'oMathPara':
            continue
        raw = etree.tounicode(om, with_tail=False)
        math_list.append({'type': 'inline', 'xml': raw, 'text': _math_text(om)})

    # Reconstruct text without formula garbling: get text from runs, skip math-only runs
    text_parts = []
    for child in root:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'r':
            # Skip runs that contain only math (no w:t)
            has_text = child.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') is not None
            if has_text:
                text_parts.append(child.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text or '')
        elif tag == 'oMathPara' or tag == 'oMath':
            # Don't add math XML to text
            pass
        elif tag == 'pPr':
            pass
        else:
            # Other elements (like w:r with math only) — skip
            pass

    text = ''.join(text_parts).strip()
    return text, math_list


def detect_heading_level(para):
    """Detect heading level using OOXML-direct size + heuristics."""
    if not para.runs:
        return 0
    text = para.text.strip()
    if not text:
        return 0
    # Figure/table captions are captions, not outline headings or TOC entries.
    if re.match(r'^(图|表)\s*\d+(?:[.-]\d+)?\s*', text):
        return 0

    # ── Label-style headings (Abstract:, Key words:, 摘要：, 关键词：) ──
    # These are detected by text pattern regardless of paragraph length or OOXML formatting,
    # so that mid-length paragraphs like "Key words: memetics; ..." (~100 chars) are caught.
    # Case-insensitive for English labels so "key words:", "KEY WORDS:", etc. are all detected.
    label_patterns = [
        (r'(?i)^(Abstract\s*:?)', 2),
        (r'(?i)^(Key\s*words?\s*:?)', 2),
        (r'^(摘要\s*[：:]?)', 2),
        (r'^(关键词\s*[：:]?)', 2),
    ]
    for pat, lvl in label_patterns:
        if re.match(pat, text):
            return lvl

    # ── Explicit heading patterns should win over font heuristics ──
    # Chapter: 第1章 / 第一章 / 1 绪论 / Chapter 1
    if re.match(r'^第[一二三四五六七八九十\d]+章', text):
        return 1
    if re.match(r'^(?:Chapter\s*)?\d+\s+[\u4e00-\u9fffA-Za-z]', text) and not re.match(r'^\d+\.\d+', text):
        return 1
    if len(text) <= 80 and re.match(r'^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343]+[\u3001\uff0e.]\s*\S+', text):
        return 1
    if len(text) <= 80 and re.match(r'^\d+[\u3001\uff0e]\s*\S+', text):
        return 1
    if len(text) <= 80 and re.match(r'^[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\uff09)]\s*\S+', text):
        return 2
    # Numbered: 1.1 = level 2, 1.1.1 = level 3
    if re.match(r'^\d+\.\d+\.\d+\s*', text):
        return 3
    if re.match(r'^\d+\.\d+\s*', text):
        return 2
    # ── Numbered heading patterns for long paragraphs (>200 chars) ──
    if len(text) > 200:
        heading_patterns = [
            r'^(\d+\.\s+\w)', r'^(\d+\.\d+\s+\w)',
        ]
        for pat in heading_patterns:
            if re.match(pat, text):
                return 1 if re.match(r'^(\d+\.\s+\w)', text) else 2
        return 0

    # Skip formatting notes (Chinese parenthetical instructions or short bracketed text)
    if text.startswith('（') and (text.endswith('）') or len(text) < 60):
        return 0

    # Get size from OOXML directly (bypass python-docx API None issue)
    max_size = 0
    is_bold = False
    for r in para.runs:
        rPr = r._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        if rPr is not None:
            for child in rPr:
                tag = child.tag.split('}')[-1]
                if tag == 'sz' or tag == 'szCs':
                    try:
                        max_size = max(max_size, int(child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')) / 2.0)
                    except:
                        pass
                if tag == 'b':
                    val = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if str(val).lower() not in ('0', 'false', 'off'):
                        is_bold = True

    if not is_bold or max_size < 12:
        return 0

    # Exclude non-heading patterns
    if re.search(r'\(\d{4}[-–]\d{4}\)', text):  # year range
        return 0
    if re.search(r'\d{6}', text):  # postal code
        return 0
    if 'University' in text and len(text) > 100:  # long affiliation
        return 0
    if len(text.split()) <= 2 and not any(c.isdigit() for c in text):
        return 0
    # Formulas: contain math operators and are centered/italic
    if re.search(r'[=×÷ΣΠ∫√∞≈±]', text) and len(text) < 100:
        return 0

    # A heading must be reasonably short and not a full sentence
    if len(text) > 80:
        return 0
    # Full Chinese sentences (ending with 。) are body text, not headings
    if text.endswith('。'):
        return 0

    # Level detection
    numbered = bool(re.match(r'^[\d]+\.[\d]*\s', text))  # "1. " or "2.1 "

    if max_size >= 15:
        return 1  # 小三号及以上 = h1
    if max_size >= 14:
        return 2 if (numbered or len(text) < 60) else 1
    if max_size >= 12:
        return 2 if numbered else 3
    return 0


class ImageRegistry:
    """Per-extraction image registry.

    The same DOCX image relationship can be encountered more than once during
    verification or when Word duplicates drawing markup.  Saving by a content
    hash avoids the historical bug where one logical figure produced hundreds
    or thousands of copied files.  The registry is local to one extraction, so
    there is no university-, title-, or path-specific hardcoding.
    """

    def __init__(self, fig_dir, prefix='img'):
        self.fig_dir = fig_dir
        self.prefix = prefix
        self.counter = 0
        self.by_hash = {}
        self.failures = []

    def save_relationship_image(self, rel):
        if 'image' not in getattr(rel, 'reltype', ''):
            return None
        try:
            blob = rel.target_part.blob
            digest = hashlib.sha256(blob).hexdigest()[:20]
            if digest in self.by_hash:
                return self.by_hash[digest]

            ext = rel.target_ref.rsplit('.', 1)[-1].lower()
            if ext not in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp'):
                ext = 'png'
            self.counter += 1
            fname = f'{self.prefix}_{self.counter:03d}.{ext}'
            fpath = os.path.join(self.fig_dir, fname)
            with open(fpath, 'wb') as f:
                f.write(blob)
            self.by_hash[digest] = fname
            return fname
        except Exception as exc:
            self.failures.append({
                'target': getattr(rel, 'target_ref', ''),
                'error': str(exc)[:200],
            })
            return None


def extract_images_from_para(para, fig_dir, prefix='img', registry=None):
    """Extract inline images by rId. Returns filenames in paragraph order.

    Fixes two duplicate sources without hardcoding:
    1) `seen_rids` is paragraph-scoped, not run-scoped;
    2) image bytes are de-duplicated by SHA-256 within the extraction pass.
    """
    saved = []
    registry = registry or ImageRegistry(fig_dir, prefix)
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    seen_rids = set()
    for run in para.runs:
        xml = run._element.xml
        if 'w:drawing' not in xml and 'wp:inline' not in xml and 'wp:anchor' not in xml:
            continue
        for blip in run._element.iter(f'{{{A_NS}}}blip'):
            embed = blip.get(f'{{{R_NS}}}embed')
            if not embed or embed in seen_rids:
                continue
            seen_rids.add(embed)
            if embed in para.part.rels:
                fname = registry.save_relationship_image(para.part.rels[embed])
                if fname:
                    saved.append(fname)
            else:
                registry.failures.append({'target': embed, 'error': 'relationship id not found'})
    return saved




def _local_name(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag


def _run_text_preserve_breaks(r_elem):
    """Return visible text carried by a run, preserving explicit breaks."""
    parts = []
    for child in r_elem:
        name = _local_name(child)
        if name == 't':
            parts.append(child.text or '')
        elif name in ('tab',):
            parts.append('\t')
        elif name in ('br', 'cr'):
            parts.append('\n')
    return ''.join(parts)


def _math_entry_from_ooxml(math_elem, math_type='inline'):
    raw = etree.tounicode(math_elem, with_tail=False)
    return {'type': math_type, 'xml': raw, 'text': _math_text(math_elem)}


def _images_from_run_ooxml(run_elem, rels, registry, seen_rids, location='body'):
    """Extract images from one OOXML run in its exact paragraph position."""
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    V_NS = 'urn:schemas-microsoft-com:vml'
    out = []

    def add_rid(rid):
        if not rid or rid in seen_rids:
            return
        seen_rids.add(rid)
        if rid in rels:
            fname = registry.save_relationship_image(rels[rid])
            if fname:
                item = {'role': 'image', 'image': fname}
                if location and location != 'body':
                    item['location'] = location
                out.append(item)
        else:
            registry.failures.append({'target': rid, 'error': 'relationship id not found'})

    for blip in run_elem.iter(f'{{{A_NS}}}blip'):
        add_rid(blip.get(f'{{{R_NS}}}embed') or blip.get(f'{{{R_NS}}}link'))
    for imagedata in run_elem.iter(f'{{{V_NS}}}imagedata'):
        add_rid(imagedata.get(f'{{{R_NS}}}id') or imagedata.get(f'{{{R_NS}}}embed'))
    return out


def _image_items_from_ooxml(container_elem, rels, registry, location='body'):
    """Extract all image runs from an arbitrary OOXML container.

    Body paragraphs use `paragraph_stream_items()`, but images can also live
    inside table cells.  Keeping this helper generic lets table-cell drawings
    enter the same content image pipeline instead of disappearing silently.
    """
    W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    seen_rids = set()
    out = []
    for run_elem in container_elem.iter(f'{{{W_NS}}}r'):
        out.extend(_images_from_run_ooxml(run_elem, rels, registry, seen_rids, location=location))
    return out


def _non_body_image_entries(doc):
    """Return header/footer images that are outside the body content stream."""
    entries = []
    seen = set()
    for sec_idx, section in enumerate(doc.sections):
        for attr in (
            'header', 'first_page_header', 'even_page_header',
            'footer', 'first_page_footer', 'even_page_footer',
        ):
            try:
                part = getattr(section, attr).part
            except Exception:
                continue
            for rid, rel in getattr(part, 'rels', {}).items():
                if 'image' not in getattr(rel, 'reltype', ''):
                    continue
                try:
                    digest = hashlib.sha256(rel.target_part.blob).hexdigest()[:20]
                except Exception:
                    digest = f'{id(part)}:{rid}'
                key = (attr, digest)
                if key in seen:
                    continue
                seen.add(key)
                entries.append({
                    'location': f'section_{sec_idx + 1}_{attr}',
                    'target': getattr(rel, 'target_ref', ''),
                })
    return entries


def paragraph_stream_items(para, registry):
    """Yield paragraph text/image/math items in true OOXML run order.

    The previous extractor saved all images first and appended paragraph text
    afterwards, which could create: image -> body text -> caption.  This routine
    flushes text before every drawing or math run, then emits the token at the
    exact location where Word stores it.  It is structural and does not depend
    on a school name, figure number, or fixed paragraph index.
    """
    items = []
    buf = []
    seen_rids = set()

    def flush_text():
        text = ''.join(buf)
        buf.clear()
        if text.strip():
            items.append({'role': 'text', 'text': text})

    def append_math(math_elem, math_type='inline'):
        entry = _math_entry_from_ooxml(math_elem, math_type)
        flush_text()
        if math_type == 'display':
            items.append({
                'role': 'formula',
                'source': 'omml',
                'text': entry.get('text') or '',
                'math': [entry],
                'numbered': _formula_should_number(entry.get('text') or ''),
            })
        else:
            items.append({
                'role': 'math_inline',
                'source': 'omml',
                'text': entry.get('text') or '',
                'math': [entry],
            })

    def consume_run(run_elem):
        for part in run_elem:
            name = _local_name(part)
            if name in ('drawing', 'pict'):
                flush_text()
                items.extend(_images_from_run_ooxml(run_elem, para.part.rels, registry, seen_rids))
            elif name == 'oMath':
                append_math(part, 'inline')
            elif name == 't':
                if part.text:
                    buf.append(part.text)
            elif name in ('tab',):
                buf.append('\t')
            elif name in ('br', 'cr'):
                buf.append('\n')

    for child in para._element:
        name = _local_name(child)
        if name == 'r':
            consume_run(child)
        elif name in ('hyperlink',):
            for r in child:
                if _local_name(r) != 'r':
                    continue
                consume_run(r)
        elif name in ('oMath', 'oMathPara'):
            append_math(child, 'display' if name == 'oMathPara' else 'inline')
    flush_text()
    return items


def _append_stream_run_group(section, runs, in_appendix=False):
    if not runs:
        return
    text = ''.join(str(r.get('text') or '') for r in runs).strip()
    math_items = []
    for run in runs:
        if run.get('type') == 'math':
            math_items.extend(run.get('math') or [])
    if not math_items:
        _append_text_or_code(section, text, in_appendix=in_appendix)
        return
    non_math_text = ''.join(str(r.get('text') or '') for r in runs if r.get('type') != 'math').strip()
    if non_math_text:
        section['paragraphs'].append({
            'role': 'rich_text',
            'text': text,
            'runs': runs,
            'math': math_items,
        })
    else:
        section['paragraphs'].append({
            'role': 'formula',
            'source': 'omml',
            'text': text,
            'math': math_items,
            'numbered': _formula_should_number(text),
        })


def _caption_kind(item):
    if isinstance(item, dict):
        role = item.get('role')
        if role in ('figure_caption', 'table_caption'):
            return role
        text = item.get('text') or ''
    else:
        text = str(item or '')
    if _is_figure_caption(text):
        return 'figure_caption'
    if _is_table_caption(text):
        return 'table_caption'
    return None


def _is_image_item(item):
    return isinstance(item, dict) and item.get('role') == 'image' and item.get('image')


def _caption_text(item):
    if isinstance(item, dict):
        return item.get('text') or ''
    return str(item or '')


def pair_figure_blocks(paragraphs):
    """Pair images with captions while preserving all text.

    Besides the normal `image -> caption` layout, this fixes the two observed
    drift patterns without hardcoding figure numbers or page positions:
      * image, body text, caption  -> image+caption, body text
      * image, image, caption, caption -> image+caption, image+caption

    Look-ahead is deliberately small.  When no nearby figure caption exists the
    image token is left unchanged so content is never dropped or invented.
    """
    out = []
    i = 0
    n = len(paragraphs or [])
    while i < n:
        item = paragraphs[i]
        if not _is_image_item(item):
            out.append(item)
            i += 1
            continue

        images = []
        while i < n and _is_image_item(paragraphs[i]):
            images.append(paragraphs[i])
            i += 1

        # Look ahead through a small local window for captions.  Non-caption
        # text between image and caption is temporarily held, then emitted after
        # the paired figure so it can no longer split picture and title.
        j = i
        held_text = []
        captions = []
        max_probe = min(n, i + max(8, len(images) * 3 + 3))
        while j < max_probe and len(captions) < len(images):
            nxt = paragraphs[j]
            if _is_image_item(nxt):
                break
            kind = _caption_kind(nxt)
            if kind == 'figure_caption':
                captions.append(nxt)
                j += 1
                # Continue consuming immediately stacked figure captions for
                # the image,image,caption,caption case.
                continue
            if kind == 'table_caption':
                break
            held_text.append(nxt)
            j += 1

        if captions:
            for idx, img in enumerate(images):
                cap = captions[idx] if idx < len(captions) else None
                if cap is not None:
                    out.append({'role': 'figure', 'image': img.get('image'), 'caption': _caption_text(cap)})
                else:
                    out.append(img)
            if len(captions) > len(images):
                out.extend(captions[len(images):])
            out.extend(held_text)
            i = j
            continue

        # No nearby caption: keep all images in their original place and leave
        # subsequent text to be processed normally.
        out.extend(images)
    return out




def _normalize_role_heading(text):
    return re.sub(r'\s+', ' ', str(text or '').strip())


def _classify_section_role(heading, level=0):
    """Map a detected heading to a semantic role used by the renderer."""
    h = _normalize_role_heading(heading)
    h_compact = re.sub(r'[\s：:]+', '', h).lower()
    if h_compact in ('摘要', '中文摘要'):
        return 'cn_abstract'
    if h_compact in ('关键词', '关键字') or h.startswith('关键词'):
        return 'cn_keywords'
    if h_compact in ('abstract', 'englishabstract'):
        return 'en_abstract'
    if h.upper().replace(' ', '').startswith('KEYWORDS') or re.match(r'(?i)^key\s*words?', h):
        return 'en_keywords'
    if re.match(r'(?i)^references?$', h) or h.startswith('参考文献'):
        return 'references'
    if re.search(r'致\s*谢', h):
        return 'acknowledgement'
    if re.search(r'附\s*录', h):
        return 'appendix'
    if level and level > 0:
        return 'heading'
    return 'body'


def _split_heading_number(text):
    """Return (number, title) for headings such as '1.1标题' or '第1章绪论'."""
    t = str(text or '').strip()
    m = re.match(r'^(第[一二三四五六七八九十百千万\d]+章)\s*(.+)$', t)
    if m:
        return m.group(1), m.group(2).strip()
    m = re.match(r'^(\d+(?:\.\d+)*)\s*(.+)$', t)
    if m:
        return m.group(1), m.group(2).strip()
    return '', t


def _normalize_heading_spacing(text):
    num, title = _split_heading_number(text)
    return f'{num} {title}'.strip() if num and title else str(text or '').strip()


def _is_figure_caption(text):
    t = str(text or '').strip()
    return bool(
        re.match(r'^图\s*\d+(?:[.-]\d+)?\s*[^\d\s]', t)
        or re.match(r'(?i)^(?:fig\.?|figure)\s*\d+(?:[.-]\d+)?\s+[^\d\s]', t)
    )


def _is_table_caption(text):
    t = str(text or '').strip()
    return bool(
        re.match(r'^表\s*\d+(?:[.-]\d+)?\s*[^\d\s]', t)
        or re.match(r'(?i)^table\s*\d+(?:[.-]\d+)?\s+[^\d\s]', t)
    )


def _normalize_caption_spacing(text):
    t = str(text or '').strip()
    t = re.sub(r'(?i)^(fig\.?|figure)\s*(\d+(?:[.-]\d+)?)\s*', r'Fig. \2 ', t)
    t = re.sub(r'(?i)^table\s*(\d+(?:[.-]\d+)?)\s*', r'Table \1 ', t)
    return re.sub(r'^(图|表)\s*(\d+(?:[.-]\d+)?)\s*', r'\1 \2 ', t).strip()


def _clean_markdown_links(text):
    def repl(m):
        label = (m.group(1) or '').strip()
        target = (m.group(2) or '').strip()
        return label or target
    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', repl, str(text or ''))


def _clean_text_artifacts(text, preserve_newlines=False):
    """Remove generic editor/clipboard artifacts without changing content semantics."""
    t = _clean_markdown_links(text)
    t = t.replace('\u00a0', ' ')
    if preserve_newlines:
        lines = []
        for line in t.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
            s = re.sub(r'[ \t]+', ' ', line).strip()
            if _is_noise_text(s):
                continue
            lines.append(s)
        return '\n'.join(lines).strip()
    t = re.sub(r'\s+', ' ', t).strip()
    return '' if _is_noise_text(t) else t


def _is_noise_text(text):
    t = str(text or '').strip()
    return t in {'复制', 'Copy', 'Plain Text', '纯文本'}


def _clean_code_text(text):
    return _clean_text_artifacts(text, preserve_newlines=True)


def _clean_formula_text(text):
    t = _clean_text_artifacts(text)
    if t.count('|') >= 3:
        t = t.replace('|', '')
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def _paragraph_plain_text_from_ooxml(p_elem):
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    pieces = []
    for r in p_elem.findall(f'{{{W}}}r'):
        part = ''.join(t.text or '' for t in r.findall(f'{{{W}}}t'))
        if r.find(f'{{{W}}}br') is not None and not part:
            part = '\n'
        pieces.append(part)
    return ''.join(pieces)


def _extract_table_rows_from_ooxml(tbl_elem):
    """Preserve cell paragraph breaks so code/config tables do not collapse into one line."""
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    rows = []
    for tr in tbl_elem.findall(f'{{{W}}}tr'):
        cells = []
        for tc in tr.findall(f'{{{W}}}tc'):
            paras = []
            for p in tc.findall(f'{{{W}}}p'):
                txt = _clean_text_artifacts(_paragraph_plain_text_from_ooxml(p), preserve_newlines=True).rstrip()
                if txt:
                    paras.append(txt)
            cells.append('\n'.join(paras).strip())
        rows.append(cells)
    return rows

def _looks_like_code_line(text):
    """Heuristic for network/device configuration or command-line code.

    Kept generic: it detects syntax-like command lines rather than any
    particular vendor, university, or template.
    """
    t = (text or '').strip()
    if not t or len(t) > 220:
        return False
    if re.match(r'^[A-Za-z0-9_.-]+[>#]', t):
        return True
    if re.match(r'^(interface|vlan|ip route|ip address|router|switchport|acl|rule|nat|dhcp|dns|ospf|bgp|display|show|ping|tracert|undo|quit|return|sysname|description|gateway|firewall|security-policy)\b', t, re.I):
        return True
    if re.match(r'^[a-z][a-z0-9_-]+\s+[-A-Za-z0-9_/.:]+', t) and any(ch in t for ch in ['/', '.', '-', '_']):
        return True
    return False



def _table_rows_look_like_code(rows):
    """Classify one-/two-column command tables as code, not academic tables."""
    flat = []
    for row in rows or []:
        for cell in row or []:
            for line in str(cell or '').splitlines():
                if line.strip():
                    flat.append(line.strip())
    if not flat:
        return False
    ncols = max((len(r) for r in rows or []), default=0)
    hits = sum(1 for x in flat if _looks_like_code_line(x))
    if ncols <= 1 and len(flat) >= 2 and hits >= 2:
        return True
    if ncols <= 2 and len(flat) >= 4 and hits >= max(2, len(flat) // 3):
        return True
    return False


def _code_text_from_table_rows(rows):
    lines = []
    for row in rows or []:
        cells = [str(c or '').rstrip() for c in row]
        if len(cells) == 1:
            lines.append(cells[0])
        else:
            lines.append('    '.join(cells).rstrip())
    return _clean_code_text('\n'.join(lines).rstrip())


def _looks_like_formula_text(text):
    """Detect standalone calculation/formula paragraphs.

    The rule is intentionally structural: formulas are short standalone lines
    with equality/calculation operators. Some thesis sources store formulas as
    plain text, including definition lines without numbers and continuation
    lines that start with "=".
    """
    t = str(text or '').strip()
    if not t or len(t) > 180:
        return False
    if t.endswith(('。', '！', '？', '；', ';')):
        return False
    if re.search(r'^\[\d+\]', t) or re.search(r'\[\d+(?:[-,，、]\d+)*\]', t):
        return False
    if _latex_from_formula_text(t):
        return True
    has_equal = bool(re.search(r'[=＝≈≤≥<>]', t))
    has_operator = bool(re.search(r'[+\-*/×÷%]', t))
    has_digit = bool(re.search(r'\d', t))
    starts_continuation = bool(re.match(r'^[=＝≈≤≥<>]', t)) and has_digit
    if starts_continuation:
        return True
    if has_equal and has_operator:
        return True
    return False


def _latex_escape_text(text):
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def _latex_from_formula_text(text):
    t = str(text or '').strip()
    if t.startswith('$$') and t.endswith('$$'):
        return t[2:-2].strip()
    if t.startswith('$') and t.endswith('$'):
        return t[1:-1].strip()
    return ''


def _formula_should_number(text):
    t = str(text or '').strip()
    if not re.search(r'\d', t):
        return False
    has_equal = bool(re.search(r'[=＝≈≤≥<>]', t))
    has_operator = bool(re.search(r'[+*/×÷%]|\s-\s', t))
    return bool(has_equal and has_operator)


def _formula_item_from_text(text):
    clean = _clean_formula_text(text)
    item = {'role': 'formula', 'source': 'text', 'text': clean, 'numbered': _formula_should_number(clean)}
    latex = _latex_from_formula_text(clean)
    if latex:
        item['source'] = 'latex'
        item['latex'] = latex
    return item

def _append_text_or_code(section, text, in_appendix=False):
    """Append semantic blocks while preserving captions, code and inline citations."""
    if not text:
        return
    text = _clean_text_artifacts(text)
    if not text:
        return
    if _is_figure_caption(text):
        section['paragraphs'].append({'role': 'figure_caption', 'text': _normalize_caption_spacing(text)})
    elif _is_table_caption(text):
        section['paragraphs'].append({'role': 'table_caption', 'text': _normalize_caption_spacing(text)})
    elif _looks_like_formula_text(text):
        section['paragraphs'].append(_formula_item_from_text(text))
    elif in_appendix and (_looks_like_code_line(text) or '\n' in text):
        section['paragraphs'].append({'role': 'code', 'code': _clean_code_text(text)})
    else:
        section['paragraphs'].append(text)

def extract(docx_path, output_dir='Inputs'):
    """Extract content from a content docx into structured JSON + copy images."""
    doc = Document(docx_path)
    base = os.path.splitext(os.path.basename(docx_path))[0]

    # Setup output dirs.  Recreate figures for each extraction so repeated
    # verification passes do not accumulate stale/duplicated files.
    content_dir = os.path.join(output_dir, base)
    fig_dir = os.path.join(content_dir, 'figures')
    shutil.rmtree(fig_dir, ignore_errors=True)
    os.makedirs(fig_dir, exist_ok=True)
    image_registry = ImageRegistry(fig_dir, f'{base}_img')

    content = {
        '_meta': {
            'source': os.path.basename(docx_path),
            'sha256': hashlib.sha256(open(docx_path, 'rb').read()).hexdigest()[:16],
            'paragraphs': len(doc.paragraphs),
            'tables_count': len(doc.tables),
        },
        'title_info': {},
        'sections': [],
        'references': [],
    }

    # ── Extract title info ──
    # Find the largest-text paragraph in the first 20 paragraphs
    text_start = 0
    best_title = ('', 0, 0)  # (text, size, index)
    for i, p in enumerate(doc.paragraphs[:20]):
        txt = p.text.strip()
        if not txt or len(txt) < 10:
            continue
        if txt.startswith('（') and txt.endswith('）'):
            continue
        # Get max font size from runs
        max_sz = 0
        for r in p.runs:
            if r.font.size:
                max_sz = max(max_sz, r.font.size.pt)
        if max_sz >= 14 and max_sz > best_title[1]:
            clean = txt.split('（')[0].strip().replace('\n', ' ').replace('\r', '')
            if clean and not clean.startswith('年') and not clean.startswith('本科'):
                best_title = (clean, max_sz, i)

    if best_title[0]:
        content['title_info']['title_cn'] = best_title[0]
        text_start = best_title[2] + 1

    # ── Extract cover info from content docx tables ──
    cover_info = {}
    _COVER_LABEL_MAP = {
        '学校编码': 'school_code', '学位编码': 'degree_code', '论文题目': 'paper_title',
        '学生姓名': 'student_name', '学号': 'student_id', '学    号': 'student_id',
        '所属学院': 'college', '专业班级': 'class_name',
        '指导老师': 'advisor', '指导教师': 'advisor',
    }
    for table in doc.tables[:5]:
        for row in table.rows:
            if len(row.cells) >= 2:
                label = row.cells[0].text.strip()
                value = row.cells[1].text.strip()
                if label and value:
                    for kw, key in _COVER_LABEL_MAP.items():
                        if kw in label:
                            content['cover_info'] = content.get('cover_info', {})
                            content['cover_info'][key] = value
                            cover_info[key] = value
                            break
    # Cover table is the most reliable source for the paper title.
    # The old heuristic sometimes picked declaration/footer text with large font.
    if cover_info.get('paper_title'):
        content['title_info']['title_cn'] = cover_info['paper_title']

    # ── Parse sections ──
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    current_section = {'heading': '正文', 'level': 1, 'role': 'body', 'paragraphs': [], 'images': []}
    sections = [current_section]
    ref_section = None

    # Skip body elements before text_start
    _body_children = list(doc.element.body)
    _p_idx = 0  # paragraph index in body children
    _started = False
    for _child in _body_children:
        _tag = _child.tag.split('}')[-1]

        if _tag == 'p':
            if _p_idx < text_start:
                _p_idx += 1
                continue
            _started = True
            p = doc.paragraphs[_p_idx]
            _p_idx += 1
            text = p.text.strip()

            level = detect_heading_level(p)

            if re.match(r'(?i)^references?\b', text) or text.startswith('参考文献'):
                ref_section = {'heading': text, 'entries': []}
                continue

            # Back matter after references (致谢/附录) must not be swallowed
            # by the reference collector. Treat it as normal sections again.
            if ref_section is not None and level > 0 and re.search(r'(致\s*谢|附\s*录)', text):
                _h = _normalize_heading_spacing(text.split('（')[0].strip())
                current_section = {'heading': _h, 'level': level, 'role': _classify_section_role(_h, level), 'paragraphs': [], 'images': []}
                sections.append(current_section)
                ref_section = None
                continue

            if ref_section is not None:
                clean_ref_text = _clean_text_artifacts(text)
                if clean_ref_text:
                    ref_section['entries'].append(clean_ref_text)
            elif level > 0:
                clean_heading = text.split('（')[0].strip()
                if not clean_heading:
                    continue
                m = re.match(r'(?i)^(Abstract\s*:?|Key\s*words?\s*:?|摘要\s*[：:]|关键词\s*[：:])\s*', text)
                if m:
                    heading_part = m.group(1).strip()
                    body_part = text[m.end():].strip()
                    body_part = re.sub(r'^[（(][^）)]*[）)]\s*', '', body_part)
                else:
                    heading_part = clean_heading
                    body_part = ''
                heading_part = _normalize_heading_spacing(heading_part)
                role = _classify_section_role(heading_part, level)
                current_section = {'heading': heading_part, 'level': level, 'role': role, 'paragraphs': [], 'images': []}
                if role == 'en_abstract':
                    current_section['page_break_before'] = True
                sections.append(current_section)
                if body_part:
                    _append_text_or_code(current_section, body_part, in_appendix=False)
            else:
                # Preserve exact OOXML run order within the paragraph.  This is
                # crucial when Word stores explanatory text, a drawing, and a
                # caption in nearby runs/paragraphs.
                stream_items = paragraph_stream_items(p, image_registry)
                if not stream_items and text:
                    stream_items = [{'role': 'text', 'text': text}]
                rich_runs = []
                in_appendix = bool(re.search(r'(附\s*录|配置|命令|代码)', current_section.get('heading','')))

                def _flush_rich_runs():
                    nonlocal rich_runs
                    _append_stream_run_group(current_section, rich_runs, in_appendix=in_appendix)
                    rich_runs = []

                for _it in stream_items:
                    if _it.get('role') == 'image':
                        _flush_rich_runs()
                        current_section['images'].append(_it.get('image'))
                        current_section['paragraphs'].append(_it)
                    elif _it.get('role') == 'math_inline':
                        rich_runs.append({'type': 'math', 'text': _it.get('text') or '', 'math': _it.get('math') or []})
                    elif _it.get('role') == 'formula':
                        _flush_rich_runs()
                        current_section['paragraphs'].append(_it)
                    elif _it.get('role') == 'text':
                        txt = _it.get('text') or ''
                        if txt:
                            rich_runs.append({'type': 'text', 'text': txt})
                _flush_rich_runs()

        elif _tag == 'tbl' and _started:
            # Body table — preserve paragraph breaks inside each cell.
            _rows = _extract_table_rows_from_ooxml(_child)
            _table_images = _image_items_from_ooxml(_child, doc.part.rels, image_registry, location='table_cell')
            if _rows:
                if ref_section is not None:
                    if _table_rows_look_like_code(_rows):
                        ref_section['entries'].append({'role': 'code', 'code': _code_text_from_table_rows(_rows), 'table_rows': _rows})
                    else:
                        ref_section['entries'].append({'role': 'table', 'table_rows': _rows})
                elif _table_rows_look_like_code(_rows):
                    current_section['paragraphs'].append({'role': 'code', 'code': _code_text_from_table_rows(_rows), 'table_rows': _rows})
                else:
                    current_section['paragraphs'].append({'role': 'table', 'table_rows': _rows})
            if _table_images and ref_section is None:
                for _img in _table_images:
                    current_section['images'].append(_img.get('image'))
                    current_section['paragraphs'].append(_img)

    if ref_section and ref_section['entries']:
        content['references'] = ref_section['entries']

    # Filter empty placeholder sections, but KEEP structural headings.
    # A level-1 chapter can legitimately have no direct paragraphs because it is
    # followed immediately by 1.1/1.2 subsections. Dropping it breaks TOC and body
    # structure, so keep every non-placeholder heading even when empty.
    content['sections'] = []
    for s in sections:
        # If real headings were detected, the initial placeholder section is
        # almost always cover/declaration/title residue. Keep it only when it
        # is the sole section in an unstructured document.
        if s['heading'] == '正文' and len(sections) > 1:
            continue
        if s['paragraphs'] or s['images'] or (s.get('heading') and s.get('heading') != '正文'):
            content['sections'].append(s)
    if ref_section:
        content['references'] = ref_section['entries']

    # Mark the first real body chapter so renderers can start it on a new page.
    _front_roles = {'cn_abstract', 'cn_keywords', 'en_abstract', 'en_keywords'}
    for _s in content['sections']:
        if _s.get('role') not in _front_roles and not _s.get('page_break_before'):
            _s.setdefault('page_break_before', True)
            break

    # Pair images with following figure captions after all sections are built.
    for _s in content['sections']:
        _s['paragraphs'] = pair_figure_blocks(_s.get('paragraphs') or [])

    # Count saved images without running a second extraction pass.
    # Re-extracting here used to create duplicate filenames and made figure
    # captions drift away from their intended images.
    content['_meta']['images_extracted'] = len([
        f for f in os.listdir(fig_dir)
        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff'))
    ])
    content['_meta']['images_dir'] = os.path.abspath(fig_dir)
    content['_meta']['image_extract_failures'] = image_registry.failures
    content['_meta']['non_body_images'] = _non_body_image_entries(doc)

    return content


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Templates/模版.docx'
    content = extract(path)
    json_path = os.path.splitext(path)[0] + '_content.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    print(f'Content JSON → {json_path}')
    print(f'Sections: {len(content["sections"])}  References: {len(content["references"])}  Images: {content["_meta"]["images_extracted"]}')

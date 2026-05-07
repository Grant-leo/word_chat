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

    text = ''.join(text_parts).strip() if text_parts else para.text.strip()
    return text, math_list


def detect_heading_level(para):
    """Detect heading level using OOXML-direct size + heuristics."""
    if not para.runs:
        return 0
    text = para.text.strip()
    if not text:
        return 0

    # ── Label-style headings (Abstract:, Key words:, 摘要：, 关键词：) ──
    # These are detected by text pattern regardless of paragraph length or OOXML formatting,
    # so that mid-length paragraphs like "Key words: memetics; ..." (~100 chars) are caught.
    # Case-insensitive for English labels so "key words:", "KEY WORDS:", etc. are all detected.
    label_patterns = [
        (r'(?i)^(Abstract\s*:?)', 2),
        (r'(?i)^(Key\s*words?\s*:?)', 2),
        (r'^(摘要\s*[：:])', 2),
        (r'^(关键词\s*[：:])', 2),
    ]
    for pat, lvl in label_patterns:
        if re.match(pat, text):
            return lvl

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


def extract_images_from_para(para, fig_dir, prefix='img'):
    """Extract inline images, save with plain-ASCII safe names. Returns list of filenames."""
    saved = []
    for run in para.runs:
        xml = run._element.xml
        if 'w:drawing' not in xml and 'wp:inline' not in xml:
            continue
        for rel in para.part.rels.values():
            if 'image' not in rel.reltype:
                continue
            try:
                ext = rel.target_ref.rsplit('.', 1)[-1]
                if ext.lower() not in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff'):
                    ext = 'png'
                seq = len([f for f in os.listdir(fig_dir) if f.startswith(prefix)]) + 1
                fname = f'{prefix}_{seq:03d}.{ext}'
                fpath = os.path.join(fig_dir, fname)
                if not os.path.exists(fpath):
                    with open(fpath, 'wb') as f:
                        f.write(rel.target_part.blob)
                saved.append(fname)
            except:
                pass
    return saved


def extract(docx_path, output_dir='Inputs'):
    """Extract content from a content docx into structured JSON + copy images."""
    doc = Document(docx_path)
    base = os.path.splitext(os.path.basename(docx_path))[0]

    # Setup output dirs
    content_dir = os.path.join(output_dir, base)
    fig_dir = os.path.join(content_dir, 'figures')
    os.makedirs(fig_dir, exist_ok=True)

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

    # ── Parse sections ──
    current_section = {'heading': '正文', 'level': 1, 'paragraphs': [], 'images': []}
    sections = [current_section]
    ref_section = None

    for p in doc.paragraphs[text_start:]:
        text = p.text.strip()
        level = detect_heading_level(p)

        # Detect references section
        if re.match(r'(?i)^references?\b', text) or text.startswith('参考文献'):
            ref_section = {'heading': text, 'entries': []}
            continue

        if level > 0:
            # Clean heading text: remove inline formatting notes
            clean_heading = text.split('（')[0].strip()
            if not clean_heading:
                continue
            # Check for label-style headings (Abstract:, Key words:, 摘要：, 关键词：)
            # that may have inline body text — split at any paragraph length
            m = re.match(r'(?i)^(Abstract\s*:?|Key\s*words?\s*:?|摘要\s*[：:]|关键词\s*[：:])\s*', text)
            if m:
                heading_part = m.group(1).strip()
                body_part = text[m.end():].strip()
                # Remove leading formatting notes
                body_part = re.sub(r'^[（(][^）)]*[）)]\s*', '', body_part)
            else:
                heading_part = clean_heading
                body_part = ''
            current_section = {'heading': heading_part, 'level': level, 'paragraphs': [], 'images': []}
            sections.append(current_section)
            if body_part:
                current_section['paragraphs'].append(body_part)
        elif ref_section is not None:
            # In references section
            if text:
                ref_section['entries'].append(text)
        else:
            # Body paragraph — check for math formulas
            imgs = extract_images_from_para(p, fig_dir, f'{base}_img')
            current_section['images'].extend(imgs)
            clean_text, math_list = extract_math(p)
            if math_list:
                current_section['paragraphs'].append({
                    'text': clean_text,
                    'math': math_list,
                })
            elif clean_text:
                current_section['paragraphs'].append(clean_text)

    if ref_section and ref_section['entries']:
        content['references'] = ref_section['entries']

    # Filter empty sections and placeholder "正文"
    content['sections'] = [s for s in sections
                           if (s['paragraphs'] or s['images'])
                           and (s['heading'] != '正文' or len(sections) == 1)]
    if ref_section:
        content['references'] = ref_section['entries']

    # Extract all images from entire document (including those not in body text)
    all_imgs = []
    for p in doc.paragraphs:
        all_imgs.extend(extract_images_from_para(p, fig_dir, f'{base}_img'))

    # Also extract from tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    all_imgs.extend(extract_images_from_para(p, fig_dir, f'{base}_tbl_img'))

    content['_meta']['images_extracted'] = len(all_imgs)
    content['_meta']['images_dir'] = fig_dir

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

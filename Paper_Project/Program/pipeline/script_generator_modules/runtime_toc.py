"""TOC and page-resolution runtime template fragment for generated build scripts."""
from __future__ import annotations

TOC_RUNTIME = r'''
def enable_update_fields_on_open():
    settings = doc.settings._element
    upd = settings.find(qn('w:updateFields'))
    if upd is None:
        upd = OxmlElement('w:updateFields')
        settings.append(upd)
    upd.set(qn('w:val'), 'true')



def collect_toc_entries():
    entries = []
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if not h or h == '正文':
            continue
        if is_reference_heading(h) or is_backmatter_heading(h) or is_caption_heading(h) or role in ('references', 'acknowledgement', 'appendix'):
            continue
        level = max(1, min(int(sec.get('level') or 1), 3))
        entries.append({'level': level, 'text': normalize_heading_spacing(h)})
    if DATA.get('references'):
        entries.append({'level': 1, 'text': '参考文献'})
    ack_sections, app_sections = collect_structural_backmatter()
    pure_refs, ack_from_refs, app_from_refs = split_refs_backmatter(DATA.get('references') or [])
    if ack_sections or ack_from_refs:
        entries.append({'level': 1, 'text': '致  谢'})
    if app_sections or app_from_refs:
        entries.append({'level': 1, 'text': '附  录'})
    return entries


def add_toc_line(text, level, page_text=''):
    level = max(1, min(int(level or 1), 3))
    prof = profile('toc' + str(level))
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=0)
    indents = (DATA.get('rules') or {}).get('toc_indents_cm') or [0.0, 0.74, 1.48]
    p.paragraph_format.left_indent = Cm(float(indents[level - 1] if len(indents) >= level else 0.0))
    tabs = p.paragraph_format.tab_stops
    tabs.add_tab_stop(Cm(max(1.0, text_width_cm() - 0.15)), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    r = p.add_run(text)
    apply_run_profile(r, prof, text)
    p.add_run('\t')
    r2 = p.add_run(str(page_text))
    page_prof = dict(profile('toc2'))
    apply_run_profile(r2, page_prof, str(page_text))
    return p


def add_wps_toc_field():
    """Insert a real Word/WPS TOC field instead of a fake static directory.

    The generated DOCX only needs correct heading styles/outline levels.  WPS
    can then populate page numbers by Update Field / Generate Directory.  This
    avoids any dependency on LibreOffice or pdftotext and does not guess pages.
    """
    p = doc.add_paragraph()
    apply_paragraph_profile(p, profile('body'), first_indent_override=0)

    def append_run_with(el):
        r = OxmlElement('w:r')
        r.append(el)
        p._element.append(r)
        return r

    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    append_run_with(begin)

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = r' TOC \o "1-3" \h \z \u '
    append_run_with(instr)

    sep = OxmlElement('w:fldChar')
    sep.set(qn('w:fldCharType'), 'separate')
    append_run_with(sep)

    hint = p.add_run('请在 WPS 中右键“更新域”或“生成目录”，目录将按正文标题自动生成。')
    hint_prof = dict(profile('body'))
    hint_prof['italic'] = True
    apply_run_profile(hint, hint_prof, hint.text)

    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    append_run_with(end)


def add_toc():
    enable_update_fields_on_open()
    configure_global_styles()
    add_text('目  录', role='toc_title', first_indent=False)
    if USE_NATIVE_TOC:
        add_wps_toc_field()
        return
    entries = collect_toc_entries()
    for ent in entries:
        key = _norm_for_pdf_match(ent.get('text') or '')
        add_toc_line(ent.get('text') or '', ent.get('level') or 1, TOC_PAGE_MAP.get(key, ''))



def _norm_for_pdf_match(text):
    return re.sub(r'\s+', '', str(text or '')).lower()


def _extract_pdf_pages(pdf_path):
    exe = shutil.which('pdftotext')
    if not exe:
        return []
    with tempfile.TemporaryDirectory() as td:
        txt = os.path.join(td, 'out.txt')
        cmd = [exe, '-layout', pdf_path, txt]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
        if not os.path.exists(txt):
            return []
        data = open(txt, 'r', encoding='utf-8', errors='ignore').read()
    return data.split('\f')


def _make_pdf_for_pagination(docx_path):
    soffice = shutil.which('libreoffice') or shutil.which('soffice')
    if not soffice:
        for p in [r'C:\Program Files\LibreOffice\program\soffice.exe',
                  r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']:
            if os.path.exists(p):
                soffice = p
                break
    if not soffice:
        return None
    td = tempfile.mkdtemp(prefix='toc_pages_')
    profile = os.path.join(td, 'lo_profile')
    home = os.path.join(td, 'home')
    os.makedirs(profile, exist_ok=True)
    os.makedirs(home, exist_ok=True)
    try:
        from pathlib import Path
        profile_uri = Path(profile).as_uri()
    except Exception:
        profile_uri = 'file://' + profile.replace(' ', '%20')
    cmd = [
        soffice, '--headless', '--norestore', '--nofirststartwizard',
        f'-env:UserInstallation={profile_uri}',
        '--convert-to', 'pdf', '--outdir', td, docx_path,
    ]
    env = dict(os.environ)
    env['HOME'] = home
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120, env=env)
    base = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf = os.path.join(td, base)
    return pdf if os.path.exists(pdf) and os.path.getsize(pdf) > 0 else None


def _infer_heading_pages_from_pdf(docx_path=None):
    """Render the current DOCX to PDF and infer static TOC page numbers.

    This is the no-hardcode TOC pass: page numbers are not guessed from chapter
    count, school name, or fixed page offsets.  The first rendered body heading
    defines Arabic page 1, then every TOC entry is located in the rendered PDF
    text after the TOC pages.
    """
    try:
        pdf = _make_pdf_for_pagination(docx_path or OUT)
        if not pdf:
            return {}
        pages = _extract_pdf_pages(pdf)
        if not pages:
            return {}
        norm_pages = [_norm_for_pdf_match(p) for p in pages]
        entries = collect_toc_entries()
        if not entries:
            return {}

        first = _norm_for_pdf_match(entries[0]['text'])
        toc_last = 0
        for i, text in enumerate(norm_pages):
            if _norm_for_pdf_match('目录') in text:
                toc_last = i

        body_start = None
        for i in range(toc_last + 1, len(norm_pages)):
            if first and first in norm_pages[i]:
                body_start = i
                break
        if body_start is None:
            return {}

        page_map = {}
        search_from = body_start
        for ent in entries:
            key = _norm_for_pdf_match(ent['text'])
            if not key:
                continue
            found = None
            for i in range(search_from, len(norm_pages)):
                if key in norm_pages[i]:
                    found = i
                    break
            if found is None:
                # Some PDF extractors insert or drop punctuation/spaces.  Fall
                # back to a looser key built from the first substantial token.
                loose = re.sub(r'[^0-9a-zA-Z\u4e00-\u9fff]+', '', key)[:16]
                if loose:
                    for i in range(search_from, len(norm_pages)):
                        if loose in norm_pages[i]:
                            found = i
                            break
            if found is not None:
                page_map[key] = found - body_start + 1
                search_from = min(found, len(norm_pages) - 1)
        return page_map
    except Exception:
        return {}


def _rewrite_static_toc_pages(page_map):
    if not page_map:
        return False
    try:
        d = Document(OUT)
        in_toc = False
        changed = False
        for p in d.paragraphs:
            txt = p.text.strip()
            if txt == '目录':
                in_toc = True
                continue
            if not in_toc:
                continue
            pPr = p._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                break
            if '\t' not in p.text:
                continue
            label = p.text.split('\t', 1)[0].strip()
            key = _norm_for_pdf_match(label)
            if key not in page_map:
                continue
            # Preserve paragraph formatting; rebuild simple runs only.
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            prof = profile('body')
            r1 = p.add_run(label)
            apply_run_profile(r1, prof, label)
            p.add_run('\t')
            r2 = p.add_run(str(page_map[key]))
            apply_run_profile(r2, prof, str(page_map[key]))
            changed = True
        if changed:
            d.save(OUT)
        return changed
    except Exception:
        return False


def update_static_toc_pages():
    page_map = _infer_heading_pages_from_pdf()
    _rewrite_static_toc_pages(page_map)


def _infer_heading_pages_from_word_com(docx_path=None):
    """Use Word pagination to compute TOC page numbers without updating fields."""
    try:
        import win32com.client  # type: ignore
    except Exception:
        return {}
    word = None
    doc_obj = None
    try:
        entries = collect_toc_entries()
        if not entries:
            return {}
        entry_keys = [_norm_for_pdf_match(e.get('text') or '') for e in entries]
        wanted = set(k for k in entry_keys if k)
        found = {}
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        doc_obj = word.Documents.Open(os.path.abspath(docx_path or OUT), ReadOnly=True)
        try:
            doc_obj.Repaginate()
        except Exception:
            pass
        for para in doc_obj.Paragraphs:
            try:
                if int(para.OutlineLevel) not in (1, 2, 3):
                    continue
            except Exception:
                continue
            text = str(para.Range.Text or '').replace('\r', '').replace('\x07', '').strip()
            key = _norm_for_pdf_match(text)
            if key in wanted and key not in found:
                try:
                    found[key] = int(para.Range.Information(3))
                except Exception:
                    pass
            if len(found) >= len(wanted):
                break
        if not found:
            return {}
        first_key = next((k for k in entry_keys if k in found), None)
        if not first_key:
            return {}
        first_page = found[first_key]
        return {k: max(1, v - first_page + 1) for k, v in found.items()}
    except Exception:
        return {}
    finally:
        try:
            if doc_obj is not None:
                doc_obj.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def update_fields_with_word_com():
    """Update TOC/fields through Microsoft Word when available on Windows."""
    try:
        import win32com.client  # type: ignore
    except Exception:
        return False
    word = None
    doc_obj = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        doc_obj = word.Documents.Open(os.path.abspath(OUT), ReadOnly=False)
        try:
            for toc in doc_obj.TablesOfContents:
                toc.Update()
        except Exception:
            pass
        try:
            doc_obj.Fields.Update()
        except Exception:
            pass
        doc_obj.Save()
        return True
    except Exception:
        return False
    finally:
        try:
            if doc_obj is not None:
                doc_obj.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


'''

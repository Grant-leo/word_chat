"""OOXML Math builder helpers for the LaTeX converter."""
from __future__ import annotations

from lxml import etree

try:
    from latex_omath_modules.symbols import M, XML_SPACE, _MATRIX_BRACKETS
except ImportError:  # pragma: no cover - package-style imports
    from .symbols import M, XML_SPACE, _MATRIX_BRACKETS


def _is_math_identifier_text(text):
    """True for variable-like literal letters that should keep math italic."""
    if not text:
        return False
    for ch in str(text):
        if ("A" <= ch <= "Z") or ("a" <= ch <= "z") or ("\u0370" <= ch <= "\u03ff"):
            continue
        return False
    return True


def _make_literal_run(text):
    """Create a run for literal LaTeX chars.

    Raw identifiers keep Word's default math style; punctuation, digits,
    brackets, and operators are forced upright for Word/WPS consistency.
    """
    style = None if _is_math_identifier_text(text) else "plain"
    return _make_run(text, style=style)


def _make_run(text, style=None):
    """Create m:r element with text content."""
    r = etree.Element(f"{{{M}}}r")
    rPr = etree.SubElement(r, f"{{{M}}}rPr")
    sty_map = {
        "plain": "p", "bold": "b", "italic": "i",
        "bold-italic": "bi", "script": "scr",
    }
    if style and style in sty_map:
        sty = etree.SubElement(rPr, f"{{{M}}}sty")
        sty.set(f"{{{M}}}val", sty_map[style])
    elif style and style not in sty_map:
        nor = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        w_rPr = etree.SubElement(rPr, f"{nor}rPr")
        if style == "sans":
            rf = etree.SubElement(w_rPr, f"{nor}rFonts")
            rf.set(f"{nor}ascii", "Arial")
        elif style == "mono":
            rf = etree.SubElement(w_rPr, f"{nor}rFonts")
            rf.set(f"{nor}ascii", "Consolas")
    t = etree.SubElement(r, f"{{{M}}}t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    return r


def _make_fraction(num_el, den_el, frac_type="bar"):
    """Create m:f element. frac_type: bar, nobar (binomial), lin, skw."""
    f = etree.Element(f"{{{M}}}f")
    fPr = etree.SubElement(f, f"{{{M}}}fPr")
    typ = etree.SubElement(fPr, f"{{{M}}}type")
    typ.set(f"{{{M}}}val", frac_type)
    num = etree.SubElement(f, f"{{{M}}}num")
    num.append(num_el)
    den = etree.SubElement(f, f"{{{M}}}den")
    den.append(den_el)
    return f


def _make_sup(base_el, sup_el):
    s = etree.Element(f"{{{M}}}sSup")
    etree.SubElement(s, f"{{{M}}}sSupPr")
    e = etree.SubElement(s, f"{{{M}}}e")
    e.append(base_el)
    sp = etree.SubElement(s, f"{{{M}}}sup")
    sp.append(sup_el)
    return s


def _make_sub(base_el, sub_el):
    s = etree.Element(f"{{{M}}}sSub")
    etree.SubElement(s, f"{{{M}}}sSubPr")
    e = etree.SubElement(s, f"{{{M}}}e")
    e.append(base_el)
    sb = etree.SubElement(s, f"{{{M}}}sub")
    sb.append(sub_el)
    return s


def _make_supsub(base_el, sub_el, sup_el):
    s = etree.Element(f"{{{M}}}sSubSup")
    etree.SubElement(s, f"{{{M}}}sSubSupPr")
    e = etree.SubElement(s, f"{{{M}}}e")
    e.append(base_el)
    sb = etree.SubElement(s, f"{{{M}}}sub")
    sb.append(sub_el)
    sp = etree.SubElement(s, f"{{{M}}}sup")
    sp.append(sup_el)
    return s


def _make_nary(chr_char, sub_el, sup_el, e_el, limloc="undOvr"):
    n = etree.Element(f"{{{M}}}nary")
    nPr = etree.SubElement(n, f"{{{M}}}naryPr")
    c = etree.SubElement(nPr, f"{{{M}}}chr")
    c.set(f"{{{M}}}val", chr_char)
    g = etree.SubElement(nPr, f"{{{M}}}grow")
    g.set(f"{{{M}}}val", "1")
    ll = etree.SubElement(nPr, f"{{{M}}}limLoc")
    ll.set(f"{{{M}}}val", limloc)
    if sub_el is not None:
        sb = etree.SubElement(n, f"{{{M}}}sub")
        sb.append(sub_el)
    if sup_el is not None:
        sp = etree.SubElement(n, f"{{{M}}}sup")
        sp.append(sup_el)
    e = etree.SubElement(n, f"{{{M}}}e")
    e.append(e_el)
    return n


def _make_radical(deg_el, e_el):
    rad = etree.Element(f"{{{M}}}rad")
    etree.SubElement(rad, f"{{{M}}}radPr")
    if deg_el is not None:
        d = etree.SubElement(rad, f"{{{M}}}deg")
        d.append(deg_el)
    e = etree.SubElement(rad, f"{{{M}}}e")
    e.append(e_el)
    return rad


def _make_delimiter(left_chr, right_chr, content_el, grow=True):
    d = etree.Element(f"{{{M}}}d")
    dPr = etree.SubElement(d, f"{{{M}}}dPr")
    beg = etree.SubElement(dPr, f"{{{M}}}begChr")
    beg.set(f"{{{M}}}val", left_chr or "")
    sep = etree.SubElement(dPr, f"{{{M}}}sepChr")
    sep.set(f"{{{M}}}val", "")
    end = etree.SubElement(dPr, f"{{{M}}}endChr")
    end.set(f"{{{M}}}val", right_chr or "")
    if grow:
        gr = etree.SubElement(dPr, f"{{{M}}}grow")
        gr.set(f"{{{M}}}val", "1")
    items = content_el if isinstance(content_el, list) else [content_el]
    for item in items:
        if item is not None:
            e = etree.SubElement(d, f"{{{M}}}e")
            e.append(item)
    return d


def _make_accent(chr_char, content_el):
    a = etree.Element(f"{{{M}}}acc")
    aPr = etree.SubElement(a, f"{{{M}}}accPr")
    c = etree.SubElement(aPr, f"{{{M}}}chr")
    c.set(f"{{{M}}}val", chr_char)
    e = etree.SubElement(a, f"{{{M}}}e")
    e.append(content_el)
    return a


def _make_bar(content_el, pos="top"):
    b = etree.Element(f"{{{M}}}bar")
    bPr = etree.SubElement(b, f"{{{M}}}barPr")
    p = etree.SubElement(bPr, f"{{{M}}}pos")
    p.set(f"{{{M}}}val", pos)
    e = etree.SubElement(b, f"{{{M}}}e")
    e.append(content_el)
    return b


def _make_function(name, arg_el):
    fn = etree.Element(f"{{{M}}}func")
    fName = etree.SubElement(fn, f"{{{M}}}fName")
    nr = etree.SubElement(fName, f"{{{M}}}r")
    nrPr = etree.SubElement(nr, f"{{{M}}}rPr")
    sty = etree.SubElement(nrPr, f"{{{M}}}sty")
    sty.set(f"{{{M}}}val", "p")
    nt = etree.SubElement(nr, f"{{{M}}}t")
    nt.set(XML_SPACE, "preserve")
    nt.text = name
    e = etree.SubElement(fn, f"{{{M}}}e")
    e.append(arg_el)
    return fn


def _make_limlow(e_el, lim_el):
    """Make m:limLow for \\lim_{x\\to 0} f(x) style."""
    ll = etree.Element(f"{{{M}}}limLow")
    e = etree.SubElement(ll, f"{{{M}}}e")
    e.append(e_el)
    lm = etree.SubElement(ll, f"{{{M}}}lim")
    lm.append(lim_el)
    return ll


def _make_groupChr(chr_char, pos, content_el):
    gc = etree.Element(f"{{{M}}}groupChr")
    gcPr = etree.SubElement(gc, f"{{{M}}}groupChrPr")
    c = etree.SubElement(gcPr, f"{{{M}}}chr")
    c.set(f"{{{M}}}val", chr_char)
    p = etree.SubElement(gcPr, f"{{{M}}}pos")
    p.set(f"{{{M}}}val", pos)
    e = etree.SubElement(gc, f"{{{M}}}e")
    e.append(content_el)
    return gc


def _make_borderBox(content_el):
    bb = etree.Element(f"{{{M}}}borderBox")
    etree.SubElement(bb, f"{{{M}}}borderBoxPr")
    e = etree.SubElement(bb, f"{{{M}}}e")
    e.append(content_el)
    return bb


def _make_matrix(rows_data, cols, bracket_type="matrix"):
    """rows_data: list of lists of element trees."""
    m = etree.Element(f"{{{M}}}m")
    mPr = etree.SubElement(m, f"{{{M}}}mPr")
    mcs = etree.SubElement(mPr, f"{{{M}}}mcs")
    mc = etree.SubElement(mcs, f"{{{M}}}mc")
    mcPr = etree.SubElement(mc, f"{{{M}}}mcPr")
    cnt = etree.SubElement(mcPr, f"{{{M}}}count")
    cnt.set(f"{{{M}}}val", str(cols))
    for row in rows_data:
        mr = etree.SubElement(m, f"{{{M}}}mr")
        for cell in row:
            e = etree.SubElement(mr, f"{{{M}}}e")
            if isinstance(cell, str):
                e.append(_make_run(cell))
            elif cell is not None:
                e.append(cell)
    left, right = _MATRIX_BRACKETS.get(bracket_type, ("", ""))
    if left or right:
        return _make_delimiter(left, right, m)
    return m

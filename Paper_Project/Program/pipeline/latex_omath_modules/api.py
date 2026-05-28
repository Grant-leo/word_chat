"""Public API helpers for LaTeX-to-OMML conversion."""
from __future__ import annotations

from lxml import etree

from .parser import _LaTeXParser
from .symbols import M, XML_SPACE
from .tokenizer import _tokenize

def latex_to_omath(latex_str, display=False):
    """Convert LaTeX math string to OOXML oMath XML string.

    Supports: fractions (\\frac), binomial (\\binom), roots (\\sqrt, \\sqrt[n]),
    sums (\\sum), integrals (\\int, \\iint, \\iiint, \\oint), products (\\prod),
    matrices (\\begin{pmatrix}...), cases (\\begin{cases}...),
    Greek letters (\\alpha, \\Gamma, ...), math symbols (\\infty, \\partial, ...),
    arrows (\\to, \\rightarrow, \\Rightarrow, ...),
    accents (\\hat, \\bar, \\vec, \\dot, \\ddot, \\tilde),
    overline/underline (\\overline, \\underline),
    named functions (\\sin, \\cos, \\log, \\lim, ...),
    braces (\\overbrace, \\underbrace), boxed (\\boxed),
    text mode (\\text, \\mathrm, \\mathbf), and more.

    Args:
        latex_str: LaTeX math string (e.g. r"\\frac{a}{b}")
        display: If True, wrap in m:oMathPara for display mode

    Returns:
        XML string of oMath element (or oMathPara if display=True)
    """
    if not latex_str or not latex_str.strip():
        omath = etree.Element(f'{{{M}}}oMath')
        return etree.tounicode(omath, with_tail=False)

    try:
        tokens = _tokenize(latex_str)
        parser = _LaTeXParser(tokens)
        omath = parser.parse()
    except Exception as e:
        omath = etree.Element(f'{{{M}}}oMath')
        err_r = etree.SubElement(omath, f'{{{M}}}r')
        etree.SubElement(err_r, f'{{{M}}}rPr')
        err_t = etree.SubElement(err_r, f'{{{M}}}t')
        err_t.set(XML_SPACE, 'preserve')
        err_t.text = f'[LaTeX error: {e}]'

    if display:
        omp = etree.Element(f'{{{M}}}oMathPara')
        omp.append(omath)
        return etree.tounicode(omp, with_tail=False)
    return etree.tounicode(omath, with_tail=False)


def formula_text_from_omath(xml_str):
    """Extract plain text from OOXML math formula."""
    parts = []
    for t in etree.fromstring(xml_str).iter(f'{{{M}}}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


def body_latex(doc, text, latex_str, display=True):
    """Convenience: add a paragraph with LaTeX formula to a python-docx Document.

    Usage identical to body_with_formula() but takes LaTeX string instead of OOXML.

    Args:
        doc: python-docx Document object
        text: preceding text (empty string if standalone)
        latex_str: LaTeX math string
        display: if True, center as display formula (m:oMathPara)
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if display else WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    if not display:
        pf.first_line_indent = Pt(21)

    if text.strip():
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        rp = r._element.get_or_add_rPr()
        rf = rp.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts'); rp.insert(0, rf)
        rf.set(qn('w:eastAsia'), '宋体')

    xml_str = latex_to_omath(latex_str, display=display)
    math_el = etree.fromstring(xml_str)
    p._element.append(math_el)
    return p


# ═══════════════════════════════════════════════════════════════════════════
#  SELF-TEST
# ═══════════════════════════════════════════════════════════════════════════


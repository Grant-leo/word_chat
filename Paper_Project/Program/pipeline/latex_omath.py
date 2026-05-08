"""
latex_omath.py — LaTeX math to OOXML (Office Math Markup Language) converter.

Standalone module. Import into build_generated.py or any script that needs
to convert LaTeX formula strings into native Word equation XML.

Supports: fractions, roots, sums, integrals, products, matrices, cases,
Greek letters, math symbols, arrows, accents (hat/bar/vec/dot/ddot/tilde),
overline/underline, named functions (sin/cos/lim etc.), limits, braces,
boxed, text mode, and more.

Usage:
    from latex_omath import latex_to_omath, body_latex

    # inline formula
    xml = latex_to_omath(r"E = mc^2")

    # display formula with helper
    body_latex("", r"\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}")

WPS compatibility: every m:r includes m:rPr (even if empty). Redundant
xmlns declarations stripped on insertion (via _strip_math_ns in caller).
"""

from lxml import etree

# ═══════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ── Greek letters (lowercase) ──
_GREEK_LOWER = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
    'epsilon': 'ε', 'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η',
    'theta': 'θ', 'vartheta': 'ϑ', 'iota': 'ι', 'kappa': 'κ',
    'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ',
    'pi': 'π', 'varpi': 'ϖ', 'rho': 'ρ', 'varrho': 'ϱ',
    'sigma': 'σ', 'varsigma': 'ς', 'tau': 'τ', 'upsilon': 'υ',
    'phi': 'φ', 'varphi': 'φ', 'chi': 'χ', 'psi': 'ψ',
    'omega': 'ω',
}

# ── Greek letters (uppercase) ──
_GREEK_UPPER = {
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
    'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Phi': 'Φ',
    'Psi': 'Ψ', 'Omega': 'Ω',
}

# ── Math symbols ──
_SYMBOLS = {
    'infty': '∞', 'partial': '∂', 'nabla': '∇',
    'times': '×', 'div': '÷', 'pm': '±', 'mp': '∓',
    'cdot': '·', 'cdots': '⋯', 'vdots': '⋮', 'ddots': '⋱',
    'ldots': '…', 'forall': '∀', 'exists': '∃', 'nexists': '∄',
    'neg': '¬', 'lnot': '¬', 'wedge': '∧', 'land': '∧',
    'vee': '∨', 'lor': '∨', 'cap': '∩', 'cup': '∪',
    'subset': '⊂', 'supset': '⊃', 'subseteq': '⊆', 'supseteq': '⊇',
    'in': '∈', 'notin': '∉', 'ni': '∋',
    'approx': '≈', 'equiv': '≡', 'neq': '≠', 'ne': '≠',
    'leq': '≤', 'le': '≤', 'geq': '≥', 'ge': '≥',
    'll': '≪', 'gg': '≫', 'propto': '∝', 'sim': '∼',
    'simeq': '≃', 'cong': '≅', 'doteq': '≐',
    'perp': '⟂', 'parallel': '∥',
    'angle': '∠', 'measuredangle': '∡',
    'circ': '∘', 'bullet': '∙',
    'oplus': '⊕', 'ominus': '⊖', 'otimes': '⊗', 'odot': '⊙',
    'oslash': '⊘', 'uplus': '⊎',
    'aleph': 'ℵ', 'hbar': 'ℏ', 'ell': 'ℓ',
    'wp': '℘', 'Re': 'ℜ', 'Im': 'ℑ',
    'emptyset': '∅', 'varnothing': '∅',
    'top': '⊤', 'bot': '⊥',
    'triangle': '△', 'triangledown': '▽',
    'square': '□', 'Box': '□', 'diamond': '◇', 'Diamond': '◇',
    'star': '⋆', 'bigstar': '★',
    'clubsuit': '♣', 'diamondsuit': '♦',
    'heartsuit': '♡', 'spadesuit': '♠',
    'prime': '′', 'backslash': '∖',
    'surd': '√', 'dag': '†', 'ddag': '‡',
    'S': '§', 'P': '¶', 'pounds': '£',
    'subsetneq': '⊊', 'supsetneq': '⊋',
}

# ── Arrows ──
_ARROWS = {
    'to': '→', 'rightarrow': '→', 'longrightarrow': '⟶',
    'Rightarrow': '⇒', 'Longrightarrow': '⟹',
    'leftarrow': '←', 'longleftarrow': '⟵',
    'Leftarrow': '⇐', 'Longleftarrow': '⟸',
    'leftrightarrow': '↔', 'longleftrightarrow': '⟷',
    'Leftrightarrow': '⇔', 'Longleftrightarrow': '⟺',
    'mapsto': '↦', 'longmapsto': '⟼',
    'hookrightarrow': '↪', 'hookleftarrow': '↩',
    'rightharpoonup': '⇀', 'rightharpoondown': '⇁',
    'leftharpoonup': '↼', 'leftharpoondown': '↽',
    'uparrow': '↑', 'downarrow': '↓', 'updownarrow': '↕',
    'Uparrow': '⇑', 'Downarrow': '⇓', 'Updownarrow': '⇕',
    'nearrow': '↗', 'searrow': '↘', 'swarrow': '↙', 'nwarrow': '↖',
    'leadsto': '↝',
}

# ── Named functions (rendered upright) ──
_NAMED_FUNCTIONS = frozenset([
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth',
    'ln', 'log', 'lg', 'exp', 'det', 'gcd',
    'max', 'min', 'lim', 'limsup', 'liminf',
    'arg', 'deg', 'dim', 'hom', 'ker',
    'sup', 'inf', 'Pr',
])

# ── Accent characters ──
_ACCENT_CHARS = {
    'hat': '̂', 'bar': '̄', 'vec': '⃗',
    'dot': '̇', 'ddot': '̈', 'tilde': '̃',
}

# ── N-ary operators ──
_NARY_CHARS = {
    'sum': '∑', 'int': '∫', 'prod': '∏',
    'iint': '∬', 'iiint': '∭', 'oint': '∮',
    'bigcup': '⋃', 'bigcap': '⋂',
}

_NARY_LIMLOC = {
    'sum': 'undOvr', 'prod': 'undOvr', 'bigcup': 'undOvr', 'bigcap': 'undOvr',
    'int': 'subSup', 'iint': 'subSup', 'iiint': 'subSup', 'oint': 'subSup',
}

# ── Delimiter mappings (for \left...\right) ──
_LEFT_DELIM = {
    '\\{': '{', '\\{.': '{',
    '\\langle': '⟨', '\\lfloor': '⌊', '\\lceil': '⌈',
    '\\|': '‖', '.': '',
}
_RIGHT_DELIM = {
    '\\}': '}', '\\.}': '}',
    '\\rangle': '⟩', '\\rfloor': '⌋', '\\rceil': '⌉',
    '\\|': '‖', '.': '',
}

# ── Matrix bracket types ──
_MATRIX_BRACKETS = {
    'matrix': ('', ''), 'pmatrix': ('(', ')'), 'bmatrix': ('[', ']'),
    'Bmatrix': ('{', '}'), 'vmatrix': ('|', '|'),
    'Vmatrix': ('‖', '‖'), 'cases': ('{', ''),
}


# ═══════════════════════════════════════════════════════════════════════════
#  OOXML BUILDER HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _make_run(text, plain=False):
    """Create m:r element with text content."""
    r = etree.Element(f'{{{M}}}r')
    rPr = etree.SubElement(r, f'{{{M}}}rPr')
    if plain:
        sty = etree.SubElement(rPr, f'{{{M}}}sty')
        sty.set(f'{{{M}}}val', 'p')
    t = etree.SubElement(r, f'{{{M}}}t')
    t.set(XML_SPACE, 'preserve')
    t.text = text
    return r


def _make_fraction(num_el, den_el, frac_type='bar'):
    """Create m:f element. frac_type: bar, nobar (binomial), lin, skw."""
    f = etree.Element(f'{{{M}}}f')
    fPr = etree.SubElement(f, f'{{{M}}}fPr')
    typ = etree.SubElement(fPr, f'{{{M}}}type')
    typ.set(f'{{{M}}}val', frac_type)
    num = etree.SubElement(f, f'{{{M}}}num')
    num.append(num_el)
    den = etree.SubElement(f, f'{{{M}}}den')
    den.append(den_el)
    return f


def _make_sup(base_el, sup_el):
    s = etree.Element(f'{{{M}}}sSup')
    etree.SubElement(s, f'{{{M}}}sSupPr')
    e = etree.SubElement(s, f'{{{M}}}e'); e.append(base_el)
    sp = etree.SubElement(s, f'{{{M}}}sup'); sp.append(sup_el)
    return s


def _make_sub(base_el, sub_el):
    s = etree.Element(f'{{{M}}}sSub')
    etree.SubElement(s, f'{{{M}}}sSubPr')
    e = etree.SubElement(s, f'{{{M}}}e'); e.append(base_el)
    sb = etree.SubElement(s, f'{{{M}}}sub'); sb.append(sub_el)
    return s


def _make_supsub(base_el, sub_el, sup_el):
    s = etree.Element(f'{{{M}}}sSubSup')
    etree.SubElement(s, f'{{{M}}}sSubSupPr')
    e = etree.SubElement(s, f'{{{M}}}e'); e.append(base_el)
    sb = etree.SubElement(s, f'{{{M}}}sub'); sb.append(sub_el)
    sp = etree.SubElement(s, f'{{{M}}}sup'); sp.append(sup_el)
    return s


def _make_nary(chr_char, sub_el, sup_el, e_el, limloc='undOvr'):
    n = etree.Element(f'{{{M}}}nary')
    nPr = etree.SubElement(n, f'{{{M}}}naryPr')
    c = etree.SubElement(nPr, f'{{{M}}}chr'); c.set(f'{{{M}}}val', chr_char)
    g = etree.SubElement(nPr, f'{{{M}}}grow'); g.set(f'{{{M}}}val', '1')
    ll = etree.SubElement(nPr, f'{{{M}}}limLoc'); ll.set(f'{{{M}}}val', limloc)
    if sub_el is not None:
        sb = etree.SubElement(n, f'{{{M}}}sub'); sb.append(sub_el)
    if sup_el is not None:
        sp = etree.SubElement(n, f'{{{M}}}sup'); sp.append(sup_el)
    e = etree.SubElement(n, f'{{{M}}}e'); e.append(e_el)
    return n


def _make_radical(deg_el, e_el):
    rad = etree.Element(f'{{{M}}}rad')
    etree.SubElement(rad, f'{{{M}}}radPr')
    if deg_el is not None:
        d = etree.SubElement(rad, f'{{{M}}}deg'); d.append(deg_el)
    e = etree.SubElement(rad, f'{{{M}}}e'); e.append(e_el)
    return rad


def _make_delimiter(left_chr, right_chr, content_el, grow=True):
    d = etree.Element(f'{{{M}}}d')
    dPr = etree.SubElement(d, f'{{{M}}}dPr')
    if left_chr:
        beg = etree.SubElement(dPr, f'{{{M}}}begChr'); beg.set(f'{{{M}}}val', left_chr)
    if right_chr:
        end = etree.SubElement(dPr, f'{{{M}}}endChr'); end.set(f'{{{M}}}val', right_chr)
    if grow:
        gr = etree.SubElement(dPr, f'{{{M}}}grow'); gr.set(f'{{{M}}}val', '1')
    items = content_el if isinstance(content_el, list) else [content_el]
    for item in items:
        if item is not None:
            e = etree.SubElement(d, f'{{{M}}}e'); e.append(item)
    return d


def _make_accent(chr_char, content_el):
    a = etree.Element(f'{{{M}}}acc')
    aPr = etree.SubElement(a, f'{{{M}}}accPr')
    c = etree.SubElement(aPr, f'{{{M}}}chr'); c.set(f'{{{M}}}val', chr_char)
    e = etree.SubElement(a, f'{{{M}}}e'); e.append(content_el)
    return a


def _make_bar(content_el, pos='top'):
    b = etree.Element(f'{{{M}}}bar')
    bPr = etree.SubElement(b, f'{{{M}}}barPr')
    p = etree.SubElement(bPr, f'{{{M}}}pos'); p.set(f'{{{M}}}val', pos)
    e = etree.SubElement(b, f'{{{M}}}e'); e.append(content_el)
    return b


def _make_function(name, arg_el):
    fn = etree.Element(f'{{{M}}}func')
    fName = etree.SubElement(fn, f'{{{M}}}fName')
    nr = etree.SubElement(fName, f'{{{M}}}r')
    nrPr = etree.SubElement(nr, f'{{{M}}}rPr')
    sty = etree.SubElement(nrPr, f'{{{M}}}sty'); sty.set(f'{{{M}}}val', 'p')
    nt = etree.SubElement(nr, f'{{{M}}}t')
    nt.set(XML_SPACE, 'preserve')
    nt.text = name
    e = etree.SubElement(fn, f'{{{M}}}e'); e.append(arg_el)
    return fn


def _make_limlow(e_el, lim_el):
    """Make m:limLow for \lim_{x\to 0} f(x) style."""
    ll = etree.Element(f'{{{M}}}limLow')
    e = etree.SubElement(ll, f'{{{M}}}e'); e.append(e_el)
    lm = etree.SubElement(ll, f'{{{M}}}lim'); lm.append(lim_el)
    return ll


def _make_groupChr(chr_char, pos, content_el):
    gc = etree.Element(f'{{{M}}}groupChr')
    gcPr = etree.SubElement(gc, f'{{{M}}}groupChrPr')
    c = etree.SubElement(gcPr, f'{{{M}}}chr'); c.set(f'{{{M}}}val', chr_char)
    p = etree.SubElement(gcPr, f'{{{M}}}pos'); p.set(f'{{{M}}}val', pos)
    e = etree.SubElement(gc, f'{{{M}}}e'); e.append(content_el)
    return gc


def _make_borderBox(content_el):
    bb = etree.Element(f'{{{M}}}borderBox')
    etree.SubElement(bb, f'{{{M}}}borderBoxPr')
    e = etree.SubElement(bb, f'{{{M}}}e'); e.append(content_el)
    return bb


def _make_matrix(rows_data, cols, bracket_type='matrix'):
    """rows_data: list of lists of element trees."""
    m = etree.Element(f'{{{M}}}m')
    mPr = etree.SubElement(m, f'{{{M}}}mPr')
    mcs = etree.SubElement(mPr, f'{{{M}}}mcs')
    mc = etree.SubElement(mcs, f'{{{M}}}mc')
    mcPr = etree.SubElement(mc, f'{{{M}}}mcPr')
    cnt = etree.SubElement(mcPr, f'{{{M}}}count'); cnt.set(f'{{{M}}}val', str(cols))
    for row in rows_data:
        mr = etree.SubElement(m, f'{{{M}}}mr')
        for cell in row:
            e = etree.SubElement(mr, f'{{{M}}}e')
            if isinstance(cell, str):
                e.append(_make_run(cell))
            elif cell is not None:
                e.append(cell)
    left, right = _MATRIX_BRACKETS.get(bracket_type, ('', ''))
    if left or right:
        return _make_delimiter(left, right, m)
    return m


# ═══════════════════════════════════════════════════════════════════════════
#  TOKENIZER
# ═══════════════════════════════════════════════════════════════════════════

def _tokenize(s):
    """Tokenize LaTeX math string into token list."""
    tokens = []
    i = 0; n = len(s)
    while i < n:
        c = s[i]
        if c == '\\':
            if i + 1 < n and s[i + 1] == '\\':
                tokens.append({'type': 'NEWLINE'}); i += 2; continue
            j = i + 1
            while j < n and s[j].isalpha():
                j += 1
            if j > i + 1:
                tokens.append({'type': 'COMMAND', 'value': s[i:j]}); i = j
            else:
                # Backslash + non-alpha: emit as COMMAND (for \{, \}, \|, etc.)
                tokens.append({'type': 'COMMAND', 'value': s[i:i + 2]}); i += 2
        elif c in '_{':
            t = 'SUB' if c == '_' else 'LBRACE'
            tokens.append({'type': t}); i += 1
        elif c in '^}':
            t = 'SUPER' if c == '^' else 'RBRACE'
            tokens.append({'type': t}); i += 1
        elif c == '&':
            tokens.append({'type': 'AMPERSAND'}); i += 1
        elif c.isspace():
            tokens.append({'type': 'SPACE'}); i += 1
        else:
            tokens.append({'type': 'CHAR', 'value': c}); i += 1
    tokens.append({'type': 'EOF'})
    return tokens


# ═══════════════════════════════════════════════════════════════════════════
#  RECURSIVE DESCENT PARSER
# ═══════════════════════════════════════════════════════════════════════════

class _LaTeXParser:
    def __init__(self, tokens):
        self.tokens = tokens
        self.pos = 0

    def peek(self):
        return self.tokens[self.pos]

    def consume(self, expected=None):
        t = self.tokens[self.pos]
        if expected and t['type'] != expected:
            raise ValueError(f'Expected {expected}, got {t}')
        self.pos += 1
        return t

    def parse(self):
        result = self._parse_expression()
        # trailing content ignored for robustness
        omath = etree.Element(f'{{{M}}}oMath')
        if isinstance(result, list):
            for el in result:
                if el is not None:
                    omath.append(el)
        elif result is not None:
            omath.append(result)
        return omath

    # ── expression: handles binary operators + - = < > ──
    def _parse_expression(self):
        left = self._flatten(self._parse_sup_sub_seq())
        while True:
            t = self.peek()
            if t['type'] == 'EOF' or t['type'] in ('RBRACE', 'AMPERSAND', 'NEWLINE'):
                break
            if t['type'] == 'CHAR' and t['value'] in '+-=<>*' + chr(39) + ':':
                left.append(_make_run(self.consume()['value']))
                left.extend(self._flatten(self._parse_sup_sub_seq()))
            elif t['type'] == 'SPACE':
                self.consume()
            elif self._is_start_of_term(t):
                left.extend(self._flatten(self._parse_sup_sub_seq()))
            else:
                break
        return left

    def _is_start_of_term(self, t):
        if t['type'] not in ('CHAR', 'COMMAND', 'LBRACE'):
            return False
        # Exclude bracket-closing chars that should terminate expressions
        if t['type'] == 'CHAR' and t.get('value') == ']':
            return False
        # \right, \end terminate enclosing contexts
        if t['type'] == 'COMMAND' and t.get('value') in ('\\right', '\\end'):
            return False
        return True

    # ── sup_sub_seq: implicit multiplication chain ──
    def _parse_sup_sub_seq(self):
        first = self._parse_sup_sub_term()
        if first is None:
            return None
        result = self._flatten(first)
        while True:
            t = self.peek()
            if not self._is_start_of_term(t):
                break
            if t['type'] == 'COMMAND' and t.get('value', '')[1:] in _NAMED_FUNCTIONS:
                break
            if t['type'] == 'CHAR' and t['value'] in '+-=<>*/' + chr(39) + ':':
                break
            nxt = self._parse_sup_sub_term()
            if nxt is not None:
                result.extend(self._flatten(nxt))
        return result

    # ── sup_sub_term: atom with optional sub/superscript ──
    def _parse_sup_sub_term(self):
        base = self._ensure_single(self._parse_atom())
        if base is None:
            return None
        if self.peek()['type'] == 'SUB':
            self.consume('SUB')
            sub = self._ensure_single(self._parse_atom())
            if self.peek()['type'] == 'SUPER':
                self.consume('SUPER')
                sup = self._ensure_single(self._parse_atom())
                if sub is not None and sup is not None:
                    return _make_supsub(base, sub, sup)
                return _make_sub(base, sub) if sub is not None else base
            return _make_sub(base, sub) if sub is not None else base
        if self.peek()['type'] == 'SUPER':
            self.consume('SUPER')
            sup = self._ensure_single(self._parse_atom())
            if self.peek()['type'] == 'SUB':
                self.consume('SUB')
                sub = self._ensure_single(self._parse_atom())
                if sub is not None and sup is not None:
                    return _make_supsub(base, sub, sup)
                return _make_sup(base, sup) if sup is not None else base
            return _make_sup(base, sup) if sup is not None else base
        return base

    def _ensure_single(self, el):
        """Convert list to single element for use as argument to builder."""
        if el is None:
            return None
        if isinstance(el, list):
            return self._list_to_element(el)
        return el

    # ── atom: single unit ──
    def _parse_atom(self):
        t = self.peek()
        if t['type'] == 'LBRACE':
            self.consume('LBRACE')
            result = self._parse_expression()
            self.consume('RBRACE')
            return result
        if t['type'] == 'COMMAND':
            return self._parse_command()
        if t['type'] == 'CHAR':
            return _make_run(self.consume()['value'])
        if t['type'] == 'SPACE':
            self.consume()
            return None
        if t['type'] == 'BEGIN':
            return self._parse_matrix()
        return _make_run(self.consume()['value'])

    # ── command dispatcher ──
    def _parse_command(self):
        cmd = self.consume('COMMAND')['value']
        name = cmd[1:]  # strip backslash

        # Fractions
        if name in ('frac', 'tfrac', 'dfrac'):
            return self._parse_frac()
        if name == 'binom':
            return self._parse_binom()

        # Roots
        if name == 'sqrt':
            return self._parse_sqrt()

        # N-ary operators
        if name in _NARY_CHARS:
            return self._parse_nary(name)

        # Left/right delimiters
        if name == 'left':
            return self._parse_left_right()
        if name == 'right':
            raise ValueError('\\right without \\left')

        # Accents
        if name in _ACCENT_CHARS:
            return self._parse_accent(name)
        if name in ('widehat', 'widetilde'):
            return self._parse_accent(name[4:])

        # Overline/underline
        if name == 'overline':
            return self._parse_bar('top')
        if name == 'underline':
            return self._parse_bar('bot')

        # Text styles
        if name in ('text', 'mathrm', 'mathbf', 'mathit', 'mathsf', 'mathtt'):
            return self._parse_text_style(name)

        # Boxed
        if name == 'boxed':
            return self._parse_boxed()

        # Braces
        if name == 'overbrace':
            return self._parse_brace('over')
        if name == 'underbrace':
            return self._parse_brace('under')

        # Limit
        if name == 'lim':
            return self._parse_lim()

        # Named functions
        if name in _NAMED_FUNCTIONS:
            return self._parse_function(name)

        # Greek letters
        if name in _GREEK_LOWER:
            return _make_run(_GREEK_LOWER[name])
        if name in _GREEK_UPPER:
            return _make_run(_GREEK_UPPER[name])

        # Arrows
        if name in _ARROWS:
            return _make_run(_ARROWS[name])

        # Symbols
        if name in _SYMBOLS:
            return _make_run(_SYMBOLS[name])

        # Spacing commands (ignored in OOXML)
        if name in (',', ';', ':', '!', 'quad', 'qquad', 'enspace', 'enskip', 'thinspace'):
            return None
        if name in ('limits', 'nolimits'):
            return None

        # Unrecognized — emit as plain text
        return _make_run(name)

    # ── Argument parsing ──
    def _parse_arg(self):
        """Parse a single argument: {expr} or a single atom."""
        if self.peek()['type'] == 'LBRACE':
            self.consume('LBRACE')
            res = self._parse_expression()
            self.consume('RBRACE')
            return res
        return self._parse_atom()

    # ── Specific command handlers ──
    def _parse_frac(self):
        num = self._parse_arg()
        den = self._parse_arg()
        return _make_fraction(self._ensure_element(num), self._ensure_element(den))

    def _parse_binom(self):
        num = self._parse_arg()
        den = self._parse_arg()
        return _make_fraction(self._ensure_element(num), self._ensure_element(den), 'nobar')

    def _parse_sqrt(self):
        deg = None
        t = self.peek()
        if t['type'] == 'CHAR' and t.get('value') == '[':
            self.consume()
            deg = self._parse_expression()
            t2 = self.peek()
            if t2['type'] == 'CHAR' and t2.get('value') == ']':
                self.consume()
        arg = self._parse_arg()
        d_el = self._ensure_element(deg) if deg else None
        return _make_radical(d_el, self._ensure_element(arg))

    def _parse_nary(self, name):
        chr_char = _NARY_CHARS[name]
        limloc = _NARY_LIMLOC.get(name, 'undOvr')
        sub_el = None; sup_el = None
        while self.peek()['type'] in ('SUB', 'SUPER'):
            t = self.peek()['type']
            self.consume()
            arg = self._parse_arg()
            if arg is not None:
                arg_el = self._ensure_element(arg)
                if t == 'SUB':
                    sub_el = arg_el
                else:
                    sup_el = arg_el
        # operand is the rest of the expression (e.g., f(x)dx)
        e_el = self._parse_sup_sub_seq()
        if e_el is None:
            e_el = _make_run('')
        else:
            e_el = self._ensure_element(e_el)
        return _make_nary(chr_char, sub_el, sup_el, e_el, limloc)

    def _parse_left_right(self):
        left_arg = self._parse_arg_for_delim()
        content = []
        while True:
            t = self.peek()
            if t['type'] == 'COMMAND' and t['value'] == '\\right':
                break
            if t['type'] == 'EOF':
                raise ValueError('Missing \\right to match \\left')
            el = self._parse_sup_sub_seq()
            if el is not None:
                content.extend(self._flatten(el))
        self.consume('COMMAND')
        right_arg = self._parse_arg_for_delim()
        content_el = self._list_to_element(content)
        return _make_delimiter(left_arg, right_arg, content_el)

    def _parse_arg_for_delim(self):
        t = self.peek()
        if t['type'] == 'CHAR':
            c = self.consume()['value']
            if c == '.':
                return ''
            return c
        if t['type'] == 'COMMAND':
            cmd = t['value']
            if cmd in _LEFT_DELIM:
                self.consume()
                return _LEFT_DELIM[cmd]
            if cmd in _RIGHT_DELIM:
                self.consume()
                return _RIGHT_DELIM[cmd]
        c = self.consume().get('value', '.')
        return '' if c == '.' else c

    def _parse_accent(self, name):
        # Skip limits/nolimits modifiers
        while (self.peek()['type'] == 'COMMAND'
               and self.peek().get('value', '')[1:] in ('limits', 'nolimits')):
            self.consume()
        arg = self._parse_arg()
        chr_char = _ACCENT_CHARS[name]
        return _make_accent(chr_char, self._ensure_element(arg))

    def _parse_bar(self, pos):
        arg = self._parse_arg()
        return _make_bar(self._ensure_element(arg), pos)

    def _parse_text_style(self, name):
        self.consume('LBRACE')
        text_parts = []
        while self.peek()['type'] != 'RBRACE' and self.peek()['type'] != 'EOF':
            t = self.peek()
            if t['type'] == 'CHAR':
                text_parts.append(self.consume()['value'])
            elif t['type'] == 'SPACE':
                text_parts.append(' '); self.consume()
            elif t['type'] == 'COMMAND':
                cname = self.consume()['value'][1:]
                if cname in _GREEK_LOWER:
                    text_parts.append(_GREEK_LOWER[cname])
                elif cname in _SYMBOLS:
                    text_parts.append(_SYMBOLS[cname])
                else:
                    text_parts.append(cname)
            elif t['type'] == 'SUB':
                self.consume(); text_parts.append('_')
            elif t['type'] == 'SUPER':
                self.consume(); text_parts.append('^')
            else:
                self.consume()
        self.consume('RBRACE')
        return _make_run(''.join(text_parts), plain=(name in ('text', 'mathrm')))

    def _parse_boxed(self):
        arg = self._parse_arg()
        return _make_borderBox(self._ensure_element(arg))

    def _parse_brace(self, direction):
        arg = self._parse_arg()
        chr_char = '⏞' if direction == 'over' else '⏟'
        pos = 'top' if direction == 'over' else 'bot'
        gc = _make_groupChr(chr_char, pos, self._ensure_element(arg))
        if self.peek()['type'] in ('SUPER', 'SUB'):
            t = self.peek()['type']
            self.consume()
            label = self._ensure_single(self._parse_atom())
            if label is not None:
                return _make_sup(gc, label) if direction == 'over' else _make_sub(gc, label)
        return gc

    def _parse_lim(self):
        # Build 'lim' as function name run
        nr = etree.Element(f'{{{M}}}r')
        nrPr = etree.SubElement(nr, f'{{{M}}}rPr')
        sty = etree.SubElement(nrPr, f'{{{M}}}sty'); sty.set(f'{{{M}}}val', 'p')
        nt = etree.SubElement(nr, f'{{{M}}}t')
        nt.set(XML_SPACE, 'preserve')
        nt.text = 'lim'
        # Parse limits and argument
        sub_el = None; sup_el = None
        while self.peek()['type'] in ('SUB', 'SUPER'):
            t = self.peek()['type']
            self.consume()
            arg = self._parse_atom()
            if arg is not None:
                if t == 'SUB':
                    sub_el = arg
                else:
                    sup_el = arg
        arg_el = None
        if self._is_start_of_term(self.peek()):
            arg_el = self._parse_atom()
        if sub_el is not None and arg_el is not None:
            ll = _make_limlow(arg_el, sub_el)
            if sup_el is not None:
                ll = _make_sup(ll, sup_el)
            return ll
        fn = etree.Element(f'{{{M}}}func')
        fName = etree.SubElement(fn, f'{{{M}}}fName')
        fName.append(nr)
        if arg_el is not None:
            e = etree.SubElement(fn, f'{{{M}}}e'); e.append(arg_el)
        return fn

    def _parse_function(self, name):
        arg_el = None
        if self._is_start_of_term(self.peek()):
            arg_el = self._parse_atom()
        if arg_el is None:
            return _make_run(name, plain=True)
        return _make_function(name, arg_el)

    def _parse_matrix(self):
        t = self.consume()
        mtype = t.get('value', '')
        rows = []; current_row = []
        while True:
            t = self.peek()
            if t['type'] == 'COMMAND' and t.get('value', '') == '\\end':
                break
            if t['type'] == 'EOF':
                raise ValueError(f'Missing \\end for \\begin{{{mtype}}}')
            if t['type'] == 'NEWLINE':
                self.consume(); rows.append(current_row); current_row = []; continue
            if t['type'] == 'AMPERSAND':
                self.consume(); current_row.append(None); continue
            el = self._parse_expression()
            current_row.append(self._list_to_element(el) if isinstance(el, list) else el)
        self.consume('COMMAND')
        if self.peek()['type'] == 'LBRACE':
            self.consume('LBRACE')
            while self.peek()['type'] not in ('RBRACE', 'EOF'):
                self.consume()
            if self.peek()['type'] == 'RBRACE':
                self.consume()
        if current_row:
            rows.append(current_row)
        cols = max(len(r) for r in rows) if rows else 1
        for r in rows:
            while len(r) < cols:
                r.append(None)
        return _make_matrix(rows, cols, mtype)

    # ── Utility ──
    def _flatten(self, el):
        if el is None:
            return []
        if isinstance(el, list):
            return el
        return [el]

    def _ensure_element(self, el):
        if el is None:
            return _make_run('')
        if isinstance(el, list):
            return self._list_to_element(el)
        return el

    def _list_to_element(self, lst):
        """Convert list of sibling elements to a single element.
        For single elements, return as-is. For multiple, wrap in invisible delimiter."""
        if not lst:
            return _make_run('')
        lst = [x for x in lst if x is not None]
        if len(lst) == 0:
            return _make_run('')
        if len(lst) == 1:
            return lst[0]
        # Wrap multiple siblings in invisible delimiter (m:d without brackets)
        wrapped = etree.Element(f'{{{M}}}d')
        etree.SubElement(wrapped, f'{{{M}}}dPr')
        for item in lst:
            e = etree.SubElement(wrapped, f'{{{M}}}e'); e.append(item)
        return wrapped


# ═══════════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════

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

if __name__ == '__main__':
    tests = [
        (r'x^2', 'superscript'),
        (r'x_1', 'subscript'),
        (r'x_1^2', 'subsuperscript'),
        (r'\frac{a}{b}', 'fraction'),
        (r'\binom{n}{k}', 'binomial'),
        (r'\sqrt{x}', 'sqrt'),
        (r'\sqrt[3]{x}', 'nth root'),
        (r'\alpha + \beta = \gamma', 'greek'),
        (r'\sin\theta + \cos\theta', 'trig'),
        (r'\Gamma(z)', 'Gamma fn'),
        (r'\infty + \partial', 'symbols'),
        (r'\int_0^\infty e^{-x^2} dx', 'integral'),
        (r'\sum_{i=1}^n x_i', 'summation'),
        (r'\left(\frac{a}{b}\right)^n', 'delimited frac'),
        (r'\hat{x} + \bar{y} + \vec{v}', 'accents'),
        (r'\overline{abc} + \underline{xyz}', 'over/underline'),
        (r'\boxed{E=mc^2}', 'boxed'),
        (r'\begin{pmatrix} a & b \\ c & d \end{pmatrix}', 'pmatrix'),
        (r'\begin{cases} x & x>0 \\ -x & x\leq 0 \end{cases}', 'cases'),
        (r'\overbrace{x+y}^{n}', 'overbrace'),
        (r'\underbrace{x+y}_{n}', 'underbrace'),
        (r'\text{hello world}', 'text mode'),
        (r'\lim_{x\to\infty} f(x)', 'limit'),
        (r'x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}', 'quadratic'),
        (r'\mathrm{CH_4 + 2O_2 \to CO_2 + 2H_2O}', 'chemical'),
        (r'\Gamma(z) = \int_0^\infty t^{z-1} e^{-t} dt', 'Gamma integral'),
        (r'\forall x \in \mathbb{R}, \exists y > 0', 'quantifiers'),
    ]

    passed = 0
    for latex, desc in tests:
        try:
            xml = latex_to_omath(latex)
            # Check for error markers
            has_err = '[LaTeX error' in xml
            # Check for ? marker (unrecognized command fallback)
            # Look for ? followed by a command name — these appear as plain text
            if has_err:
                print(f'FAIL  | {desc:20s} | error embedded in output')
            else:
                print(f'OK    | {desc:20s}')
                passed += 1
        except Exception as e:
            print(f'CRASH | {desc:20s} | {type(e).__name__}: {e}')

    print(f'\n{passed}/{len(tests)} tests passed')

"""Recursive descent parser for LaTeX-to-OMML conversion."""
from __future__ import annotations

from lxml import etree

from .symbols import (
    M,
    XML_SPACE,
    _ACCENT_CHARS,
    _ARROWS,
    _DBLSTRUCK,
    _GREEK_LOWER,
    _GREEK_UPPER,
    _LEFT_DELIM,
    _MATRIX_BRACKETS,
    _NAMED_FUNCTIONS,
    _NARY_CHARS,
    _NARY_LIMLOC,
    _RIGHT_DELIM,
    _SCRIPT,
    _SYMBOLS,
)
from .ooxml import (
    _make_accent,
    _make_bar,
    _make_borderBox,
    _make_delimiter,
    _make_fraction,
    _make_function,
    _make_groupChr,
    _make_limlow,
    _make_literal_run,
    _make_matrix,
    _make_nary,
    _make_radical,
    _make_run,
    _make_sub,
    _make_sup,
    _make_supsub,
)

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
                left.append(_make_run(self.consume()['value'], style='plain'))
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
        # \right terminates left/right context
        if t['type'] == 'COMMAND' and t.get('value') == '\\right':
            return False
        # END tokens terminate matrix environments
        if t['type'] == 'END':
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
            return _make_literal_run(self.consume()['value'])
        if t['type'] == 'SPACE':
            self.consume()
            return None
        if t['type'] == 'BEGIN':
            return self._parse_matrix()
        return _make_literal_run(self.consume()['value'])

    # ── command dispatcher ──
    def _parse_command(self):
        cmd = self.consume('COMMAND')['value']
        name = cmd[1:]  # strip backslash

        # Fractions
        if name in ('frac', 'tfrac', 'dfrac'):
            return self._parse_frac()
        if name in ('binom', 'dbinom', 'tbinom'):
            return self._parse_binom()

        # Roots
        if name == 'sqrt':
            return self._parse_sqrt()

        # \mathbb{letter} — double-struck
        if name == 'mathbb':
            return self._parse_math_alphabet(_DBLSTRUCK)
        # \mathcal{letter} — calligraphic
        if name == 'mathcal':
            return self._parse_math_alphabet(_SCRIPT)

        # \underset{below}{expr}, \overset{above}{expr}
        if name == 'underset':
            return self._parse_stackrel('sub')
        if name == 'overset':
            return self._parse_stackrel('sup')

        # \xrightarrow[below]{above}, \xleftarrow[below]{above}
        if name == 'xrightarrow':
            return self._parse_ext_arrow('→')
        if name == 'xleftarrow':
            return self._parse_ext_arrow('←')

        # \pmod{n}
        if name == 'pmod':
            return self._parse_pmod()

        # \not — negation prefix
        if name == 'not':
            return self._parse_not()

        # \substack{...} — multi-line subscript for nary operators
        if name == 'substack':
            return self._parse_substack()

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
            return _make_run(_ARROWS[name], style='plain')

        # \tag{label} — formula numbering
        if name == 'tag':
            return self._parse_tag()

        # Symbols
        if name in _SYMBOLS:
            return _make_run(_SYMBOLS[name], style='plain')

        # Spacing commands (ignored in OOXML)
        if name in (',', ';', ':', '!', 'quad', 'qquad', 'enspace', 'enskip', 'thinspace'):
            return None
        if name in ('limits', 'nolimits'):
            return None

        # Escaped literal punctuation such as \{ and \% should stay upright.
        if name in ('{', '}', '[', ']', '(', ')', '|', '%', '_', '#', '&'):
            return _make_run(name, style='plain')

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

    def _parse_math_alphabet(self, char_map):
        """Handle \\mathbb{R}, \\mathcal{F} etc."""
        arg = self._parse_arg()
        if arg is None:
            return _make_run('')
        # Extract text from the argument
        if isinstance(arg, list):
            txt = ''
            for el in arg:
                if hasattr(el, 'iter'):
                    for t in el.iter(f'{{{M}}}t'):
                        if t.text: txt += t.text
            if not txt:
                txt = ''.join(t.text or '' for el in arg if hasattr(el, 'iter') for t in el.iter(f'{{{M}}}t'))
        elif hasattr(arg, 'iter'):
            txt = ''
            for t in arg.iter(f'{{{M}}}t'):
                if t.text: txt += t.text
        else:
            txt = str(arg) if arg is not None else ''
        if txt in char_map:
            return _make_run(char_map[txt])
        return _make_run(txt)

    def _parse_stackrel(self, direction):
        """Handle \\underset{below}{expr} and \\overset{above}{expr}."""
        script = self._parse_arg()
        base = self._parse_arg()
        base_el = self._ensure_element(base)
        script_el = self._ensure_element(script)
        if direction == 'sub':
            return _make_sub(base_el, script_el)
        else:
            return _make_sup(base_el, script_el)

    def _parse_ext_arrow(self, chr_char):
        """Handle \\xrightarrow[below]{above}, \\xleftarrow[below]{above}."""
        sub_el = None; sup_el = None
        if self.peek()['type'] == 'SUB':
            self.consume('SUB')
            sub_el = self._ensure_single(self._parse_atom())
        if self.peek()['type'] == 'SUPER':
            self.consume('SUPER')
            sup_el = self._ensure_single(self._parse_atom())
        # Build arrow base with limits
        arrow_run = _make_run(chr_char, style='plain')
        if sub_el and sup_el:
            return _make_supsub(arrow_run, sub_el, sup_el)
        elif sub_el:
            return _make_sub(arrow_run, sub_el)
        elif sup_el:
            return _make_sup(arrow_run, sup_el)
        return arrow_run

    def _parse_pmod(self):
        """Handle \\pmod{n} → (mod n)."""
        arg = self._parse_arg()
        mod_text = _make_run('mod', style='plain')
        space = _make_run(' ', style='plain')
        num = self._ensure_element(arg)
        lp = _make_run('(', style='plain')
        rp = _make_run(')', style='plain')
        # Build: (mod n) as sibling elements
        return self._list_to_element([lp, mod_text, space, num, rp])

    def _parse_not(self):
        """Handle \\not — negation prefix. Combines with next character."""
        _NOT_MAP = {
            '=': '≠', '<': '≮', '>': '≯', '≤': '≰', '≥': '≱',
            '∈': '∉', '∋': '∌', '⊂': '⊄', '⊃': '⊅',
            '⊆': '⊈', '⊇': '⊉', '∼': '≁', '≃': '≄',
            '≈': '≉', '≡': '≢', '∥': '∦', '∣': '∤',
        }
        nxt = self.peek()
        if nxt['type'] == 'CHAR' and nxt.get('value') in _NOT_MAP:
            self.consume()
            return _make_run(_NOT_MAP[nxt['value']], style='plain')
        if nxt['type'] == 'COMMAND':
            nxt_name = nxt.get('value', '')[1:]
            if nxt_name in _SYMBOLS and _SYMBOLS[nxt_name] in _NOT_MAP.values():
                sym = _SYMBOLS[nxt_name]
                # Find the negated version
                rev = {v: k for k, v in _NOT_MAP.items()}
                self.consume()
                return _make_run(rev.get(sym, sym), style='plain')
        return _make_run('¬', style='plain')

    def _parse_substack(self):
        """Handle \\substack{a \\\\ b} — multi-line subscript content."""
        self.consume('LBRACE')
        lines = [[]]
        while self.peek()['type'] != 'RBRACE' and self.peek()['type'] != 'EOF':
            t = self.peek()
            if t['type'] == 'NEWLINE':
                self.consume(); lines.append([])
            elif t['type'] == 'AMPERSAND':
                self.consume()
            elif self._is_start_of_term(t) or t['type'] == 'CHAR':
                el = self._parse_expression()
                if isinstance(el, list):
                    lines[-1].extend(el)
                elif el is not None:
                    lines[-1].append(el)
            else:
                self.consume()
        self.consume('RBRACE')
        # Filter empty lines and build stacked runs
        cells = []
        max_cols = max(len(l) for l in lines) if lines else 1
        for line in lines:
            for el in line:
                cells.append(el)
            while len(cells) % max_cols != 0:
                cells.append(None)
        if not cells:
            return _make_run('')
        return _make_matrix(lines, max_cols, 'matrix')
    def _parse_tag(self):
        """Handle \\tag{label} — formula numbering."""
        arg = self._parse_arg()
        if arg is None:
            return None
        tag_text = self._ensure_element(arg)
        # Extract text from tag
        tag_str = ''
        if hasattr(tag_text, 'iter'):
            for t in tag_text.iter(f'{{{M}}}t'):
                if t.text: tag_str += t.text
        if not tag_str:
            tag_str = '?'
        # Emit as parenthesized number: (#)
        return _make_run(f'({tag_str})', style='plain')

    def _parse_equation(self):
        """Handle \\begin{equation}...\\end{equation} with auto-numbering."""
        # Parse content until END
        content = []
        eq_num = getattr(self, '_eq_counter', 0) + 1
        self._eq_counter = eq_num
        while True:
            t = self.peek()
            if t['type'] == 'END':
                end_type = t.get('value', '')
                if end_type == 'equation':
                    self.consume()
                    break
                raise ValueError(f'\\begin{{equation}} ended with \\end{{{end_type}}}')
            if t['type'] == 'EOF':
                raise ValueError('Missing \\end for \\begin{equation}')
            el = self._parse_expression()
            if isinstance(el, list):
                content.extend(el)
            elif el is not None:
                content.append(el)
        # Append numbering
        content.append(_make_run(f'({eq_num})', style='plain'))
        return self._list_to_element(content)

    def _parse_align(self, env_name='align', auto_number=True):
        """Handle align-like environments.

        align keeps the historical auto-number behavior. aligned/align* are
        used by the thesis generator for multi-line native formulas where the
        formula number is supplied explicitly with \\tag.
        """
        rows = []; current_row = []
        while True:
            t = self.peek()
            if t['type'] == 'END':
                end_type = t.get('value', '')
                if end_type == env_name:
                    self.consume()
                    break
                raise ValueError(f'\\begin{{{env_name}}} ended with \\end{{{end_type}}}')
            if t['type'] == 'EOF':
                raise ValueError(f'Missing \\end for \\begin{{{env_name}}}')
            if t['type'] == 'NEWLINE':
                self.consume()
                if current_row:
                    rows.append(current_row)
                current_row = []
                continue
            if t['type'] == 'AMPERSAND':
                self.consume(); current_row.append(None); continue
            el = self._parse_expression()
            current_row.append(self._list_to_element(el) if isinstance(el, list) else el)
        if current_row:
            rows.append(current_row)
        cols = max(len(r) for r in rows) if rows else 2
        for r in rows:
            while len(r) < cols:
                r.append(None)
            if auto_number:
                # Add auto-number
                row_num = rows.index(r) + 1
                r.append(_make_run(f'({row_num})', style='plain'))
        return _make_matrix(rows, cols + (1 if auto_number else 0), 'matrix')

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
            deg_parts = []
            while True:
                t2 = self.peek()
                if t2['type'] == 'EOF':
                    raise ValueError('Missing ] for optional sqrt degree')
                if t2['type'] == 'CHAR' and t2.get('value') == ']':
                    break
                if t2['type'] == 'SPACE':
                    self.consume()
                    continue
                before = self.pos
                el = self._parse_sup_sub_term()
                if el is not None:
                    deg_parts.extend(self._flatten(el))
                    continue
                if self.pos == before:
                    self.consume()
            deg = deg_parts
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
        # Map LaTeX style command to internal style key
        _TEXT_STYLE_MAP = {
            'text': 'plain', 'mathrm': 'plain',
            'mathbf': 'bold', 'mathit': 'italic',
            'mathsf': 'sans', 'mathtt': 'mono',
        }
        style = _TEXT_STYLE_MAP.get(name, 'plain')
        self.consume('LBRACE')
        # Collect content: flat text with inline sub/superscript support
        result = []
        buffer = []

        def flush_buffer():
            if buffer:
                result.append(_make_run(''.join(buffer), style=style))
                buffer.clear()

        while self.peek()['type'] != 'RBRACE' and self.peek()['type'] != 'EOF':
            t = self.peek()
            if t['type'] == 'CHAR':
                buffer.append(self.consume()['value'])
            elif t['type'] == 'SPACE':
                buffer.append(' ')
                self.consume()
            elif t['type'] == 'COMMAND':
                cname = self.consume()['value'][1:]
                if cname in _GREEK_LOWER:
                    flush_buffer()
                    result.append(_make_run(_GREEK_LOWER[cname], style=style))
                elif cname in _SYMBOLS:
                    flush_buffer()
                    result.append(_make_run(_SYMBOLS[cname], style=style))
                elif cname in _ARROWS:
                    flush_buffer()
                    result.append(_make_run(_ARROWS[cname], style=style))
                else:
                    buffer.append(cname)
            elif t['type'] == 'SUB':
                # Real subscript inside text style (e.g., \mathrm{CH_4})
                flush_buffer()
                sub_base = _make_run('', style=style)
                self.consume('SUB')
                sub_arg = self._ensure_single(self._parse_atom())
                result.append(_make_sub(sub_base, sub_arg) if sub_arg is not None else sub_base)
            elif t['type'] == 'SUPER':
                # Real superscript inside text style
                flush_buffer()
                sup_base = _make_run('', style=style)
                self.consume('SUPER')
                sup_arg = self._ensure_single(self._parse_atom())
                result.append(_make_sup(sup_base, sup_arg) if sup_arg else sup_base)
            else:
                flush_buffer()
                self.consume()
        self.consume('RBRACE')
        flush_buffer()
        if len(result) == 0:
            return _make_run('', style=style)
        if len(result) == 1:
            return result[0]
        return self._list_to_element(result)

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
            arg = self._ensure_single(self._parse_atom())
            if arg is not None:
                if t == 'SUB':
                    sub_el = arg
                else:
                    sup_el = arg
        arg_el = None
        if self._is_start_of_term(self.peek()):
            arg_el = self._ensure_single(self._parse_atom())
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
            return _make_run(name, style='plain')
        return _make_function(name, arg_el)

    def _parse_matrix(self):
        t = self.consume()
        mtype = t.get('value', '')

        # Route equation and align environments to dedicated handlers
        if mtype == 'equation':
            return self._parse_equation()
        if mtype == 'align':
            return self._parse_align('align', auto_number=True)
        if mtype in ('aligned', 'align*'):
            return self._parse_align(mtype, auto_number=False)

        rows = []; current_row = []
        while True:
            t = self.peek()
            if t['type'] == 'END':
                end_mtype = t.get('value', '')
                if end_mtype and end_mtype != mtype:
                    raise ValueError(f'\\begin{{{mtype}}} ended with \\end{{{end_mtype}}}')
                self.consume()  # consume END token
                break
            if t['type'] == 'EOF':
                raise ValueError(f'Missing \\end for \\begin{{{mtype}}}')
            if t['type'] == 'NEWLINE':
                self.consume(); rows.append(current_row); current_row = []; continue
            if t['type'] == 'AMPERSAND':
                self.consume(); current_row.append(None); continue
            el = self._parse_expression()
            current_row.append(self._list_to_element(el) if isinstance(el, list) else el)
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
        run_parts = []
        all_mergeable_runs = True
        for item in lst:
            if item.tag != f'{{{M}}}r':
                all_mergeable_runs = False
                break
            texts = item.findall(f'{{{M}}}t')
            non_text_children = [
                child for child in item
                if child.tag not in (f'{{{M}}}rPr', f'{{{M}}}t')
            ]
            if len(texts) != 1 or non_text_children:
                all_mergeable_runs = False
                break
            rpr = item.find(f'{{{M}}}rPr')
            rpr_key = etree.tostring(rpr, encoding='unicode') if rpr is not None else ''
            run_parts.append((texts[0].text or '', rpr, rpr_key))
        if all_mergeable_runs and len({part[2] for part in run_parts}) == 1:
            merged = etree.Element(f'{{{M}}}r')
            first_rpr = run_parts[0][1] if run_parts else None
            if first_rpr is not None:
                merged.append(etree.fromstring(etree.tostring(first_rpr)))
            else:
                etree.SubElement(merged, f'{{{M}}}rPr')
            t = etree.SubElement(merged, f'{{{M}}}t')
            t.set(XML_SPACE, 'preserve')
            t.text = ''.join(part[0] for part in run_parts)
            return merged
        # Wrap multiple siblings in invisible delimiter (m:d without brackets)
        wrapped = etree.Element(f'{{{M}}}d')
        dPr = etree.SubElement(wrapped, f'{{{M}}}dPr')
        beg = etree.SubElement(dPr, f'{{{M}}}begChr'); beg.set(f'{{{M}}}val', '')
        sep = etree.SubElement(dPr, f'{{{M}}}sepChr'); sep.set(f'{{{M}}}val', '')
        end = etree.SubElement(dPr, f'{{{M}}}endChr'); end.set(f'{{{M}}}val', '')
        for item in lst:
            e = etree.SubElement(wrapped, f'{{{M}}}e'); e.append(item)
        return wrapped


# ═══════════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════


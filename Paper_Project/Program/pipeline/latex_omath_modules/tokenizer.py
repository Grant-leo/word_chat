"""Tokenizer for the LaTeX-to-OMML converter."""
from __future__ import annotations

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
                cmd = s[i:j]
                # \begin{type} → BEGIN token
                if cmd == '\\begin':
                    k = j
                    while k < n and s[k].isspace():
                        k += 1
                    if k < n and s[k] == '{':
                        k += 1; start = k
                        while k < n and s[k] != '}':
                            k += 1
                        if k < n:
                            mtype = s[start:k]
                            tokens.append({'type': 'BEGIN', 'value': mtype})
                            i = k + 1; continue
                # \end{type} → END token
                if cmd == '\\end':
                    k = j
                    while k < n and s[k].isspace():
                        k += 1
                    if k < n and s[k] == '{':
                        k += 1; start = k
                        while k < n and s[k] != '}':
                            k += 1
                        if k < n:
                            mtype = s[start:k]
                            tokens.append({'type': 'END', 'value': mtype})
                            i = k + 1; continue
                tokens.append({'type': 'COMMAND', 'value': cmd}); i = j
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


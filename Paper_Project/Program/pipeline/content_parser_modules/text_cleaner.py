"""Shared text cleanup helpers for content parsing."""
import re


_NOISE_TEXT = {
    '\u590d\u5236',
    'Copy',
    'Plain Text',
    '\u7eaf\u6587\u672c',
}


def clean_markdown_links(text):
    def repl(match):
        label = (match.group(1) or '').strip()
        target = (match.group(2) or '').strip()
        return label or target

    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', repl, str(text or ''))


def is_noise_text(text):
    value = str(text or '').strip()
    return value in _NOISE_TEXT


def clean_text_artifacts(text, preserve_newlines=False):
    """Remove editor/clipboard artifacts without changing content semantics."""
    value = clean_markdown_links(text)
    value = value.replace('\u00a0', ' ')
    if preserve_newlines:
        lines = []
        for line in value.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
            cleaned = re.sub(r'[ \t]+', ' ', line).strip()
            if is_noise_text(cleaned):
                continue
            lines.append(cleaned)
        return '\n'.join(lines).strip()
    value = re.sub(r'\s+', ' ', value).strip()
    return '' if is_noise_text(value) else value


def clean_code_text(text):
    return clean_text_artifacts(text, preserve_newlines=True)

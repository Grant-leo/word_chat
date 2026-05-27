"""Reference-section collection state for content_parser.py."""
import re
from typing import Any, Callable, Dict, List, Optional


def is_reference_heading(text: str) -> bool:
    value = str(text or '').strip()
    return bool(re.match(r'(?i)^references?\b', value) or value.startswith('\u53c2\u8003\u6587\u732e'))


class ReferenceCollector:
    """Collect reference entries until a back-matter heading exits the section."""

    def __init__(
        self,
        clean_text_func: Callable[..., str],
        is_backmatter_heading_func: Callable[[str], bool],
        normalize_heading_spacing_func: Callable[[str], str],
        classify_section_role_func: Callable[[str, int], str],
        table_rows_look_like_code_func: Callable[[List[List[str]]], bool],
        code_text_from_table_rows_func: Callable[..., str],
        clean_code_func: Callable[[str], str],
    ):
        self.clean_text_func = clean_text_func
        self.is_backmatter_heading_func = is_backmatter_heading_func
        self.normalize_heading_spacing_func = normalize_heading_spacing_func
        self.classify_section_role_func = classify_section_role_func
        self.table_rows_look_like_code_func = table_rows_look_like_code_func
        self.code_text_from_table_rows_func = code_text_from_table_rows_func
        self.clean_code_func = clean_code_func
        self.active_section: Optional[Dict[str, Any]] = None
        self.collected: List[Any] = []

    @property
    def active(self) -> bool:
        return self.active_section is not None

    def start_if_heading(self, text: str) -> bool:
        if is_reference_heading(text):
            self.active_section = {'heading': str(text or '').strip(), 'entries': []}
            return True
        return False

    def _flush_active(self) -> None:
        if self.active_section and self.active_section.get('entries'):
            self.collected.extend(self.active_section['entries'])
        self.active_section = None

    def exit_to_backmatter_section(self, text: str, level: int) -> Optional[Dict[str, Any]]:
        if self.active_section is None or not self.is_backmatter_heading_func(text):
            return None
        self._flush_active()
        heading = self.normalize_heading_spacing_func(re.split(r'[:\uff1a]', str(text or ''), maxsplit=1)[0].strip())
        section_level = level or 1
        return {
            'heading': heading,
            'level': section_level,
            'role': self.classify_section_role_func(heading, section_level),
            'paragraphs': [],
            'images': [],
        }

    def consume_text(self, text: str) -> bool:
        if self.active_section is None:
            return False
        clean = self.clean_text_func(text)
        if clean:
            self.active_section['entries'].append(clean)
        return True

    def consume_table_rows(self, rows: List[List[str]]) -> bool:
        if self.active_section is None:
            return False
        if not rows:
            return True
        if self.table_rows_look_like_code_func(rows):
            self.active_section['entries'].append({
                'role': 'code',
                'code': self.code_text_from_table_rows_func(rows, clean_code_func=self.clean_code_func),
                'table_rows': rows,
            })
        else:
            self.active_section['entries'].append({'role': 'table', 'table_rows': rows})
        return True

    def finish(self) -> List[Any]:
        self._flush_active()
        return list(self.collected)

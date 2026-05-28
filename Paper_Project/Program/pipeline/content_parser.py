"""Stable DOCX content extraction entry point."""
from __future__ import annotations

import json
import os

try:
    from content_parser_modules.extractor import _content_placeholder_samples, _count_structured_body_tables, extract
except ImportError:  # pragma: no cover - package-style imports
    from .content_parser_modules.extractor import _content_placeholder_samples, _count_structured_body_tables, extract

__all__ = ["_content_placeholder_samples", "_count_structured_body_tables", "extract"]

if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Templates/模版.docx'
    content = extract(path)
    json_path = os.path.splitext(path)[0] + '_content.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    print(f'Content JSON → {json_path}')
    print(f'Sections: {len(content["sections"])}  References: {len(content["references"])}  Images: {content["_meta"]["images_extracted"]}')

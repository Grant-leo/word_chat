"""Generated build-script runtime template assembly."""
from __future__ import annotations

from .runtime_base import BASE_RUNTIME
from .runtime_body import BODY_RUNTIME
from .runtime_build import BUILD_RUNTIME
from .runtime_content_helpers import CONTENT_HELPERS_RUNTIME
from .runtime_cover import COVER_RUNTIME
from .runtime_formula import FORMULA_RUNTIME
from .runtime_formula_render import FORMULA_RENDER_RUNTIME
from .runtime_formula_text import FORMULA_TEXT_RUNTIME
from .runtime_front_matter import FRONT_MATTER_RUNTIME
from .runtime_media_tables import MEDIA_TABLE_RUNTIME
from .runtime_notes import NOTES_RUNTIME
from .runtime_references import REFERENCES_RUNTIME
from .runtime_toc import TOC_RUNTIME

RUNTIME_TEMPLATE = r'''
__BASE_RUNTIME__

__COVER_RUNTIME__

__NOTES_RUNTIME__

__FORMULA_RUNTIME__

__TOC_RUNTIME__

__FRONT_MATTER_RUNTIME__

__CONTENT_HELPERS_RUNTIME__

__FORMULA_TEXT_RUNTIME__

__FORMULA_RENDER_RUNTIME__

__MEDIA_TABLE_RUNTIME__

__REFERENCES_RUNTIME__

__BODY_RUNTIME__

__BUILD_RUNTIME__
'''

RUNTIME_TEMPLATE = (
    RUNTIME_TEMPLATE
    .replace('__BASE_RUNTIME__', BASE_RUNTIME)
    .replace('__COVER_RUNTIME__', COVER_RUNTIME)
    .replace('__NOTES_RUNTIME__', NOTES_RUNTIME)
    .replace('__FORMULA_RUNTIME__', FORMULA_RUNTIME)
    .replace('__CONTENT_HELPERS_RUNTIME__', CONTENT_HELPERS_RUNTIME)
    .replace('__FORMULA_TEXT_RUNTIME__', FORMULA_TEXT_RUNTIME)
    .replace('__FORMULA_RENDER_RUNTIME__', FORMULA_RENDER_RUNTIME)
    .replace('__MEDIA_TABLE_RUNTIME__', MEDIA_TABLE_RUNTIME)
    .replace('__REFERENCES_RUNTIME__', REFERENCES_RUNTIME)
    .replace('__TOC_RUNTIME__', TOC_RUNTIME)
    .replace('__FRONT_MATTER_RUNTIME__', FRONT_MATTER_RUNTIME)
    .replace('__BODY_RUNTIME__', BODY_RUNTIME)
    .replace('__BUILD_RUNTIME__', BUILD_RUNTIME)
)

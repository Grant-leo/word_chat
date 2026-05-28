"""
regression_suite.py - synthetic regression checks for the Word pipeline.

The suite creates temporary, non-private DOCX/MD fixtures and verifies the
engine behavior that is hard to cover by manual inspection alone:

- inline math stays in the current paragraph
- display math remains native OMML display math
- mixed text/image/math paragraphs do not drop tokens
- markdown tables/code/math keep structure
- build_manifest.json drives body element QA counts
- template profiles avoid private source filenames
- visual QA fails closed when required render tools are unavailable
"""
from __future__ import annotations

import sys
from pathlib import Path

PIPELINE_DIR = Path(__file__).resolve().parent
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from regression_suite_modules.harness import main
import regression_suite_modules.pipeline_cases  # noqa: F401 - registers pipeline cases
import regression_suite_modules.content_cases  # noqa: F401 - registers content parser cases
import regression_suite_modules.formula_cases  # noqa: F401 - registers formula cases
import regression_suite_modules.md_cases  # noqa: F401 - registers Markdown cases
import regression_suite_modules.qa_cases  # noqa: F401 - registers QA cases
import regression_suite_modules.generator_cases  # noqa: F401 - registers script generator cases
import regression_suite_modules.template_format_cases  # noqa: F401 - registers template/format cases
import regression_suite_modules.operational_cases  # noqa: F401 - registers privacy/visual/CLI cases

if __name__ == "__main__":
    main()

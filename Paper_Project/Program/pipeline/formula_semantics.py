"""
formula_semantics.py - lightweight semantic guards for formula extraction.

The first stage is deliberately deterministic: it separates plain quantities
from standalone equations and identifies paragraphs where narrative text was
captured as a formula. Heavier OCR/ML extraction can build on these labels
without changing the content.json contract again.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
import re
from typing import Dict


CATEGORY_TEXT = "TEXT"
CATEGORY_INLINE_MATH = "INLINE_MATH"
CATEGORY_DISPLAY_MATH = "DISPLAY_MATH"
CATEGORY_QUANTITY_TEXT = "QUANTITY_TEXT"
CATEGORY_UNIT_TEXT = "UNIT_TEXT"
CATEGORY_CITATION = "CITATION"
CATEGORY_FORMULA_LABEL = "FORMULA_LABEL"
CATEGORY_CONTAMINATED = "CONTAMINATED"


_TRANSLATION = str.maketrans(
    {
        "\uff1d": "=",
        "\uff0b": "+",
        "\uff0d": "-",
        "\u2212": "-",
        "\uff0a": "*",
        "\u00d7": "*",
        "\u00b7": "*",
        "\uff0f": "/",
        "\u00f7": "/",
        "\uff1c": "<",
        "\uff1e": ">",
        "\uff05": "%",
        "\uff08": "(",
        "\uff09": ")",
        "\uff0c": ",",
        "\u3002": ".",
        "\uff1a": ":",
        "\uff1b": ";",
    }
)

_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
_DIGIT_RE = re.compile(r"\d")
_LATEX_DELIMITED_RE = re.compile(r"^\s*\${1,2}.+\${1,2}\s*$", re.S)
_CITATION_RE = re.compile(r"^\s*(?:\[\d+(?:[-,\s\u2013\u2014]\d+)*\]|\(\d{4}[a-z]?\))\s*$")
_FORMULA_LABEL_RE = re.compile(r"^\s*[\(\uff08]\s*\d+(?:\s*[-.]\s*\d+)?\s*[\)\uff09]\s*$")
_RELATION_RE = re.compile(r"[=<>]|\u2264|\u2265|\u2260|\u2248|\u223c|\u221d|\u2208|\u2209")
_MATH_OP_RE = re.compile(r"[+\-*/%]|\u2211|\u220f|\u222b|\u221a|\^|_")
_SENTENCE_PUNCT_RE = re.compile(r"[.;!?]|\uff1b|\uff01|\uff1f")
_NARRATIVE_WORD_RE = re.compile(
    "|".join(
        [
            r"\u5bf9\u5e94",
            r"\u7ea6\u675f",
            r"\u5f53.+\u65f6",
            r"\u65e0\u6cd5",
            r"\u4f4e\u4e8e",
            r"\u7ea6\u4e3a",
            r"\u89d2\u5ea6",
            r"\u6a21\u5f0f",
            r"\u6536\u5165",
            r"\u6210\u672c",
            r"\u4ea7\u91cf",
            r"\u5f00\u673a",
            r"\u7ed3\u679c",
            r"\u8868\u660e",
            r"\u5206\u6790",
            r"\u9009\u62e9",
        ]
    )
)
_UNIT_RE = re.compile(
    r"(?i)\b(?:"
    r"kw|mw|gw|kwh|mwh|gwh|v|kv|a|ma|hz|khz|mhz|w|j|kj|mj|n|pa|kpa|mpa|"
    r"kg|g|mg|t|ton|h|min|s|ms|m|cm|mm|km|m2|m3|m\^2|m\^3|"
    r"yuan|rmb|usd|cny"
    r")\b|[%\u2103\u5143\u4e07\u5143\u5428\u65e5\u5e74\u6708\u5468\u5c0f\u65f6]"
)


@dataclass(frozen=True)
class FormulaSemanticResult:
    category: str
    confidence: float
    reason: str
    should_number: bool = False

    def to_dict(self) -> Dict[str, object]:
        return asdict(self)


@dataclass(frozen=True)
class FormulaSpan:
    start: int
    end: int
    text: str
    category: str
    confidence: float
    reason: str
    latex: str = ""

    def to_dict(self) -> Dict[str, object]:
        return asdict(self)


def normalize_math_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").translate(_TRANSLATION)).strip()


def _cjk_count(text: str) -> int:
    return len(_CJK_RE.findall(text))


def _has_relation(text: str) -> bool:
    return bool(_RELATION_RE.search(normalize_math_text(text)))


def _has_math_operator(text: str) -> bool:
    return bool(_MATH_OP_RE.search(normalize_math_text(text)))


def is_citation_text(text: str) -> bool:
    return bool(_CITATION_RE.match(str(text or "").strip()))


def is_formula_label(text: str) -> bool:
    return bool(_FORMULA_LABEL_RE.match(str(text or "").strip()))


def is_quantity_text(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or not _DIGIT_RE.search(t):
        return False
    if _has_relation(t):
        return False
    if _has_math_operator(t) and not re.search(r"\d\s*-\s*\d", t):
        return False
    return bool(_UNIT_RE.search(t))


def formula_text_looks_contaminated(text: str) -> bool:
    t = normalize_math_text(text)
    if not t:
        return False
    cjk = _cjk_count(t)
    if cjk < 4:
        return False
    has_formula_signal = bool(
        _has_relation(t)
        or re.search(r"[A-Za-z\u0370-\u03ff][A-Za-z0-9_\u0370-\u03ff]*\s*\(", t)
    )
    if not has_formula_signal or not _DIGIT_RE.search(t):
        return False
    has_sentence_punct = bool(_SENTENCE_PUNCT_RE.search(t))
    has_narrative_word = bool(_NARRATIVE_WORD_RE.search(t))
    if len(t) > 30 and has_sentence_punct and has_narrative_word:
        return True
    if len(t) > 70 and cjk > 12 and has_sentence_punct:
        return True
    if len(t) > 90 and cjk > 18:
        return True
    return False


def is_formula_problem_text(text: str) -> bool:
    """True only for short, dense formula-like text that should block QA."""
    t = normalize_math_text(text)
    if not formula_text_looks_contaminated(t):
        return False
    if len(t) > 180:
        return False
    if str(text or "").strip().endswith(("\u3002", "\uff1b", ";", ".")) and not re.search(r"[=<>]\s*$|\u2264\s*$|\u2265\s*$|\u2260\s*$|\u2248\s*$|\u2208\s*$|\u2209\s*$", t):
        return False
    if split_inline_math_spans(text):
        return False
    math_marks = len(re.findall(r"[=<>+\-*/%]|\u2264|\u2265|\u2260|\u2248|\u2208|\u2209", t))
    cjk = _cjk_count(t)
    if re.search(r"[=<>]|\u2264|\u2265|\u2260|\u2248|\u2208|\u2209", t[-3:]):
        return True
    return bool(math_marks >= 3 and cjk <= 20 and len(t) <= 150)


def formula_should_number(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or formula_text_looks_contaminated(t):
        return False
    if not _DIGIT_RE.search(t):
        return False
    return bool(_has_relation(t) and _has_math_operator(t))


def looks_like_formula_text(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 180:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if _LATEX_DELIMITED_RE.match(str(text or "").strip()):
        return True
    if formula_text_looks_contaminated(t):
        return False
    cjk = _cjk_count(t)
    if (cjk > 3 and _SENTENCE_PUNCT_RE.search(t)) or cjk > 8:
        return False
    starts_continuation = bool(re.match(r"^[=<>]", t)) and bool(_DIGIT_RE.search(t))
    if starts_continuation:
        return True
    if _has_relation(t) and (_has_math_operator(t) or _DIGIT_RE.search(t)):
        return True
    return False


def classify_formula_text(text: str) -> FormulaSemanticResult:
    raw = str(text or "").strip()
    t = normalize_math_text(raw)
    if not t:
        return FormulaSemanticResult(CATEGORY_TEXT, 1.0, "empty")
    if is_citation_text(t):
        return FormulaSemanticResult(CATEGORY_CITATION, 0.98, "citation marker")
    if is_formula_label(t):
        return FormulaSemanticResult(CATEGORY_FORMULA_LABEL, 0.98, "standalone formula label")
    if formula_text_looks_contaminated(t):
        return FormulaSemanticResult(CATEGORY_CONTAMINATED, 0.92, "narrative text mixed with math operators")
    if is_quantity_text(t):
        return FormulaSemanticResult(CATEGORY_QUANTITY_TEXT, 0.86, "quantity/unit expression without equation relation")
    if _LATEX_DELIMITED_RE.match(raw):
        return FormulaSemanticResult(CATEGORY_DISPLAY_MATH, 0.96, "latex delimiter", should_number=formula_should_number(t))
    if looks_like_formula_text(t):
        return FormulaSemanticResult(CATEGORY_DISPLAY_MATH, 0.88, "standalone equation", should_number=formula_should_number(t))
    return FormulaSemanticResult(CATEGORY_TEXT, 0.74, "plain text")


_DOLLAR_INLINE_RE = re.compile(r"(?<!\$)\$(?!\$)([^$\n]{1,120})(?<!\$)\$(?!\$)")
_INLINE_EQUATION_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"("
    r"[A-Za-z\u0370-\u03ff][A-Za-z0-9_\u0370-\u03ff]*"
    r"(?:\s*\([^,;\n\u3002\uff0c\uff1b\uff1a]{0,40}\))?"
    r"\s*(?:<=|>=|=|<|>|\u2264|\u2265|\u2248|\u2260)\s*"
    r"[^,;\n\u3002\uff0c\uff1b\uff1a]{1,100}"
    r")"
)


def _valid_inline_math_candidate(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 140:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if formula_text_looks_contaminated(t):
        return False
    if _cjk_count(t) > 0:
        return False
    for left, right in (("(", ")"), ("{", "}"), ("[", "]")):
        if t.count(left) != t.count(right):
            return False
    if re.search(r"[\u3002\uff0c\uff1b\uff1a,;]|\s[\u4e00-\u9fff]", str(text or "")):
        return False
    if not _has_relation(t):
        return False
    if _has_math_operator(t) or _DIGIT_RE.search(t):
        return True
    parts = re.split(r"<=|>=|=|<|>|\u2264|\u2265|\u2248|\u2260", t, maxsplit=1)
    return len(parts) == 2 and all(re.search(r"[A-Za-z\u0370-\u03ff]", p) for p in parts)


def _valid_dollar_inline_math(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 120:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if _cjk_count(t) > 0:
        return False
    for left, right in (("(", ")"), ("{", "}"), ("[", "]")):
        if t.count(left) != t.count(right):
            return False
    if re.search(r"\\[A-Za-z]+|[_^]|[=<>]|\u2264|\u2265|\u2248|\u2260|[+\-*/%]|\u2211|\u222b|\u221a", t):
        return True
    if re.fullmatch(r"[A-Za-z\u0370-\u03ff][A-Za-z0-9\u0370-\u03ff]*(?:\([A-Za-z0-9]+\))?", t):
        return True
    return False


def _trim_inline_candidate(text: str) -> str:
    t = str(text or "").strip()
    while t and t[-1] in ")]\uff09\uff3d":
        normalized = normalize_math_text(t)
        if normalized.count(")") <= normalized.count("(") and normalized.count("]") <= normalized.count("["):
            break
        t = t[:-1].rstrip()
    return t


def split_inline_math_spans(text: str) -> list[Dict[str, object]]:
    """Return conservative inline math spans from a normal prose paragraph."""
    raw = str(text or "")
    if not raw or len(raw) > 1200:
        return []
    if re.match(r"^\s*\$\$.+\$\$\s*$", raw, re.S):
        return []
    spans: list[FormulaSpan] = []

    def overlaps(start: int, end: int) -> bool:
        return any(not (end <= span.start or start >= span.end) for span in spans)

    for match in _DOLLAR_INLINE_RE.finditer(raw):
        inner = match.group(1).strip()
        if not inner or overlaps(match.start(), match.end()) or not _valid_dollar_inline_math(inner):
            continue
        spans.append(
            FormulaSpan(
                match.start(),
                match.end(),
                inner,
                CATEGORY_INLINE_MATH,
                0.96,
                "dollar-delimited inline math",
                latex=inner,
            )
        )

    for match in _INLINE_EQUATION_RE.finditer(raw):
        candidate = _trim_inline_candidate(match.group(1))
        start = match.start(1) + (len(match.group(1)) - len(match.group(1).lstrip()))
        end = start + len(candidate)
        if overlaps(start, end) or not _valid_inline_math_candidate(candidate):
            continue
        spans.append(
            FormulaSpan(
                start,
                end,
                candidate,
                CATEGORY_INLINE_MATH,
                0.84,
                "inline equation span",
            )
        )

    spans.sort(key=lambda span: span.start)
    return [span.to_dict() for span in spans]

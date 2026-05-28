"""Regex patterns and category constants for formula semantics."""
from __future__ import annotations

import re

CATEGORY_TEXT = "TEXT"
CATEGORY_INLINE_MATH = "INLINE_MATH"
CATEGORY_DISPLAY_MATH = "DISPLAY_MATH"
CATEGORY_QUANTITY_TEXT = "QUANTITY_TEXT"
CATEGORY_UNIT_TEXT = "UNIT_TEXT"
CATEGORY_CITATION = "CITATION"
CATEGORY_FORMULA_LABEL = "FORMULA_LABEL"
CATEGORY_CONTAMINATED = "CONTAMINATED"

TRANSLATION = str.maketrans(
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

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
DIGIT_RE = re.compile(r"\d")
LATEX_DELIMITED_RE = re.compile(r"^\s*\${1,2}.+\${1,2}\s*$", re.S)
CITATION_RE = re.compile(r"^\s*(?:\[\d+(?:[-,\s\u2013\u2014]\d+)*\]|\(\d{4}[a-z]?\))\s*$")
FORMULA_LABEL_RE = re.compile(r"^\s*[\(\uff08]\s*\d+(?:\s*[-.]\s*\d+)?\s*[\)\uff09]\s*$")
RELATION_RE = re.compile(r"[=<>]|\u2264|\u2265|\u2260|\u2248|\u223c|\u221d|\u2208|\u2209")
MATH_OP_RE = re.compile(r"[+\-*/%]|\u2211|\u220f|\u222b|\u221a|\^|_")
SENTENCE_PUNCT_RE = re.compile(r"[.;!?]|\uff1b|\uff01|\uff1f")
NARRATIVE_WORD_RE = re.compile(
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
UNIT_RE = re.compile(
    r"(?i)\b(?:"
    r"kw|mw|gw|kwh|mwh|gwh|v|kv|a|ma|hz|khz|mhz|w|j|kj|mj|n|pa|kpa|mpa|"
    r"kg|g|mg|t|ton|h|min|s|ms|m|cm|mm|km|m2|m3|m\^2|m\^3|"
    r"yuan|rmb|usd|cny"
    r")\b|[%\u2103\u5143\u4e07\u5143\u5428\u65e5\u5e74\u6708\u5468\u5c0f\u65f6]"
)

DOLLAR_INLINE_RE = re.compile(r"(?<!\$)\$(?!\$)([^$\n]{1,120})(?<!\$)\$(?!\$)")
INLINE_EQUATION_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"("
    r"[A-Za-z\u0370-\u03ff][A-Za-z0-9_\u0370-\u03ff]*"
    r"(?:\s*\([^,;\n\u3002\uff0c\uff1b\uff1a]{0,40}\))?"
    r"\s*(?:<=|>=|=|<|>|\u2264|\u2265|\u2248|\u2260)\s*"
    r"[^,;\n\u3002\uff0c\uff1b\uff1a]{1,100}"
    r")"
)

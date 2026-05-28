"""Data contracts for formula semantic classification."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Dict


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

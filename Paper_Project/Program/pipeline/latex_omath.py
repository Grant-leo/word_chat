"""Stable public entry point for LaTeX math to OOXML Math conversion."""
from __future__ import annotations

from latex_omath_modules.api import body_latex, formula_text_from_omath, latex_to_omath
from latex_omath_modules.parser import _LaTeXParser
from latex_omath_modules.tokenizer import _tokenize

__all__ = [
    "_LaTeXParser",
    "_tokenize",
    "body_latex",
    "formula_text_from_omath",
    "latex_to_omath",
]


def _run_self_test():
    tests = [
        (r"x^2", "superscript"),
        (r"x_1", "subscript"),
        (r"x_1^2", "subsuperscript"),
        (r"\frac{a}{b}", "fraction"),
        (r"\binom{n}{k}", "binomial"),
        (r"\sqrt{x}", "sqrt"),
        (r"\sqrt[3]{x}", "nth root"),
        (r"\alpha + \beta = \gamma", "greek"),
        (r"\sin\theta + \cos\theta", "trig"),
        (r"\Gamma(z)", "Gamma fn"),
        (r"\infty + \partial", "symbols"),
        (r"\int_0^\infty e^{-x^2} dx", "integral"),
        (r"\sum_{i=1}^n x_i", "summation"),
        (r"\left(\frac{a}{b}\right)^n", "delimited frac"),
        (r"\hat{x} + \bar{y} + \vec{v}", "accents"),
        (r"\overline{abc} + \underline{xyz}", "over/underline"),
        (r"\boxed{E=mc^2}", "boxed"),
        (r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", "pmatrix"),
        (r"\begin{cases} x & x>0 \\ -x & x\leq 0 \end{cases}", "cases"),
        (r"\overbrace{x+y}^{n}", "overbrace"),
        (r"\underbrace{x+y}_{n}", "underbrace"),
        (r"\text{hello world}", "text mode"),
        (r"\lim_{x\to\infty} f(x)", "limit"),
        (r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}", "quadratic"),
        (r"\mathrm{CH_4 + 2O_2 \to CO_2 + 2H_2O}", "chemical"),
        (r"\Gamma(z) = \int_0^\infty t^{z-1} e^{-t} dt", "Gamma integral"),
        (r"\forall x \in \mathbb{R}, \exists y > 0", "quantifiers"),
    ]

    passed = 0
    for latex, desc in tests:
        try:
            xml = latex_to_omath(latex)
            if "[LaTeX error" in xml:
                print(f"FAIL  | {desc:20s} | error embedded in output")
            else:
                print(f"OK    | {desc:20s}")
                passed += 1
        except Exception as exc:
            print(f"CRASH | {desc:20s} | {type(exc).__name__}: {exc}")

    print(f"\n{passed}/{len(tests)} tests passed")


if __name__ == "__main__":
    _run_self_test()

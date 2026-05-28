"""Symbol registries for the LaTeX to OOXML Math converter."""
from __future__ import annotations

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_GREEK_LOWER = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ε", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ",
    "sigma": "σ", "varsigma": "ς", "tau": "τ", "upsilon": "υ",
    "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ",
    "omega": "ω",
}

_GREEK_UPPER = {
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Phi": "Φ",
    "Psi": "Ψ", "Omega": "Ω",
}

_SYMBOLS = {
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "·", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱",
    "ldots": "…", "forall": "∀", "exists": "∃", "nexists": "∄",
    "neg": "¬", "lnot": "¬", "wedge": "∧", "land": "∧",
    "vee": "∨", "lor": "∨", "cap": "∩", "cup": "∪",
    "subset": "⊂", "supset": "⊃", "subseteq": "⊆", "supseteq": "⊇",
    "in": "∈", "notin": "∉", "ni": "∋",
    "approx": "≈", "equiv": "≡", "neq": "≠", "ne": "≠",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥",
    "ll": "≪", "gg": "≫", "propto": "∝", "sim": "∼",
    "simeq": "≃", "cong": "≅", "doteq": "≐",
    "perp": "⟂", "parallel": "∥",
    "angle": "∠", "measuredangle": "∡",
    "circ": "∘", "bullet": "∙",
    "oplus": "⊕", "ominus": "⊖", "otimes": "⊗", "odot": "⊙",
    "oslash": "⊘", "uplus": "⊎",
    "aleph": "ℵ", "hbar": "ℏ", "ell": "ℓ",
    "wp": "℘", "Re": "ℜ", "Im": "ℑ",
    "emptyset": "∅", "varnothing": "∅",
    "top": "⊤", "bot": "⊥",
    "triangle": "△", "triangledown": "▽",
    "square": "□", "Box": "□", "diamond": "◇", "Diamond": "◇",
    "star": "⋆", "bigstar": "★",
    "clubsuit": "♣", "diamondsuit": "♦",
    "heartsuit": "♡", "spadesuit": "♠",
    "prime": "′", "backslash": "∖",
    "surd": "√", "dag": "†", "ddag": "‡",
    "S": "§", "P": "¶", "pounds": "£",
    "subsetneq": "⊊", "supsetneq": "⊋",
    "therefore": "∴", "because": "∵",
    "implies": "⟹", "iff": "⟺", "impliedby": "⟸",
    "nsubseteq": "⊈", "nsupseteq": "⊉", "setminus": "∖", "complement": "∁",
    "vdash": "⊢", "dashv": "⊣", "models": "⊧",
    "lesssim": "≲", "gtrsim": "≳", "approxeq": "≊",
    "triangleq": "≜", "circeq": "≗",
    "imath": "ı", "jmath": "ȷ", "hslash": "ℏ",
    "bigoplus": "⨁", "bigotimes": "⨂", "bigodot": "⨀", "coprod": "∐",
    "rtimes": "⋊", "ltimes": "⋋",
    "bigsqcup": "⊔", "cdotp": "·",
    "mid": "∣", "nmid": "∤",
    "smile": "⌣", "frown": "⌢",
    "wr": "≀", "amalg": "⨿",
    "lhd": "⊲", "rhd": "⊳", "unlhd": "⊴", "unrhd": "⊵",
    "twoheadrightarrow": "↠", "twoheadleftarrow": "↞",
    "looparrowleft": "↫", "looparrowright": "↬",
    "curvearrowleft": "↶", "curvearrowright": "↷",
    "rightleftharpoons": "⇌", "leftrightharpoons": "⇋",
    "Lsh": "↰", "Rsh": "↱",
    "sphericalangle": "∢",
    "Bbbk": "𝕜",
}

_DBLSTRUCK = {
    "A": "𝔸", "B": "𝔹", "C": "ℂ", "D": "𝔻", "E": "𝔼",
    "F": "𝔽", "G": "𝔾", "H": "ℍ", "I": "𝕀", "J": "𝕁",
    "K": "𝕂", "L": "𝕃", "M": "𝕄", "N": "ℕ", "O": "𝕆",
    "P": "ℙ", "Q": "ℚ", "R": "ℝ", "S": "𝕊", "T": "𝕋",
    "U": "𝕌", "V": "𝕍", "W": "𝕎", "X": "𝕏", "Y": "𝕐", "Z": "ℤ",
    "a": "𝕒", "b": "𝕓", "c": "𝕔", "d": "𝕕", "e": "𝕖",
    "f": "𝕗", "g": "𝕘", "h": "𝕙", "i": "𝕚", "j": "𝕛",
    "k": "𝕜", "l": "𝕝", "m": "𝕞", "n": "𝕟", "o": "𝕠",
    "p": "𝕡", "q": "𝕢", "r": "𝕣", "s": "𝕤", "t": "𝕥",
    "u": "𝕦", "v": "𝕧", "w": "𝕨", "x": "𝕩", "y": "𝕪", "z": "𝕫",
    "0": "𝟘", "1": "𝟙", "2": "𝟚", "3": "𝟛", "4": "𝟜",
    "5": "𝟝", "6": "𝟞", "7": "𝟟", "8": "𝟠", "9": "𝟡",
    "gamma": "ℾ", "Gamma": "ℾ", "pi": "ℿ", "Pi": "ℿ",
    "Sigma": "⅀",
}

_SCRIPT = {
    "A": "𝒜", "B": "ℬ", "C": "𝒞", "D": "𝒟", "E": "ℰ",
    "F": "ℱ", "G": "𝒢", "H": "ℋ", "I": "ℐ", "J": "𝒥",
    "K": "𝒦", "L": "ℒ", "M": "ℳ", "N": "𝒩", "O": "𝒪",
    "P": "𝒫", "Q": "𝒬", "R": "ℛ", "S": "𝒮", "T": "𝒯",
    "U": "𝒰", "V": "𝒱", "W": "𝒲", "X": "𝒳", "Y": "𝒴", "Z": "𝒵",
}

_ARROWS = {
    "to": "→", "rightarrow": "→", "longrightarrow": "⟶",
    "Rightarrow": "⇒", "Longrightarrow": "⟹",
    "leftarrow": "←", "longleftarrow": "⟵",
    "Leftarrow": "⇐", "Longleftarrow": "⟸",
    "leftrightarrow": "↔", "longleftrightarrow": "⟷",
    "Leftrightarrow": "⇔", "Longleftrightarrow": "⟺",
    "mapsto": "↦", "longmapsto": "⟼",
    "hookrightarrow": "↪", "hookleftarrow": "↩",
    "rightharpoonup": "⇀", "rightharpoondown": "⇁",
    "leftharpoonup": "↼", "leftharpoondown": "↽",
    "uparrow": "↑", "downarrow": "↓", "updownarrow": "↕",
    "Uparrow": "⇑", "Downarrow": "⇓", "Updownarrow": "⇕",
    "nearrow": "↗", "searrow": "↘", "swarrow": "↙", "nwarrow": "↖",
    "leadsto": "↝",
}

_NAMED_FUNCTIONS = frozenset([
    "sin", "cos", "tan", "cot", "sec", "csc",
    "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "coth",
    "ln", "log", "lg", "exp", "det", "gcd",
    "max", "min", "lim", "limsup", "liminf",
    "arg", "deg", "dim", "hom", "ker",
    "sup", "inf", "Pr",
])

_ACCENT_CHARS = {
    "hat": "̂", "bar": "̄", "vec": "⃗",
    "dot": "̇", "ddot": "̈", "tilde": "̃",
}

_NARY_CHARS = {
    "sum": "∑", "int": "∫", "prod": "∏",
    "iint": "∬", "iiint": "∭", "oint": "∮",
    "bigcup": "⋃", "bigcap": "⋂",
}

_NARY_LIMLOC = {
    "sum": "undOvr", "prod": "undOvr", "bigcup": "undOvr", "bigcap": "undOvr",
    "int": "subSup", "iint": "subSup", "iiint": "subSup", "oint": "subSup",
}

_LEFT_DELIM = {
    "\\{": "{", "\\{.": "{",
    "\\langle": "⟨", "\\lfloor": "⌊", "\\lceil": "⌈",
    "\\|": "‖", ".": "",
}
_RIGHT_DELIM = {
    "\\}": "}", "\\.}": "}",
    "\\rangle": "⟩", "\\rfloor": "⌋", "\\rceil": "⌉",
    "\\|": "‖", ".": "",
}

_MATRIX_BRACKETS = {
    "matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"),
    "Vmatrix": ("‖", "‖"), "cases": ("{", ""),
}

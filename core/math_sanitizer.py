# -*- coding: utf-8 -*-
"""
core/math_sanitizer.py
Converts raw LaTeX math expressions and syntax into clean, professional Unicode text.
Eliminates stray dollar signs, backslashes, LaTeX commands, and markdown bold artifacts.
"""

import re

SUPERSCRIPTS = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
    'n': 'ⁿ', 'x': 'ˣ', 'y': 'ʸ', 'i': 'ⁱ', 't': 'ᵗ'
}

SUBSCRIPTS = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
    'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ',
    'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
    'p': 'ₚ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ',
    'v': 'ᵥ', 'x': 'ₓ'
}

GREEK_LETTERS = {
    r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
    r'\\epsilon': 'ε', r'\\varepsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η',
    r'\\theta': 'θ', r'\\vartheta': 'ϑ', r'\\iota': 'ι', r'\\kappa': 'κ',
    r'\\lambda': 'λ', r'\\mu': 'μ', r'\\nu': 'ν', r'\\xi': 'ξ',
    r'\\pi': 'π', r'\\varpi': 'ϖ', r'\\rho': 'ρ', r'\\varrho': 'ϱ',
    r'\\sigma': 'σ', r'\\varsigma': 'ς', r'\\tau': 'τ', r'\\upsilon': 'υ',
    r'\\phi': 'φ', r'\\varphi': 'ϕ', r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
    r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
    r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Upsilon': 'Υ',
    r'\\Phi': 'Φ', r'\\Psi': 'Ψ', r'\\Omega': 'Ω'
}

MATH_SYMBOLS = {
    r'\\times': '×', r'\\cdot': '·', r'\\bullet': '•',
    r'\\approx': '≈', r'\\sim': '~', r'\\equiv': '≡',
    r'\\leq': '≤', r'\\le\b': '≤', r'\\geq': '≥', r'\\ge\b': '≥',
    r'\\neq': '≠', r'\\ne\b': '≠', r'\\pm': '±', r'\\mp': '∓',
    r'\\infty': '∞', r'\\propto': '∝', r'\\forall': '∀', r'\\exists': '∃',
    r'\\in\b': '∈', r'\\notin': '∉', r'\\subset': '⊂', r'\\subseteq': '⊆',
    r'\\cup': '∪', r'\\cap': '∩', r'\\emptyset': '∅',
    r'\\rangle': '⟩', r'\\langle': '⟨',
    r'\\rightarrow': '→', r'\\to\b': '→', r'\\leftarrow': '←',
    r'\\Rightarrow': '⇒', r'\\Leftarrow': '⇐', r'\\iff': '⇔',
    r'\\log': 'log', r'\\ln': 'ln', r'\\exp': 'exp', r'\\lim': 'lim',
    r'\\sum': '∑', r'\\prod': '∏', r'\\int': '∫', r'\\partial': '∂',
    r'\\nabla': '∇', r'\\sqrt': '√'
}


def clean_latex_math(text: str) -> str:
    """
    Transforms LaTeX math syntax and equations into human-readable Unicode text.
    Examples:
        'Bloch sferasi va $\\alpha|0\\rangle + \\beta|1\\rangle$' -> 'Bloch sferasi va α|0⟩ + β|1⟩'
        '$O((\\log N)^3)$' -> 'O((log N)³)'
        'N = p \\times q' -> 'N = p × q'
    """
    if not text:
        return ""

    t = str(text)

    # 1. Clean double markdown bold artifacts: '** **text**' -> '**text**'
    t = re.sub(r'\*\*\s*\*\*', '**', t)

    # 2. Replace Greek letters
    for k, v in GREEK_LETTERS.items():
        t = re.sub(k + r'(?![a-zA-Z])', v, t)

    # 3. Replace standard math symbols
    for k, v in MATH_SYMBOLS.items():
        t = re.sub(k + r'(?![a-zA-Z])', v, t)

    # 4. Fractions: \frac{a}{b} -> (a / b)
    t = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1 / \2)', t)

    # 5. Roots: \sqrt[3]{a} -> ∛(a), \sqrt{a} -> √(a)
    t = re.sub(r'\\sqrt\[3\]\{([^}]+)\}', r'∛(\1)', t)
    t = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', t)

    # 6. Superscripts with braces: ^{2} -> ²
    def repl_sup_brace(m):
        content = m.group(1)
        return ''.join(SUPERSCRIPTS.get(c, c) for c in content)
    t = re.sub(r'\^\{([^}]+)\}', repl_sup_brace, t)

    # 7. Superscripts single char: ^2 -> ²
    def repl_sup_char(m):
        c = m.group(1)
        return SUPERSCRIPTS.get(c, '^' + c)
    t = re.sub(r'\^([0-9a-zA-Z\+\-])', repl_sup_char, t)

    # 8. Subscripts with braces: _{0} -> ₀
    def repl_sub_brace(m):
        content = m.group(1)
        return ''.join(SUBSCRIPTS.get(c, c) for c in content)
    t = re.sub(r'_\{([^}]+)\}', repl_sub_brace, t)

    # 9. Subscripts single char: _0 -> ₀
    def repl_sub_char(m):
        c = m.group(1)
        return SUBSCRIPTS.get(c, '_' + c)
    t = re.sub(r'_([0-9a-zA-Z\+\-])', repl_sub_char, t)

    # 10. Strip math delimiter signs: '$$x$$' -> 'x', '$x$' -> 'x'
    t = t.replace('$$', '').replace('$', '')

    # 11. Clean redundant font/formatting commands: \text{...}, \mathrm{...}, etc.
    t = re.sub(r'\\text\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathit\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\operatorname\{([^}]+)\}', r'\1', t)

    # 12. Clean any remaining stray backslashes before plain alphabetic words
    t = re.sub(r'\\([a-zA-Z]+)', r'\1', t)

    # 13. Clean extra whitespace
    t = re.sub(r'[ \t]+', ' ', t)

    return t

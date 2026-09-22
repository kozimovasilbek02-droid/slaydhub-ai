import re
import sys
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

def clean_latex_math(text: str) -> str:
    if not text:
        return ""
    
    # 1. Clean double markdown bold artifacts: '** **text**' -> '**text**'
    t = re.sub(r'\*\*\s*\*\*', '**', text)

    superscripts = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
        '+': '⁺', '-': '⁻', 'n': 'ⁿ', 'x': 'ˣ', 'y': 'ʸ', 'i': 'ⁱ'
    }
    subscripts = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
        '+': '₊', '-': '₋', 'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ', 'o': 'ₒ',
        'r': 'ᵣ', 'u': 'ᵤ', 'v': 'ᵥ', 'x': 'ₓ'
    }
    
    greek = {
        r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ', r'\\epsilon': 'ε',
        r'\\zeta': 'ζ', r'\\eta': 'η', r'\\theta': 'θ', r'\\iota': 'ι', r'\\kappa': 'κ',
        r'\\lambda': 'λ', r'\\mu': 'μ', r'\\nu': 'ν', r'\\xi': 'ξ', r'\\pi': 'π',
        r'\\rho': 'ρ', r'\\sigma': 'σ', r'\\tau': 'τ', r'\\upsilon': 'υ', r'\\phi': 'φ',
        r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
        r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
        r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Phi': 'Φ', r'\\Psi': 'Ψ', r'\\Omega': 'Ω'
    }
    
    symbols = {
        r'\\times': '×', r'\\cdot': '·', r'\\approx': '≈', r'\\sim': '~',
        r'\\leq': '≤', r'\\le\b': '≤', r'\\geq': '≥', r'\\ge\b': '≥',
        r'\\neq': '≠', r'\\pm': '±', r'\\mp': '∓',
        r'\\infty': '∞', r'\\propto': '∝', r'\\forall': '∀', r'\\exists': '∃',
        r'\\in\b': '∈', r'\\notin': '∉', r'\\subset': '⊂', r'\\subseteq': '⊆',
        r'\\cup': '∪', r'\\cap': '∩', r'\\emptyset': '∅',
        r'\\rangle': '⟩', r'\\langle': '⟨', r'\\rightarrow': '→', r'\\to\b': '→',
        r'\\leftarrow': '←', r'\\Rightarrow': '⇒', r'\\Leftarrow': '⇐', r'\\iff': '⇔',
        r'\\log': 'log', r'\\ln': 'ln', r'\\exp': 'exp', r'\\sum': '∑', r'\\prod': '∏',
        r'\\sqrt': '√'
    }

    # Replace Greek letters
    for k, v in greek.items():
        t = re.sub(k + r'(?![a-zA-Z])', v, t)
        
    # Replace symbols
    for k, v in symbols.items():
        t = re.sub(k + r'(?![a-zA-Z])', v, t)

    # Fractions: \frac{a}{b} -> (a / b)
    t = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1 / \2)', t)
    # Roots: \sqrt{a} -> √(a)
    t = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', t)
    
    # Powers with braces: ^{2} -> ²
    def repl_sup_brace(m):
        content = m.group(1)
        return ''.join(superscripts.get(c, c) for c in content)
    t = re.sub(r'\^\{([^}]+)\}', repl_sup_brace, t)
    
    # Powers single char: ^2 -> ²
    def repl_sup_char(m):
        c = m.group(1)
        return superscripts.get(c, '^' + c)
    t = re.sub(r'\^([0-9a-zA-Z\+\-])', repl_sup_char, t)

    # Subscripts with braces: _{0} -> ₀
    def repl_sub_brace(m):
        content = m.group(1)
        return ''.join(subscripts.get(c, c) for c in content)
    t = re.sub(r'_\{([^}]+)\}', repl_sub_brace, t)
    
    # Subscripts single char: _0 -> ₀
    def repl_sub_char(m):
        c = m.group(1)
        return subscripts.get(c, '_' + c)
    t = re.sub(r'_([0-9a-zA-Z\+\-])', repl_sub_char, t)

    # Strip enclosing $ signs: '$x$' -> 'x', '$$x$$' -> 'x'
    t = t.replace('$$', '').replace('$', '')
    # Clean redundant backslashes before words: '\text{...}' -> '...'
    t = re.sub(r'\\text\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\operatorname\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\([a-zA-Z]+)', r'\1', t)
    return t

sample1 = r"Bloch sferasi va $\alpha|0\rangle + \beta|1\rangle$ superpozitsiya holati."
sample2 = r"Shor Algoritmining $O((\log N)^3)$ Murakkabligi"
sample3 = r"N = p \times q va T_q = k \cdot (\log N)^3"
sample4 = r"** **Shor va Grover algoritmlarining** asimmetrik..."

print("Sample 1 ->", clean_latex_math(sample1))
print("Sample 2 ->", clean_latex_math(sample2))
print("Sample 3 ->", clean_latex_math(sample3))
print("Sample 4 ->", clean_latex_math(sample4))

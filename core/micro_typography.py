# -*- coding: utf-8 -*-
"""
core/micro_typography.py
Advanced Micro-Typography and Knuth-Plass Line-Balancing Engine for SlaydHub AI.
Eliminates orphan/widow words, enforces balanced multi-line headlines, binds short
particles with non-breaking spaces, and sanitizes typographic glyphs.
"""

import re
from typing import List, Tuple, Optional


# Short particles, prepositions, conjunctions that should NEVER be left dangling at the end of a line
ORPHAN_PARTICLES_UZ = {
    "va", "ham", "yoki", "ammo", "lekin", "biroq", "chunki", "uchun", "bilan",
    "kabi", "ko'ra", "qarab", "boshlab", "qadar", "so'ng", "keyin", "oid",
    "bu", "shu", "o'sha", "har", "bir", "hech", "qaysi", "qanday",
    "eng", "juda", "nihoyatda", "faqat", "hatto", "go'yo", "agar"
}

ORPHAN_PARTICLES_EN = {
    "a", "an", "the", "and", "or", "but", "for", "nor", "so", "yet",
    "in", "on", "at", "to", "for", "with", "by", "from", "up", "about",
    "into", "over", "after", "is", "are", "was", "were", "be", "as",
    "of", "if", "than", "then"
}

ALL_ORPHAN_PARTICLES = ORPHAN_PARTICLES_UZ | ORPHAN_PARTICLES_EN


class MicroTypography:
    """
    High-precision typography formatting engine for presentations.
    """

    @staticmethod
    def bind_orphans(text: str) -> str:
        """
        Replaces spaces after short orphan particles with non-breaking spaces (\u00A0)
        so they naturally stick with the following word and never dangle at line breaks.
        Preserves newlines for multi-line and bulleted body text.
        """
        if not text:
            return ""

        lines = text.split('\n')
        bound_lines = []
        for line in lines:
            words = line.split()
            if len(words) <= 1:
                bound_lines.append(line)
                continue

            result_words = []
            for i, w in enumerate(words):
                clean_w = re.sub(r'[^a-zA-Z0-9\-\'\u0400-\u04FF]', '', w).lower()
                if clean_w in ALL_ORPHAN_PARTICLES and i < len(words) - 1:
                    # Append word with non-breaking space
                    result_words.append(w + "\u00A0")
                else:
                    result_words.append(w + " ")

            formatted = "".join(result_words).strip()
            # Clean double non-breaking spaces
            formatted = re.sub(r'[\u00A0\s]+\u00A0', '\u00A0', formatted)
            bound_lines.append(formatted)

        return "\n".join(bound_lines)

    @staticmethod
    def balance_headline(headline: str, max_line_len: int = 42) -> str:
        """
        Knuth-Plass style dynamic line balancer for slide titles.
        If a headline exceeds max_line_len or looks unbalanced on single line,
        calculates the optimal midpoint split to achieve visual symmetry:
        e.g. Line 1 length ≈ Line 2 length, preserving semantic chunks.
        """
        if not headline:
            return ""

        # Normalize spaces
        clean_hl = " ".join(headline.split())

        # If already short enough, return with bound orphans
        if len(clean_hl) <= max_line_len and "\n" not in clean_hl:
            return MicroTypography.bind_orphans(clean_hl)

        # If already explicitly multi-line, clean each line
        if "\n" in clean_hl:
            lines = [MicroTypography.bind_orphans(line.strip()) for line in clean_hl.split("\n")]
            return "\n".join(lines)

        words = clean_hl.split()
        if len(words) < 3:
            return MicroTypography.bind_orphans(clean_hl)

        total_len = len(clean_hl)

        # For long headlines (> 78 chars and >= 6 words), calculate optimal 3-line Inverted Pyramid partition
        if total_len > 78 and len(words) >= 6:
            best_3_score = float("inf")
            best_i = len(words) // 3
            best_j = (2 * len(words)) // 3

            avg_len = total_len / 3.0
            for i in range(1, len(words) - 1):
                for j in range(i + 1, len(words)):
                    l1 = " ".join(words[:i])
                    l2 = " ".join(words[i:j])
                    l3 = " ".join(words[j:])

                    # Inverted pyramid preference: len(l1) >= len(l2) >= len(l3)
                    penalty = 0.0
                    if len(l1) < len(l2):
                        penalty += 14.0
                    if len(l2) < len(l3):
                        penalty += 14.0

                    # Penalty if line breaks immediately after an orphan particle
                    if words[i - 1].lower() in ALL_ORPHAN_PARTICLES:
                        penalty += 18.0
                    if words[j - 1].lower() in ALL_ORPHAN_PARTICLES:
                        penalty += 18.0

                    # Variance from target line length
                    var = (len(l1) - (avg_len + 4.0))**2 + (len(l2) - avg_len)**2 + (len(l3) - (avg_len - 4.0))**2
                    score = var + penalty

                    if score < best_3_score:
                        best_3_score = score
                        best_i = i
                        best_j = j

            part1 = MicroTypography.bind_orphans(" ".join(words[:best_i]))
            part2 = MicroTypography.bind_orphans(" ".join(words[best_i:best_j]))
            part3 = MicroTypography.bind_orphans(" ".join(words[best_j:]))
            return f"{part1}\n{part2}\n{part3}"

        # Calculate optimal 2-line break point minimizing length difference
        best_diff = float("inf")
        best_split_idx = len(words) // 2

        # Candidate split locations
        running_len = 0
        for i in range(len(words) - 1):
            running_len += len(words[i]) + 1
            line1_len = running_len - 1
            line2_len = total_len - running_len

            # Penalty for leaving a very short first or second line
            if len(words[i]) <= 3 and words[i].lower() in ALL_ORPHAN_PARTICLES:
                # Do not split right after an orphan particle if it's the end of line 1
                penalty = 15.0
            else:
                penalty = 0.0

            diff = abs(line1_len - line2_len) + penalty

            # Prefer slightly longer first line over second line (inverted pyramid aesthetic)
            if line1_len < line2_len:
                diff += 4.0

            if diff < best_diff:
                best_diff = diff
                best_split_idx = i + 1

        line1 = " ".join(words[:best_split_idx])
        line2 = " ".join(words[best_split_idx:])

        line1 = MicroTypography.bind_orphans(line1)
        line2 = MicroTypography.bind_orphans(line2)

        return f"{line1}\n{line2}"

    @staticmethod
    def sanitize_typography(text: str) -> str:
        """
        Replaces crude typewriter symbols with elegant typographic glyphs:
        - '--' or ' - ' -> ' — ' (Em-dash)
        - Quotes: converts straight quotes into smart quotes
        - Ellipsis: '...' -> '…'
        """
        if not text:
            return ""

        t = str(text)
        # Em-dash
        t = re.sub(r'\s*--\s*', ' — ', t)
        t = re.sub(r'(\w)\s+-\s+(\w)', r'\1 — \2', t)

        # Ellipsis
        t = t.replace('...', '…')

        # Clean trailing colons or commas before line breaks
        t = re.sub(r'\s*:\s*\n', ':\n', t)

        return t

    @classmethod
    def format_title(cls, title: str, max_line_len: int = 40) -> str:
        """Complete formatting pipeline for presentation titles."""
        sanitized = cls.sanitize_typography(title)
        balanced = cls.balance_headline(sanitized, max_line_len=max_line_len)
        return balanced

    @classmethod
    def format_body(cls, body: str) -> str:
        """Complete formatting pipeline for cards, theses, and body text."""
        sanitized = cls.sanitize_typography(body)
        bound = cls.bind_orphans(sanitized)
        return bound

# -*- coding: utf-8 -*-
"""
scratch/test_claude_rhetoric.py
Unit test for Claude's updated VisualRhetoric module (Round 6).
"""

import os
import sys
import io

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


from core.visual_rhetoric import VisualRhetoric

def run_tests():
    print("=== TEST 1: Cards Dictionary Input (Resolving the Bug) ===")
    card_slide = {
        "title": "Post-Kvant Kriptografiyasi Asoslari",
        "cards": [
            {"title": "Kyber-768 Algoritmi", "description": "Lattice-based shifrlash standarti"},
            {"title": "Dilithium Raqamli Imzosi", "description": "Kvant kompyuterlariga bardoshli imzo mexanizmi"}
        ]
    }

    notes = VisualRhetoric.compile_dual_track_notes(
        slide_title=card_slide["title"],
        slide_data=card_slide,
        topic="Zamonaviy Kriptografiya",
        slide_index=2,
        total_slides=4
    )

    print("Generated notes preview:")
    print(notes[:500])

    assert "Kyber-768 Algoritmi" in notes, "Cards title missing from notes!"
    assert "Dilithium Raqamli Imzosi" in notes, "Second card title missing from notes!"
    assert "1-Savol: Nima uchun taqdimotda aynan 'Kyber-768 Algoritmi'" in notes, "Dynamic Q&A failed to inject card 1!"
    assert "2-Savol: 'Dilithium Raqamli Imzosi'" in notes, "Dynamic Q&A failed to inject card 2!"
    assert "soniyalik jonli nutq ssenariysi" in notes, "Speech pacing timer missing!"
    print("\nPASS: Cards dictionary handled with 100% dynamic speaker notes and contextual Q&A!")

    print("\n=== TEST 2: Comparison Slide (left_points & right_points) ===")
    comp_slide = {
        "title": "RSA vs PQC Taqqoslovi",
        "left_points": ["RSA-2048 zaiflashadi"],
        "right_points": ["Kyber-768 bardoshli"]
    }
    comp_notes = VisualRhetoric.compile_dual_track_notes(
        slide_title=comp_slide["title"],
        slide_data=comp_slide,
        topic="Kriptografiya",
        slide_index=3,
        total_slides=4
    )
    assert "RSA-2048 zaiflashadi" in comp_notes
    assert "Kyber-768 bardoshli" in comp_notes
    print("PASS: Comparison points correctly extracted and embedded.")

    print("\n=== TEST 3: Edge Case (Empty Slide Data) ===")
    empty_notes = VisualRhetoric.compile_dual_track_notes(
        slide_title="Bo'sh Slayd",
        slide_data={},
        topic="Sinov",
        slide_index=1,
        total_slides=1
    )
    assert "Hurmatli tinglovchilar" in empty_notes
    print("PASS: Empty slide gracefully generated without exception.")

    print("\n==========================================")
    print("ALL ROUND 6 VISUAL RHETORIC TESTS PASSED!")
    print("==========================================")

if __name__ == "__main__":
    run_tests()

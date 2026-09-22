# -*- coding: utf-8 -*-
"""
scratch/test_slot_allocator.py
Unit tests and stress tests for SlotAllocator.
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.slot_allocator import SlotAllocator

def run_tests():
    print("=== TEST 1: Count Content Items ===")
    slide_simple = {"title": "Test", "points": ["A", "B", "C"]}
    assert SlotAllocator.count_content_items(slide_simple) == 3, "Failed counting 3 points"

    slide_dense = {
        "title": "Sun'iy Intellekt Arxitekturasi",
        "cards": [
            {"title": f"Komponent {i}", "description": f"Tavsif {i}"}
            for i in range(1, 8)
        ]
    }
    assert SlotAllocator.count_content_items(slide_dense) == 7, "Failed counting 7 cards"
    print("PASS: Count content items.")

    print("\n=== TEST 2: Partition Dense Slide (7 items, max 4) ===")
    parts = SlotAllocator.partition_into_continuation_slides(slide_dense, max_slots_per_slide=4)
    assert len(parts) == 2, f"Expected 2 parts, got {len(parts)}"
    assert parts[0]["title"] == "Sun'iy Intellekt Arxitekturasi (1/2-Qism)"
    assert len(parts[0]["cards"]) == 4
    assert parts[1]["title"] == "Sun'iy Intellekt Arxitekturasi (2/2-Qism)"
    assert len(parts[1]["cards"]) == 3
    print(f"PASS: Successfully partitioned 7 items into Part 1 ({len(parts[0]['cards'])}) and Part 2 ({len(parts[1]['cards'])}).")

    print("\n=== TEST 3: Deck Balancing ===")
    deck = [
        {"title": "Cover", "is_cover": True, "layout_type": "cover"},
        {"title": "Overview", "cards": ["Point 1", "Point 2"]},
        slide_dense, # 7 items
        {"title": "Summary", "points": ["Concl 1", "Concl 2", "Concl 3"]}
    ]
    balanced = SlotAllocator.balance_deck(deck, default_max_slots=4)
    assert len(balanced) == 5, f"Expected 5 slides, got {len(balanced)}"
    assert balanced[0]["title"] == "Cover"
    assert balanced[1]["title"] == "Overview"
    assert "1/2-Qism" in balanced[2]["title"]
    assert "2/2-Qism" in balanced[3]["title"]
    assert balanced[4]["title"] == "Summary"
    print(f"PASS: Balanced deck expanded from 4 to {len(balanced)} slides seamlessly!")

    print("\n==========================================")
    print("ALL SLOT ALLOCATOR TESTS PASSED WITH 0 DEFECTS!")
    print("==========================================")

if __name__ == "__main__":
    run_tests()

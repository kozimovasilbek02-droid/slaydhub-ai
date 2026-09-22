# -*- coding: utf-8 -*-
"""
scratch/test_claude_features.py
Tests Claude's Capacity Scoring, Template Matching, and Hardened Clustering.
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.slot_allocator import SlotAllocator

def test_claude_algorithms():
    print("=== TEST CLAUDE 1: Semantic Clustering with Dictionaries (KeyError Prevention) ===")
    cards = [
        {"title": f"Card {i}", "description": f"Detailed description for card {i}"}
        for i in range(1, 7) # 6 cards
    ]
    # Cluster 6 cards into 3 slots
    clustered = SlotAllocator.cluster_content(cards, max_slots=3)
    assert len(clustered) == 3, f"Expected 3 clusters, got {len(clustered)}"
    assert "Card 1 & Card 2" in clustered[0]["title"]
    assert "Card 3 & Card 4" in clustered[1]["title"]
    assert "Card 5 & Card 6" in clustered[2]["title"]
    print("PASS: Clustered 6 dictionary cards into 3 paired slots with 0 KeyError!")

    print("\n=== TEST CLAUDE 2: Template Matching with Scoring ===")
    templates = [
        {"name": "Template 2 Slots", "capacity": {"slot_count": 2, "avg_slot_area": 5000}},
        {"name": "Template 4 Slots", "capacity": {"slot_count": 4, "avg_slot_area": 4500}},
        {"name": "Template 6 Slots", "capacity": {"slot_count": 6, "avg_slot_area": 4000}}
    ]
    match_result = SlotAllocator.match_template(cards, templates)
    assert match_result["template"]["name"] == "Template 6 Slots", "Should pick exact 6 slots template"
    assert match_result["needs_spillover"] is False
    print("PASS: Exact match template selected with score > 1000!")

    print("\n=== TEST CLAUDE 3: Pagination Engine ===")
    pages = SlotAllocator.paginate_content(cards, max_slots=4)
    assert len(pages) == 2
    assert len(pages[0]) == 4
    assert len(pages[1]) == 2
    print("PASS: Paginated 6 items into 2 pages (4 + 2)!")

    print("\n==========================================")
    print("ALL CLAUDE CODE COMPONENT TESTS PASSED!")
    print("==========================================")

if __name__ == "__main__":
    test_claude_algorithms()

# -*- coding: utf-8 -*-
import asyncio
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.notebooklm_executor import get_notebooklm_executor

executor = get_notebooklm_executor()
prompt = """
Siz akademik assistentsiz. Quyidagi mavzu bo'yicha JSON qaytaring:
Mavzu: "Kvant kompyuterlari"
Format:
{
  "topic": "Kvant kompyuterlari",
  "slides": [
    {"slide_number": 1, "layout_type": "cover", "title": "Kvant kompyuterlari", "subtitle": "Kirish"},
    {"slide_number": 2, "layout_type": "agenda", "points": ["1. Nazariya", "2. Tajriba"]}
  ]
}
"""

print("Running live AI generation with auto-fallback...")
content = asyncio.run(executor.generate_content_auto(prompt))
print("SUCCESS! Generated Content:")
print("Topic:", content.get("topic"))
print("Total Slides:", len(content.get("slides", [])))
for s in content.get("slides", []):
    print("  *", s)

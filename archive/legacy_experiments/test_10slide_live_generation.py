# -*- coding: utf-8 -*-
import os
import sys
import asyncio
import json

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.academic_matcher import get_academic_matcher
from core.notebooklm_prompt_gen import generate_notebooklm_prompt
from core.notebooklm_executor import get_notebooklm_executor

matcher = get_academic_matcher()
bp = matcher.build_blueprint("Kvant kompyuterlari va asimmetrik kriptografiya", slide_count=10, language="uz")

prompt = generate_notebooklm_prompt(bp)
print(f"Blueprint slides: {len(bp['slides'])}, Prompt length: {len(prompt)}")

executor = get_notebooklm_executor()
print("\nExecuting live 10-slide AI generation...")
result = asyncio.run(executor.generate_content_auto(prompt))

print("\n🎉 SUCCESS! Full 10-slide Content Generated and Parsed:")
print("Topic:", result.get("topic"))
slides = result.get("slides", [])
print(f"Total Slides: {len(slides)}")

for s in slides:
    s_num = s.get("slide_number")
    l_type = s.get("layout_type")
    title = s.get("title", s.get("concept_name", ""))
    print(f"  Slide {s_num} [{l_type}]: {title[:60]}")

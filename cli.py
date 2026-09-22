# -*- coding: utf-8 -*-
"""
cli.py
SlaydHub AI — Professional Academic Presentation Studio CLI
Allows generating high-end PPTX presentations directly from the command line.

Usage examples:
  1. Generate from topic using AI:
     python cli.py --topic "Kvant kompyuterlari va asimmetrik kriptografiya" --slides 10 --engine harmonized

  2. Generate from NotebookLM Markdown file:
     python cli.py --markdown notes.md --topic "Kvant kompyuterlari" --engine harmonized -o output/deck.pptx

  3. Get NotebookLM research prompt:
     python cli.py --prompt-only --topic "Kiberxavfsizlik" --slides 10
"""

import os
import sys
import argparse
from pathlib import Path

# Configure Windows console to UTF-8 to prevent charmap UnicodeEncodeError
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from core.config import config
from core.academic_matcher import get_academic_matcher
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
from core.notebooklm_prompt_gen import NotebookLMPromptGenerator
from core.template_pptx_builder import get_academic_pptx_builder


def main():
    parser = argparse.ArgumentParser(
        description="SlaydHub AI — Professional Native PPTX Presentation Studio CLI"
    )
    
    parser.add_argument(
        "--topic", "-t",
        type=str,
        default="Kvant kompyuterlari va asimmetrik kriptografiya",
        help="Presentation topic / research title"
    )
    parser.add_argument(
        "--slides", "-s",
        type=int,
        default=10,
        help="Number of slides to generate (default: 10)"
    )
    parser.add_argument(
        "--markdown", "-m",
        type=str,
        default=None,
        help="Path to Markdown file exported from NotebookLM or Gemini"
    )
    parser.add_argument(
        "--engine", "-e",
        choices=["harmonized", "multi_master", "modular_stacker"],
        default="harmonized",
        help="Assembly engine: harmonized (recommended), multi_master, or modular_stacker"
    )
    parser.add_argument(
        "--lang",
        choices=["uz", "ru", "en"],
        default="uz",
        help="Presentation language (default: uz)"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default=None,
        help="Output PPTX path (default: output/academic_studio/{topic}_Customized.pptx)"
    )
    parser.add_argument(
        "--prompt-only",
        action="store_true",
        help="Print the compact NotebookLM prompt and exit"
    )
    parser.add_argument(
        "--use-gemini",
        action="store_true",
        help="Generate 10-slide content automatically using Gemini AI API"
    )

    args = parser.parse_args()

    # 1. Blueprint matching
    print(f"\n🎓 SlaydHub AI: Initializing presentation for topic: '{args.topic}'...")
    matcher = get_academic_matcher()
    bp = matcher.build_blueprint(
        topic=args.topic,
        slide_count=args.slides,
        language=args.lang
    )

    # 2. Prompt-only mode
    if args.prompt_only:
        prompt = NotebookLMPromptGenerator.generate_compact_markdown_prompt(
            topic=args.topic,
            domain=bp.get("category_name", "Akademik / Ilmiy"),
            lang=args.lang
        )
        print("\n" + "=" * 60)
        print("📌 NOTEBOOKLM RESEARCH PROMPT:")
        print("=" * 60)
        print(prompt)
        print("=" * 60 + "\n")
        return

    # 3. Obtain Content (from file, gemini, or default blueprint)
    raw_markdown = ""
    if args.markdown:
        md_p = Path(args.markdown)
        if not md_p.exists():
            print(f"❌ Error: Markdown file not found at: {md_p}")
            sys.exit(1)
        print(f"📖 Reading NotebookLM content from: {md_p}...")
        with open(md_p, "r", encoding="utf-8") as f:
            raw_markdown = f.read()
    elif args.use_gemini:
        print("🤖 Generating 10-slide academic content via Gemini AI...")
        try:
            raw_markdown = NotebookLMPromptGenerator.generate_with_gemini_direct(
                topic=args.topic,
                domain=bp.get("category_name", "Akademik / Ilmiy"),
                lang=args.lang
            )
        except Exception as e:
            print(f"❌ Error calling Gemini API: {e}")
            print("Falling back to structured template blueprint...")

    # 4. Parse content or build from blueprint
    if raw_markdown.strip():
        parsed = NotebookLMMarkdownParser.parse(raw_markdown, topic=args.topic)
    else:
        # Generate clean structured content directly from blueprint hints
        slides_list = []
        for s in bp.get("slides", []):
            slides_list.append({
                "slide_number": s["slide_number"],
                "layout_type": s.get("layout_type", "cards_grid"),
                "title": s.get("title_hint", f"{s['slide_number']}-Slayd"),
                "subtitle": "",
                "purpose": f"{s.get('layout_name', '')} doirasida ilmiy tahlil",
                "theses": [
                    f"{args.topic} bo'yicha asosiy tahliliy aspekt.",
                    "Ilmiy-amaliy ahamiyati va dolzarb tadqiqot natijalari.",
                    "Xalqaro standartlar va metodologik tavsiyalar."
                ]
            })
        parsed = {
            "topic": args.topic,
            "total_slides": len(slides_list),
            "slides": slides_list
        }

    # 5. Assemble PPTX deck
    print(f"🎨 Assembling presentation using engine: '{args.engine}'...")
    builder = get_academic_pptx_builder()
    out_file = builder.create_presentation(parsed, bp, engine_mode=args.engine)

    if args.output:
        custom_out = Path(args.output)
        custom_out.parent.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copyfile(out_file, custom_out)
        out_file = str(custom_out)

    print("\n" + "=" * 60)
    print(f"🎉 SUCCESS! Presentation generated at:")
    print(f"📁 {out_file}")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()

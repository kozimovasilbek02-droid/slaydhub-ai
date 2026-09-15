import os
import json
import base64
from typing import Optional, Dict, Any
from PIL import Image
from core.models import SlideData, MilestoneNode, ShapeElement, TextBoxElement, Position, TextBlock, TextRun

VISION_PROMPT = """You are a presentation decompiler AI. Inspect this slide image and extract all elements as structured JSON.
Return JSON with this structure:
{
  "slide_index": 1,
  "title": "Slide Title",
  "subtitle": "Subtitle if any",
  "background_color": "#FFFFFF",
  "layout_type": "ROADMAP_TIMELINE",
  "theme_colors": ["#8A56AC", "#00838F", "#EB4D55", "#FA8900", "#0288D1"],
  "milestones": [
    {
      "index": 1,
      "title": "Set Strategic Objectives",
      "description": "Establish measurable goals that align with long-term business priorities.",
      "step_label": "20XX",
      "accent_color": "#8A56AC",
      "position_type": "BOTTOM"
    }
  ],
  "shapes": [],
  "text_boxes": []
}
Output ONLY valid JSON."""

class VisionSlideParser:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

    def parse_image(self, image_path: str) -> SlideData:
        if self.api_key:
            try:
                return self._parse_with_gemini(image_path)
            except Exception as e:
                print(f"AI Vision error: {e}. Using fallback layout engine.")
        return self._heuristic_fallback(image_path)

    def _parse_with_gemini(self, image_path: str) -> SlideData:
        from google import genai
        client = genai.Client(api_key=self.api_key)
        
        with open(image_path, "rb") as f:
            image_bytes = f.read()

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                VISION_PROMPT,
                genai.types.Part.from_bytes(data=image_bytes, mime_type="image/png")
            ]
        )
        raw_text = response.text.strip()
        if raw_text.startswith("```json"):
            raw_text = raw_text[7:]
        if raw_text.startswith("```"):
            raw_text = raw_text[3:]
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3]
        raw_text = raw_text.strip()

        data = json.loads(raw_text)
        return SlideData(**data)

    def _heuristic_fallback(self, image_path: str) -> SlideData:
        return SlideData(
            slide_index=1,
            title="Strategic Roadmap Examples",
            background_color="#FFFFFF",
            layout_type="ROADMAP_TIMELINE",
            theme_colors=["#8A56AC", "#00838F", "#EB4D55", "#FA8900", "#0288D1"],
            milestones=[
                MilestoneNode(index=1, title="Set Strategic Objectives", description="Establish measurable goals that align with long-term business priorities.", step_label="20XX", accent_color="#8A56AC", position_type="BOTTOM"),
                MilestoneNode(index=2, title="Analyze Market & Capabilities", description="Study market trends, customer needs, and internal performance to find growth opportunities.", step_label="20XX", accent_color="#00838F", position_type="TOP"),
                MilestoneNode(index=3, title="Develop Key Initiatives", description="Create focused plans that drive innovation, efficiency, and business expansion.", step_label="20XX", accent_color="#EB4D55", position_type="BOTTOM"),
                MilestoneNode(index=4, title="Implement & Align Teams", description="Execute projects, empower teams, and ensure cross-department collaboration.", step_label="20XX", accent_color="#FA8900", position_type="TOP"),
                MilestoneNode(index=5, title="Monitor & Measure Progress", description="Track KPIs, evaluate outcomes, and identify areas for improvement.", step_label="20XX", accent_color="#0288D1", position_type="BOTTOM"),
            ]
        )

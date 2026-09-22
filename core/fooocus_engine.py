# -*- coding: utf-8 -*-
"""
core/fooocus_engine.py
Fooocus Local AI Image Generator & Prompt Optimizer for SlaydHub AI.

Integrates with local Fooocus instance (default http://127.0.0.1:7865) via REST/Gradio API.
Generates ultra-high quality, presentation-ready visual assets:
- 16:9 Presentation Backgrounds (Abstract, Gradient, Tech, Minimal)
- 3D & Flat Icons (512x512)
- Conceptual Scientific & Business Illustrations
Provides automatic fallback if Fooocus server is not currently running.
"""

import os
import json
import logging
import urllib.request
import urllib.error
from pathlib import Path
from typing import Optional, Dict, Any, List

from core.config import config

logger = logging.getLogger("FooocusEngine")


class FooocusImageGenerator:
    """
    Client for interacting with local Fooocus AI image generation engine.
    """

    def __init__(self, base_url: str = "http://127.0.0.1:7865"):
        self.base_url = os.getenv("FOOOCUS_API_URL", base_url).rstrip("/")
        self.output_dir = config.OUTPUT_ROOT / "fooocus_assets"
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def is_available(self) -> bool:
        """Checks if local Fooocus web UI or API endpoint is active."""
        try:
            req = urllib.request.Request(
                f"{self.base_url}/",
                headers={"User-Agent": "SlaydHub-AI-Client"}
            )
            with urllib.request.urlopen(req, timeout=2) as response:
                return response.status in (200, 302, 404)
        except Exception:
            return False

    @classmethod
    def build_slide_prompt(
        cls,
        topic: str,
        slide_title: str,
        asset_type: str = "illustration"
    ) -> str:
        """
        Builds optimized Fooocus prompt specifically crafted for modern presentations.

        asset_type: 'background' | 'illustration' | 'icon' | 'diagram'
        """
        clean_topic = topic.strip()
        clean_title = slide_title.strip()

        if asset_type == "background":
            return (
                f"Minimalist professional presentation background for {clean_topic}, "
                f"clean abstract gradients, soft blue and deep navy tones (#0A3B75, #2563EB), "
                f"high-end corporate aesthetic, subtle geometric textures, no text, "
                f"no watermark, 16:9 aspect ratio, 8k resolution"
            )
        elif asset_type == "icon":
            return (
                f"Modern 3D isometric glossy icon representing {clean_title} in {clean_topic}, "
                f"minimalist, clean smooth studio lighting, solid white background, "
                f"vibrant blue and purple accent colors (#2563EB, #7C3AED), high detail, 8k"
            )
        elif asset_type == "diagram":
            return (
                f"Scientific visual schematic concept of {clean_title} in {clean_topic}, "
                f"modern infographic aesthetic, clean vector lines, high contrast on white background, "
                f"professional academic clarity, crisp render"
            )
        else:
            return (
                f"Modern digital flat illustration of {clean_title}, {clean_topic}, "
                f"professional tech style, harmonious corporate palette, clean shapes, "
                f"isolated subject, transparent or white background, high quality vector art"
            )

    def generate(
        self,
        prompt: str,
        negative_prompt: str = "text, watermark, logo, blurry, ugly, distorted, low quality, dark background",
        aspect_ratio: str = "16:9",
        steps: int = 30,
        style: str = "Fooocus V2,Fooocus Masterpiece,Fooocus Semi Realistic"
    ) -> Optional[Path]:
        """
        Submits prompt to Fooocus API and saves generated image locally.
        Returns Path to saved image or None if failed/offline.
        """
        if not self.is_available():
            logger.info("Fooocus server is offline at %s. Skipping local generation.", self.base_url)
            return None

        try:
            payload = {
                "prompt": prompt,
                "negative_prompt": negative_prompt,
                "style_selections": style.split(","),
                "performance_selection": "Speed",
                "aspect_ratios_selection": "1152*648" if aspect_ratio == "16:9" else "896*1152",
                "image_number": 1,
                "image_seed": -1,
                "sharpness": 2,
                "guidance_scale": 4.0
            }
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                f"{self.base_url}/v1/generation/text-to-image",
                data=data,
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=45) as resp:
                result = json.loads(resp.read().decode("utf-8"))
                if isinstance(result, list) and len(result) > 0 and "url" in result[0]:
                    img_url = result[0]["url"]
                    local_file = self.output_dir / f"fooocus_{os.urandom(4).hex()}.png"
                    urllib.request.urlretrieve(img_url, local_file)
                    return local_file
        except Exception as e:
            logger.warning("Failed to generate image via Fooocus: %s", e)
        return None

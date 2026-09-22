# -*- coding: utf-8 -*-
"""
core/media_harmonizer.py
Advanced Image Replacement & Aspect-Ratio Harmonization Engine for SlaydHub AI.
- Prevents image distortion via OpenXML Smart Center-Crop (<a:srcRect>).
- Identifies and replaces images sitting deep within AutoShapes (<p:spPr//a:blipFill>)
  as well as native <p:pic> picture shapes.
- Finds safe, collision-free visual zones for placing new images algebraically.
- Protects slide background backdrops from accidental overwrite.
"""

import os
import io
import logging
from typing import List, Tuple, Any, Optional
from PIL import Image
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes.base import BaseShape
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

logger = logging.getLogger("MediaHarmonizer")


class MediaHarmonizer:

    @classmethod
    def calculate_smart_crop(cls, shape_w_pt: float, shape_h_pt: float, img_w: float, img_h: float) -> Tuple[int, int, int, int]:
        """
        Calculates the OpenXML <a:srcRect> margins to achieve a perfect Center-Crop
        without distorting the image aspect ratio.
        Returns: (l, t, r, b) in 1/100,000ths of a percentage.
        """
        if img_h <= 0 or shape_h_pt <= 0 or img_w <= 0 or shape_w_pt <= 0:
            return 0, 0, 0, 0

        shape_aspect = shape_w_pt / shape_h_pt
        img_aspect = img_w / img_h

        l, t, r, b = 0, 0, 0, 0

        # 100,000 = 100% in OpenXML DrawingML
        if img_aspect > shape_aspect:
            # Image is wider than shape -> crop left and right margins equally
            crop_lr_ratio = (1.0 - (shape_aspect / img_aspect)) / 2.0
            crop_val = max(0, min(49000, int(crop_lr_ratio * 100000)))
            l = crop_val
            r = crop_val
        elif img_aspect < shape_aspect:
            # Image is taller than shape -> crop top and bottom margins equally
            crop_tb_ratio = (1.0 - (img_aspect / shape_aspect)) / 2.0
            crop_val = max(0, min(49000, int(crop_tb_ratio * 100000)))
            t = crop_val
            b = crop_val

        return l, t, r, b

    @classmethod
    def apply_srcRect(cls, blipFill_elem: Any, l: int, t: int, r: int, b: int) -> None:
        """
        Injects or updates <a:srcRect> inside a <p:blipFill> or <a:blipFill> tag.
        Ensures strict schema compliance (must sit right after <a:blip>).
        """
        srcRect_nodes = blipFill_elem.xpath('./*[local-name()="srcRect"]')
        if srcRect_nodes:
            srcRect = srcRect_nodes[0]
        else:
            srcRect = OxmlElement('a:srcRect')
            # Insert immediately after <a:blip> if present
            blip_nodes = blipFill_elem.xpath('./*[local-name()="blip"]')
            if blip_nodes:
                blip_elem = blip_nodes[0]
                idx = list(blipFill_elem).index(blip_elem)
                blipFill_elem.insert(idx + 1, srcRect)
            else:
                blipFill_elem.insert(0, srcRect)

        if l > 0: srcRect.set('l', str(l))
        if t > 0: srcRect.set('t', str(t))
        if r > 0: srcRect.set('r', str(r))
        if b > 0: srcRect.set('b', str(b))

    @classmethod
    def find_image_shapes(cls, slide, max_area_ratio: float = 0.45) -> List[BaseShape]:
        """
        Locates all true content images:
        - MSO_SHAPE_TYPE.PICTURE shapes (<p:pic>)
        - AutoShapes mapping images as fills (<p:spPr//a:blipFill>)
        Excludes full-slide background backdrops (> max_area_ratio) to protect templates.
        """
        image_shapes = []
        try:
            prs_part = getattr(slide, "part", None)
            sw_pt = prs_part.package.presentation.slide_width.pt if prs_part else 720.0
            sh_pt = prs_part.package.presentation.slide_height.pt if prs_part else 405.0
            slide_area = max(1.0, sw_pt * sh_pt)
        except Exception:
            slide_area = 720.0 * 405.0

        for shape in slide.shapes:
            try:
                w_pt = shape.width.pt if hasattr(shape.width, "pt") else float(shape.width) / 12700.0
                h_pt = shape.height.pt if hasattr(shape.height, "pt") else float(shape.height) / 12700.0
                # Exclude full-slide background backdrops
                if (w_pt * h_pt) / slide_area > max_area_ratio:
                    continue
            except Exception:
                pass

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image_shapes.append(shape)
                continue

            # Deep OpenXML check for AutoShapes with image fills
            try:
                blipFills = shape._element.xpath('.//*[local-name()="blipFill"]')
                if blipFills:
                    image_shapes.append(shape)
            except Exception:
                pass

        return image_shapes

    @classmethod
    def replace_shape_image(cls, shape: BaseShape, new_image_blob: bytes, img_w_px: float, img_h_px: float) -> bool:
        """
        Overwrites the image payload blob securely and applies Smart Center-Crop.
        Works across both <p:pic> pictures and AutoShapes with <a:blipFill>.
        """
        blip_nodes = shape._element.xpath('.//*[local-name()="blip"]')
        if not blip_nodes:
            return False
        blip = blip_nodes[0]

        rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if not rId:
            return False

        part = shape.part
        if rId in part.rels:
            # Overwrite the payload image directly avoiding PPTX file corruption
            target_part = part.rels[rId].target_part
            target_part._blob = new_image_blob

            # Apply Smart Cropping
            try:
                w_pt = shape.width.pt if hasattr(shape.width, "pt") else float(shape.width) / 12700.0
                h_pt = shape.height.pt if hasattr(shape.height, "pt") else float(shape.height) / 12700.0
                l, t, r, b = cls.calculate_smart_crop(w_pt, h_pt, img_w_px, img_h_px)

                # Find blipFill container (either <p:blipFill> or <a:blipFill>)
                blipFill_nodes = shape._element.xpath('.//*[local-name()="blipFill"]')
                if blipFill_nodes and (l > 0 or t > 0 or r > 0 or b > 0):
                    cls.apply_srcRect(blipFill_nodes[0], l, t, r, b)
            except Exception as e:
                logger.debug(f"Smart crop application: {e}")

            return True

        return False

    @classmethod
    def replace_image_with_file(cls, shape: BaseShape, image_path: str) -> bool:
        """
        Convenience method that reads an image file, extracts natural dimensions via PIL,
        and replaces the shape's image with smart center-cropping.
        """
        if not image_path or not os.path.exists(image_path):
            return False

        try:
            with open(image_path, "rb") as f:
                blob = f.read()

            with Image.open(io.BytesIO(blob)) as img:
                w_px, h_px = img.size

            return cls.replace_shape_image(shape, blob, float(w_px), float(h_px))
        except Exception as e:
            logger.debug(f"Failed to replace image with file: {e}")
            return False

    @classmethod
    def find_safe_image_placement(
        cls,
        slide_width_pt: float,
        slide_height_pt: float,
        occupied_boxes: List[Tuple[float, float, float, float]],
        req_w_pt: float,
        req_h_pt: float
    ) -> Tuple[float, float]:
        """
        Calculates a zero-collision coordinate (left, top) for a new image.
        Scans logical grid sectors to find a safe zone. Returns default if crowded.
        """
        # Define candidate spots (Right Half, Left Half, Bottom Center, Top Right)
        candidates = [
            (slide_width_pt - req_w_pt - 40.0, (slide_height_pt - req_h_pt) / 2.0), # Right Center
            (40.0, (slide_height_pt - req_h_pt) / 2.0),                             # Left Center
            ((slide_width_pt - req_w_pt) / 2.0, slide_height_pt - req_h_pt - 40.0), # Bottom Center
            (slide_width_pt - req_w_pt - 40.0, 40.0)                                # Top Right
        ]

        for (cX, cY) in candidates:
            cR = cX + req_w_pt
            cB = cY + req_h_pt

            collision = False
            for (bx1, by1, bx2, by2) in occupied_boxes:
                pad = 15.0
                if not (cR + pad <= bx1 or cX - pad >= bx2 or cB + pad <= by1 or cY - pad >= by2):
                    collision = True
                    break

            if not collision:
                return cX, cY

        # Fallback to right side mathematically aligned
        return max(0, slide_width_pt - req_w_pt - 20), max(0, (slide_height_pt - req_h_pt) / 2)

import os
from PIL import Image, ImageOps
from pptx.util import Inches

class IconEngine:
    """Intelligent icon extraction and caching engine."""
    
    def __init__(self, cache_dir='output/icon_cache'):
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)
        
    def crop_icon_from_image(self, source_image_path: str, bbox_rel: tuple, icon_name: str) -> str:
        """
        Crops an icon from source image using relative bounding box (x1, y1, x2, y2).
        bbox_rel values are between 0.0 and 1.0.
        Returns the absolute path to the saved cropped PNG.
        """
        out_path = os.path.abspath(os.path.join(self.cache_dir, f"{icon_name}.png"))
        if os.path.exists(out_path):
            return out_path
            
        img = Image.open(source_image_path)
        w, h = img.size
        x1 = int(bbox_rel[0] * w)
        y1 = int(bbox_rel[1] * h)
        x2 = int(bbox_rel[2] * w)
        y2 = int(bbox_rel[3] * h)
        
        crop = img.crop((x1, y1, x2, y2))
        crop.save(out_path)
        return out_path

    def place_icon(self, slide, icon_path: str, x: float, y: float, size: float = 0.40):
        """Places a cropped icon onto a slide at (x, y) with specified size (square)."""
        if not os.path.exists(icon_path):
            return None
        pic = slide.shapes.add_picture(icon_path, Inches(x), Inches(y), width=Inches(size), height=Inches(size))
        return pic

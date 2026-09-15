import os
import numpy as np
from PIL import Image, ImageFilter
from scipy.ndimage import binary_fill_holes

class ObjectCutout:
    """
    Automated foreground object isolation and background cutout tool for PowerPoint slides.
    Separates product/food/icon graphics from flat slide backgrounds and shapes,
    producing clean transparent PNG assets for multi-layer Z-index PowerPoint layouts.
    """
    
    @staticmethod
    def auto_detect_bg_colors(crop_im: Image.Image) -> list:
        """Samples the 4 corners and outer borders of the image to find background colors."""
        w, h = crop_im.size
        arr = np.array(crop_im.convert('RGB'))
        corners = [
            arr[0, 0], arr[0, w-1], 
            arr[h-1, 0], arr[h-1, w-1],
            arr[0, w//2], arr[h-1, w//2]
        ]
        unique_colors = []
        for c in corners:
            if not any(np.linalg.norm(c - uc) < 15 for uc in unique_colors):
                unique_colors.append(c.tolist())
        return unique_colors

    @classmethod
    def isolate(cls, img_or_path, crop_box=None, bg_colors=None, tolerance=22, feather=1, fill_holes=True) -> Image.Image:
        """
        Isolates a foreground object from its background and returns a transparent RGBA PIL Image.
        """
        if isinstance(img_or_path, str):
            im = Image.open(img_or_path).convert('RGBA')
        else:
            im = img_or_path.convert('RGBA')
            
        w, h = im.size
        if crop_box:
            crop_im = im.crop((int(crop_box[0]*w), int(crop_box[1]*h), int(crop_box[2]*w), int(crop_box[3]*h)))
        else:
            crop_im = im
            
        if bg_colors is None:
            bg_colors = cls.auto_detect_bg_colors(crop_im)
            
        arr = np.array(crop_im, dtype=np.float32)
        rgb = arr[:, :, :3]
        
        is_bg = np.zeros((crop_im.height, crop_im.width), dtype=bool)
        for c in bg_colors:
            c_arr = np.array(c, dtype=np.float32)
            dist = np.linalg.norm(rgb - c_arr, axis=2)
            is_bg = is_bg | (dist < tolerance)
            
        fg_mask = ~is_bg
        if fill_holes:
            fg_mask = binary_fill_holes(fg_mask)
            
        alpha_channel = (fg_mask.astype(np.uint8)) * 255
        alpha_im = Image.fromarray(alpha_channel, mode='L')
        
        if feather > 0:
            alpha_im = alpha_im.filter(ImageFilter.GaussianBlur(radius=feather))
            
        crop_im.putalpha(alpha_im)
        return crop_im

    @classmethod
    def isolate_clean_icon(cls, img_or_path, crop_box=None, target_color=(45, 55, 72), bg_threshold=190) -> Image.Image:
        """
        Extracts a clean, artifact-free monochrome icon without outer circular borders or noise.
        """
        if isinstance(img_or_path, str):
            im = Image.open(img_or_path).convert('RGBA')
        else:
            im = img_or_path.convert('RGBA')
            
        w, h = im.size
        if crop_box:
            crop_im = im.crop((int(crop_box[0]*w), int(crop_box[1]*h), int(crop_box[2]*w), int(crop_box[3]*h)))
        else:
            crop_im = im
            
        arr = np.array(crop_im)
        # Background is white/light gray
        is_bg = (arr[:, :, 0] > bg_threshold) & (arr[:, :, 1] > bg_threshold) & (arr[:, :, 2] > bg_threshold)
        arr[is_bg, 3] = 0
        
        # Color foreground pixels uniformly
        is_fg = arr[:, :, 3] > 40
        arr[is_fg, 0] = target_color[0]
        arr[is_fg, 1] = target_color[1]
        arr[is_fg, 2] = target_color[2]
        
        return Image.fromarray(arr)

    @classmethod
    def make_monochrome_variant(cls, img_or_path, out_path: str, color_rgb=(255, 255, 255)) -> str:
        """Creates a recolored monochrome variant (e.g. white for dark cards)."""
        if isinstance(img_or_path, str):
            im = Image.open(img_or_path).convert('RGBA')
        else:
            im = img_or_path.convert('RGBA')
            
        arr = np.array(im)
        is_fg = arr[:, :, 3] > 40
        arr[is_fg, 0] = color_rgb[0]
        arr[is_fg, 1] = color_rgb[1]
        arr[is_fg, 2] = color_rgb[2]
        
        out_im = Image.fromarray(arr)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        out_im.save(out_path)
        return os.path.abspath(out_path)

    @classmethod
    def extract_and_save(cls, img_or_path, out_path: str, crop_box=None, bg_colors=None, tolerance=22) -> str:
        """Extracts cutout and saves to out_path, returning the absolute path."""
        res = cls.isolate(img_or_path, crop_box=crop_box, bg_colors=bg_colors, tolerance=tolerance)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        res.save(out_path)
        return os.path.abspath(out_path)

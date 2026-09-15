import os
import numpy as np
from PIL import Image, ImageChops, ImageEnhance

class VisualDiff:
    @staticmethod
    def calculate_ssim(img1_path: str, img2_path: str, target_size=(1280, 720)) -> float:
        """
        Calculates mathematical Structural Similarity Index (SSIM) between two slide images.
        Returns a float between 0.0 and 1.0 (1.0 = identical).
        """
        im1 = Image.open(img1_path).convert('L').resize(target_size, Image.Resampling.BILINEAR)
        im2 = Image.open(img2_path).convert('L').resize(target_size, Image.Resampling.BILINEAR)
        
        arr1 = np.array(im1, dtype=np.float64)
        arr2 = np.array(im2, dtype=np.float64)
        
        C1 = (0.01 * 255) ** 2
        C2 = (0.03 * 255) ** 2
        
        mu1 = np.mean(arr1)
        mu2 = np.mean(arr2)
        
        sigma1_sq = np.var(arr1)
        sigma2_sq = np.var(arr2)
        sigma12 = np.mean((arr1 - mu1) * (arr2 - mu2))
        
        ssim = ((2 * mu1 * mu2 + C1) * (2 * sigma12 + C2)) / ((mu1**2 + mu2**2 + C1) * (sigma1_sq + sigma2_sq + C2))
        return float(np.clip(ssim, 0.0, 1.0))

    @staticmethod
    def generate_diff_heatmap(img_ref_path: str, img_gen_path: str, output_path: str, target_size=(1280, 720)) -> str:
        """
        Generates a visual difference heatmap highlighting shifted elements in magenta/red.
        Saves the heatmap image to output_path.
        """
        im1 = Image.open(img_ref_path).convert('RGB').resize(target_size, Image.Resampling.BILINEAR)
        im2 = Image.open(img_gen_path).convert('RGB').resize(target_size, Image.Resampling.BILINEAR)
        
        # Absolute pixel difference
        diff = ImageChops.difference(im1, im2)
        
        # Boost contrast of diff
        enhancer = ImageEnhance.Contrast(diff)
        diff_boosted = enhancer.enhance(3.0)
        
        # Colorize difference: convert to heatmap (overlay red on top of base image)
        arr_base = np.array(im1, dtype=np.float32)
        arr_diff = np.array(diff_boosted.convert('L'), dtype=np.float32) / 255.0
        
        # Highlight changed regions in red/cyan
        overlay = arr_base.copy()
        overlay[..., 0] = np.clip(overlay[..., 0] + arr_diff * 200, 0, 255)  # Boost red
        overlay[..., 1] = np.clip(overlay[..., 1] * (1.0 - arr_diff * 0.7), 0, 255)
        overlay[..., 2] = np.clip(overlay[..., 2] * (1.0 - arr_diff * 0.7), 0, 255)
        
        result_img = Image.fromarray(overlay.astype(np.uint8))
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        result_img.save(output_path)
        return output_path

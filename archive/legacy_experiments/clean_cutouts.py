import os
import numpy as np
from PIL import Image
from scipy.ndimage import binary_fill_holes

os.makedirs('output/f11_cutouts', exist_ok=True)
folder_path = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\15_Famous_Foods_Of_The_Usa'

def isolate_food_object(img_path, crop_box, bg_colors, tolerance=15):
    """
    Isolates food object from solid/flat background colors by generating an alpha mask.
    """
    im = Image.open(img_path).convert('RGBA')
    w, h = im.size
    crop_im = im.crop((int(crop_box[0]*w), int(crop_box[1]*h), int(crop_box[2]*w), int(crop_box[3]*h)))
    
    arr = np.array(crop_im, dtype=np.float32)
    rgb = arr[:, :, :3]
    
    # Background mask: where pixel matches any of the given background colors
    is_bg = np.zeros((crop_im.height, crop_im.width), dtype=bool)
    for c in bg_colors:
        c_arr = np.array(c, dtype=np.float32)
        dist = np.linalg.norm(rgb - c_arr, axis=2)
        is_bg = is_bg | (dist < tolerance)
        
    # Invert to get foreground
    fg_mask = ~is_bg
    # Fill internal holes inside the burger so seeds/cheese with similar shades aren't transparent
    fg_mask = binary_fill_holes(fg_mask)
    
    alpha = (fg_mask.astype(np.uint8)) * 255
    arr[:, :, 3] = alpha
    
    result = Image.fromarray(arr.astype(np.uint8), mode='RGBA')
    return result

# Slide 2: Burger (Left background is #FFC000 [255, 192, 0], right is #FFFFFF [255, 255, 255])
im2_cutout = isolate_food_object(
    os.path.join(folder_path, 'Slide_02.png'),
    crop_box=(0, 0.20, 0.40, 1.0),
    bg_colors=[[255, 192, 0], [255, 255, 255]],
    tolerance=25
)
im2_cutout.save('output/f11_cutouts/s2_burger_clean.png')
print('Burger clean cutout saved!')

# Slide 7: Buffalo wings (Left side had yellow strip [255, 192, 0] or [245, 166, 35])
# Let's crop ONLY the dark platter and avoid any yellow block!
im7 = Image.open(os.path.join(folder_path, 'Slide_07.png'))
w7, h7 = im7.size
# Pure platter crop (between x=24% and x=61%)
c7 = im7.crop((int(0.238*w7), 0, int(0.612*w7), h7))
c7.save('output/f11_cutouts/s7_wings_clean.png')
print('Wings clean platter saved!')

# Slide 8: Hot dog plate
# Crop pure plate on white
im8_cutout = isolate_food_object(
    os.path.join(folder_path, 'Slide_08.png'),
    crop_box=(0.45, 0.15, 0.95, 0.75),
    bg_colors=[[255, 255, 255], [255, 192, 0]],
    tolerance=20
)
im8_cutout.save('output/f11_cutouts/s8_hotdog_clean.png')
print('Hotdog clean cutout saved!')

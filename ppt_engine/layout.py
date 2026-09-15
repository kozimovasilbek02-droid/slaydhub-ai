import math
from typing import List, Tuple

def distribute_horizontal(count: int, card_width: float, slide_width: float = 13.333, margin_x: float = 0.80) -> List[float]:
    """
    Distributes N cards horizontally across slide_width with equal spacing.
    Returns list of X coordinates (in inches).
    """
    if count == 1:
        return [(slide_width - card_width) / 2.0]
        
    usable_w = slide_width - (2 * margin_x)
    total_cards_w = count * card_width
    
    if total_cards_w >= usable_w:
        gap = 0.15
        start_x = (slide_width - (total_cards_w + (count - 1) * gap)) / 2.0
    else:
        gap = (usable_w - total_cards_w) / (count - 1)
        start_x = margin_x
        
    return [start_x + i * (card_width + gap) for i in range(count)]

def distribute_vertical(count: int, card_height: float, start_y: float = 1.60, end_y: float = 6.60) -> List[float]:
    """
    Distributes N cards vertically between start_y and end_y with equal spacing.
    Returns list of Y coordinates (in inches).
    """
    if count == 1:
        return [(start_y + end_y - card_height) / 2.0]
        
    usable_h = end_y - start_y
    total_cards_h = count * card_height
    gap = (usable_h - total_cards_h) / (count - 1)
    
    return [start_y + i * (card_height + gap) for i in range(count)]

def radial_distribution(cx: float, cy: float, radius: float, count: int, start_angle_deg: float = 0.0) -> List[Tuple[float, float, float]]:
    """
    Calculates (x, y, angle_rad) for N items evenly spaced on a circle.
    """
    coords = []
    angle_step = (2 * math.pi) / count
    start_rad = math.radians(start_angle_deg)
    
    for i in range(count):
        angle = start_rad + i * angle_step
        x = cx + radius * math.cos(angle)
        y = cy + radius * math.sin(angle)
        coords.append((x, y, angle))
    return coords

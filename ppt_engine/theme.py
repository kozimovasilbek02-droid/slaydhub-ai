from pptx.dml.color import RGBColor
from pptx.util import Pt, Inches

def hex_to_rgb(hex_code: str) -> RGBColor:
    """Converts hex string like '#E53E3E' or 'E53E3E' to pptx RGBColor."""
    h = hex_code.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

class Colors:
    # Modern Professional Palette
    NAVY = '#0A3B75'
    ORANGE = '#F15A24'
    RED = '#E53E3E'
    GOLD = '#D69E2E'
    GREEN = '#38A169'
    TEAL = '#00A3C4'
    BLUE = '#2B6CB0'
    PURPLE = '#805AD5'
    PINK = '#D53F8C'
    
    # Neutral Tones
    TEXT_MAIN = '#1A202C'       # Darkest slate for headlines
    TEXT_TITLE = '#2D3748'      # Slate 800 for titles
    TEXT_MUTED = '#718096'      # Slate 500 for body/captions
    TEXT_LIGHT = '#A0AEC0'      # Slate 400 for subtle accents
    TEXT_WHITE = '#FFFFFF'
    
    # Card Backgrounds & Borders
    BG_SLIDE = '#FFFFFF'
    BG_CARD = '#FFFFFF'
    BORDER_LIGHT = '#E2E8F0'
    BORDER_SUBTLE = '#EDF2F7'
    
    # Soft Shadows
    SHADOW_AMBIENT = '#E2E8F0'
    SHADOW_MEDIUM = '#CBD5E1'
    SHADOW_DARK = '#94A3B8'
    
    # Standard Step Colors (10 color spectrum)
    PALETTE_10 = [
        '#00A3C4',  # 01: Teal
        '#2B6CB0',  # 02: Blue
        '#805AD5',  # 03: Purple
        '#E53E3E',  # 04: Red
        '#DD6B20',  # 05: Orange
        '#D69E2E',  # 06: Gold
        '#38A169',  # 07: Green
        '#3182CE',  # 08: Bright Blue
        '#9F7AEA',  # 09: Light Purple
        '#00B4D8',  # 10: Cyan
    ]

class Fonts:
    TITLE = 'Montserrat'
    BODY = 'Segoe UI'
    ACCENT = 'Montserrat'

class Sizes:
    # Slide Dimensions (16:9 Standard)
    SLIDE_WIDTH = Inches(13.333)
    SLIDE_HEIGHT = Inches(7.500)
    
    # Header Sizes
    HEADER_X = Inches(0.80)
    HEADER_Y = Inches(0.45)
    HEADER_W = Inches(11.733)
    HEADER_H = Inches(0.85)
    
    # Typography Font Sizes
    SIZE_TITLE_MAIN = Pt(28)
    SIZE_TITLE_CARD = Pt(11.5)
    SIZE_SUBTITLE = Pt(10)
    SIZE_BODY = Pt(8.5)
    SIZE_BADGE = Pt(9.5)
    SIZE_HERO_NUM = Pt(18)

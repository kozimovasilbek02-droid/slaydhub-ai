"""
ppt_engine: Universal Professional PowerPoint Generator Framework
"""

from ppt_engine.theme import Colors, Fonts, Sizes, hex_to_rgb
from ppt_engine.geometry import (
    add_soft_shadow,
    create_seamless_callout,
    create_chevron,
    create_downward_banner,
    create_smooth_path,
    create_circular_arc,
    create_horizontal_connector,
    add_header_accent,
    add_background_wireframe_squares,
    add_dashed_vertical_divider,
    add_progress_indicators
)
from ppt_engine.typography import (
    add_slide_header,
    add_card_text,
    calculate_optimal_font_size,
    add_constrained_textbox
)
from ppt_engine.cards import (
    create_floating_card,
    create_pill_card,
    create_two_tone_card,
    create_circular_badge
)
from ppt_engine.layout import distribute_horizontal, distribute_vertical, radial_distribution
from ppt_engine.icons import IconEngine
from ppt_engine.qa import VisionQA
from ppt_engine.diff import VisualDiff
from ppt_engine.cutout import ObjectCutout
from ppt_engine.blueprints.wave_timeline import render_wave_timeline

__all__ = [
    'Colors',
    'Fonts',
    'Sizes',
    'hex_to_rgb',
    'add_soft_shadow',
    'create_seamless_callout',
    'create_chevron',
    'create_downward_banner',
    'create_smooth_path',
    'create_circular_arc',
    'create_horizontal_connector',
    'add_slide_header',
    'add_card_text',
    'calculate_optimal_font_size',
    'add_constrained_textbox',
    'create_floating_card',
    'create_pill_card',
    'create_two_tone_card',
    'create_circular_badge',
    'distribute_horizontal',
    'distribute_vertical',
    'radial_distribution',
    'IconEngine',
    'VisionQA',
    'VisualDiff',
    'render_wave_timeline'
]

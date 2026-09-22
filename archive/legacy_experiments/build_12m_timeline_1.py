import os
import math
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from ppt_engine.theme import hex_to_rgb, Colors, Fonts, Sizes
from ppt_engine.geometry import add_soft_shadow, create_smooth_path
from ppt_engine.typography import add_slide_header, add_card_text
from ppt_engine.cards import create_floating_card, create_circular_badge
from ppt_engine.qa import VisionQA

# ==================== EXTRACT ICONS ====================
def extract_clean_icons(source_slide_path: str):
    """Extracts transparent anti-aliased white icons from source slide."""
    os.makedirs('output/icons_extracted', exist_ok=True)
    icon_boxes = {
        'briefcase': (116, 341, 162, 387),
        'recycle': (345, 341, 391, 387),
        'target': (568, 341, 614, 387),
        'climber': (793, 341, 839, 387),
    }
    extracted = {}
    im = Image.open(source_slide_path)
    for name, box in icon_boxes.items():
        cropped = im.crop(box).convert('RGBA')
        arr = np.array(cropped, dtype=np.float32)
        r, g, b = arr[:,:,0], arr[:,:,1], arr[:,:,2]
        min_rgb = np.minimum(np.minimum(r, g), b)
        alpha = np.clip((min_rgb - 90) / (220 - 90) * 255.0, 0, 255).astype(np.uint8)
        
        out_arr = np.zeros((arr.shape[0], arr.shape[1], 4), dtype=np.uint8)
        out_arr[:,:,:3] = 255
        out_arr[:,:,3] = alpha
        
        out_img = Image.fromarray(out_arr)
        out_img_hires = out_img.resize((out_img.width * 4, out_img.height * 4), Image.Resampling.LANCZOS)
        out_path = os.path.abspath(f'output/icons_extracted/icon_{name}.png')
        out_img_hires.save(out_path)
        extracted[name] = out_path
    return extracted

import sys
SOURCE_DIR = sys.argv[1] if len(sys.argv) > 1 else (
    r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint_1'
    if os.path.exists(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint_1')
    else r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint_1'
)
ICONS = extract_clean_icons(os.path.join(SOURCE_DIR, 'Slide_01.png'))

# ==================== PRESENTATION SETUP ====================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

FONT_TITLE = 'Georgia'
FONT_BODY = 'Segoe UI'

def add_header(slide, title_text='Online Timeline Template', align=PP_ALIGN.CENTER, x=0.80, y=0.55, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_TITLE
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#2D3748')
    p.alignment = align
    return tb

# ==============================================================================
# SLIDE 1: 4 WHITE CARDS WITH DOWNWARD RIBBON CHEVRONS & NUMBER BADGES
# ==============================================================================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, 'Online Timeline Template')

s1_data = [
    ('20XX', 'Business refers to\nthe activities\ninvolved in\nproducing and selling\ngoods or services for\nprofit', '#F37022', '01', ICONS['briefcase'], 0.80),
    ('20XX', 'A process is a series\nof steps or actions\nthat are taken in\norder to achieve a\nparticular goal or\nobjective.', '#C92605', '02', ICONS['recycle'], 3.85),
    ('20XX', 'A target is a specific\ngoal or objective that\na person, team, or\norganization aims to\nachieve.', '#6460AA', '03', ICONS['target'], 6.90),
    ('20XX', 'Success can be\ndefined as the\nachievement of a\ndesired goal or\noutcome.', '#0CB14B', '04', ICONS['climber'], 9.95),
]

card1_w = 2.58
card1_h = 3.65
card1_y = 1.45

for yr, desc, col, num, icon_path, x in s1_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. Soft Shadow behind White Card
    add_soft_shadow(s1, MSO_SHAPE.RECTANGLE, x, card1_y, card1_w, card1_h, 0.04, 0.04, Colors.SHADOW_AMBIENT)
    
    # 2. White Card
    card = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(card1_y), Inches(card1_w), Inches(card1_h))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.fill.background()
    
    # 3. Top subtle watermark V-notch
    v_w, v_h = 0.50, 0.35
    v_pts = [
        (Inches(x + card1_w / 2 - v_w / 2), Inches(card1_y)),
        (Inches(x + card1_w / 2), Inches(card1_y + v_h)),
        (Inches(x + card1_w / 2 + v_w / 2), Inches(card1_y))
    ]
    fb_v = s1.shapes.build_freeform(v_pts[0][0], v_pts[0][1])
    fb_v.add_line_segments(v_pts[1:], close=True)
    v_shape = fb_v.convert_to_shape()
    v_shape.fill.solid()
    v_shape.fill.fore_color.rgb = hex_to_rgb('#F4F4F4')
    v_shape.line.color.rgb = hex_to_rgb('#EEEEEE')
    v_shape.line.width = Pt(1.0)
    
    # 4. Text inside Card (20XX and Description)
    tb = s1.shapes.add_textbox(Inches(x + 0.15), Inches(card1_y + 0.55), Inches(card1_w - 0.30), Inches(1.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = yr
    p1.font.name = FONT_TITLE
    p1.font.size = Pt(19)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(8)
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = hex_to_rgb('#4A5568')
    p2.alignment = PP_ALIGN.CENTER
    
    # 5. Downward Ribbon Chevron Bookmark at bottom of card
    chv_w = 1.62
    chv_x = x + (card1_w - chv_w) / 2
    chv_y = card1_y + card1_h - 1.10
    chv_h = 1.85
    notch_d = 0.50
    
    pts_chv = [
        (Inches(chv_x), Inches(chv_y)),
        (Inches(chv_x + chv_w / 2), Inches(chv_y + notch_d)),
        (Inches(chv_x + chv_w), Inches(chv_y)),
        (Inches(chv_x + chv_w), Inches(chv_y + chv_h - notch_d)),
        (Inches(chv_x + chv_w / 2), Inches(chv_y + chv_h)),
        (Inches(chv_x), Inches(chv_y + chv_h - notch_d))
    ]
    
    # Chevron Shadow
    pts_shd = [(px + Inches(0.03), py + Inches(0.03)) for px, py in pts_chv]
    fb_cs = s1.shapes.build_freeform(pts_shd[0][0], pts_shd[0][1])
    fb_cs.add_line_segments(pts_shd[1:], close=True)
    cshd = fb_cs.convert_to_shape()
    cshd.fill.solid()
    cshd.fill.fore_color.rgb = hex_to_rgb(Colors.SHADOW_AMBIENT)
    cshd.line.fill.background()
    
    # Chevron Main
    fb_c = s1.shapes.build_freeform(pts_chv[0][0], pts_chv[0][1])
    fb_c.add_line_segments(pts_chv[1:], close=True)
    chv = fb_c.convert_to_shape()
    chv.fill.solid()
    chv.fill.fore_color.rgb = c_rgb
    chv.line.fill.background()
    
    # Icon on Chevron
    if os.path.exists(icon_path):
        icon_size = 0.55
        s1.shapes.add_picture(
            icon_path,
            Inches(chv_x + (chv_w - icon_size) / 2),
            Inches(chv_y + 0.72),
            width=Inches(icon_size),
            height=Inches(icon_size)
        )
        
    # 6. Connecting vertical line below chevron tip
    tip_x = chv_x + chv_w / 2
    tip_y = chv_y + chv_h
    ln = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(tip_x - 0.01), Inches(tip_y), Inches(0.02), Inches(0.40))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#94A3B8')
    ln.line.fill.background()
    
    # 7. Circular Badge at bottom with number
    badge_diam = 0.95
    badge_y = tip_y + 0.35
    badge_x = tip_x - badge_diam / 2
    
    # Outer colored ring
    add_soft_shadow(s1, MSO_SHAPE.OVAL, badge_x - 0.08, badge_y - 0.08, badge_diam + 0.16, badge_diam + 0.16, 0.02, 0.02)
    ring = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_x - 0.08), Inches(badge_y - 0.08), Inches(badge_diam + 0.16), Inches(badge_diam + 0.16))
    ring.fill.solid()
    ring.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    ring.line.color.rgb = c_rgb
    ring.line.width = Pt(2.0)
    
    # Inner colored circle
    badge = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_x), Inches(badge_y), Inches(badge_diam), Inches(badge_diam))
    badge.fill.solid()
    badge.fill.fore_color.rgb = c_rgb
    badge.line.fill.background()
    
    p_num = badge.text_frame.paragraphs[0]
    p_num.text = num
    p_num.font.name = FONT_TITLE
    p_num.font.size = Pt(16)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 2: 4 ROUNDED CARDS WITH TOP TABS & DIAMOND TIMELINE ICONS
# ==============================================================================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, 'Online Timeline Template')

s2_data = [
    ('Business', 'Business refers to\nthe activities\ninvolved in\nproducing and\nselling goods or\nservices for profit.', '#F37022', ICONS['briefcase'], 1.00),
    ('Process', 'A process is a series\nof steps or actions\nthat are taken in\norder to achieve a\nparticular goal or\nobjective.', '#6460AA', ICONS['recycle'], 3.85),
    ('Target', 'A target is a\nspecific goal or\nobjective that a\nperson, team, or\norganization aims to\nachieve.', '#C92605', ICONS['target'], 6.70),
    ('Success', 'Success can be\ndefined as the\nachievement of a\ndesired goal or\noutcome.', '#0CB14B', ICONS['climber'], 9.55),
]

# Horizontal Blue Timeline Arrow Bar
bar_y = 5.25
bar_h = 0.09
# Start dot
dot = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.65), Inches(bar_y - 0.065), Inches(0.22), Inches(0.22))
dot.fill.solid()
dot.fill.fore_color.rgb = hex_to_rgb('#0089D1')
dot.line.fill.background()

# Bar line
bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(bar_y), Inches(11.45), Inches(bar_h))
bar.fill.solid()
bar.fill.fore_color.rgb = hex_to_rgb('#0089D1')
bar.line.fill.background()

# End Arrowhead
arrow_pts = [
    (Inches(12.20), Inches(bar_y - 0.14)),
    (Inches(12.50), Inches(bar_y + bar_h / 2)),
    (Inches(12.20), Inches(bar_y + bar_h + 0.14))
]
fb_ar = s2.shapes.build_freeform(arrow_pts[0][0], arrow_pts[0][1])
fb_ar.add_line_segments(arrow_pts[1:], close=True)
arr_head = fb_ar.convert_to_shape()
arr_head.fill.solid()
arr_head.fill.fore_color.rgb = hex_to_rgb('#0089D1')
arr_head.line.fill.background()

card2_w = 2.45
card2_h = 3.65
card2_y = 1.55

for title, desc, col, icon_path, x in s2_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. Soft Shadow behind Card
    add_soft_shadow(s2, MSO_SHAPE.ROUNDED_RECTANGLE, x, card2_y, card2_w, card2_h, 0.04, 0.04, Colors.SHADOW_AMBIENT)
    
    # 2. Gray Card Body
    body_shape = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(card2_y), Inches(card2_w), Inches(card2_h))
    body_shape.fill.solid()
    body_shape.fill.fore_color.rgb = hex_to_rgb('#B8B8B8')
    body_shape.line.fill.background()
    
    # 3. Top Tab with Pointer
    tab_h = 0.85
    tab_pts = [
        (Inches(x), Inches(card2_y + 0.20)),
        (Inches(x + 0.05), Inches(card2_y + 0.10)),
        (Inches(x + 0.12), Inches(card2_y + 0.03)),
        (Inches(x + 0.22), Inches(card2_y)),
        (Inches(x + card2_w - 0.22), Inches(card2_y)),
        (Inches(x + card2_w - 0.12), Inches(card2_y + 0.03)),
        (Inches(x + card2_w - 0.05), Inches(card2_y + 0.10)),
        (Inches(x + card2_w), Inches(card2_y + 0.20)),
        (Inches(x + card2_w), Inches(card2_y + tab_h - 0.18)),
        (Inches(x + card2_w / 2 + 0.22), Inches(card2_y + tab_h - 0.18)),
        (Inches(x + card2_w / 2), Inches(card2_y + tab_h)),
        (Inches(x + card2_w / 2 - 0.22), Inches(card2_y + tab_h - 0.18)),
        (Inches(x), Inches(card2_y + tab_h - 0.18)),
    ]
    fb_tab = s2.shapes.build_freeform(tab_pts[0][0], tab_pts[0][1])
    fb_tab.add_line_segments(tab_pts[1:], close=True)
    tab = fb_tab.convert_to_shape()
    tab.fill.solid()
    tab.fill.fore_color.rgb = c_rgb
    tab.line.fill.background()
    
    # Tab Text
    tb_tab = s2.shapes.add_textbox(Inches(x), Inches(card2_y + 0.15), Inches(card2_w), Inches(0.50))
    tf_tab = tb_tab.text_frame
    p_t = tf_tab.paragraphs[0]
    p_t.text = title
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_t.alignment = PP_ALIGN.CENTER
    
    # Body Text
    tb_body = s2.shapes.add_textbox(Inches(x + 0.16), Inches(card2_y + 1.15), Inches(card2_w - 0.32), Inches(2.20))
    tf_body = tb_body.text_frame
    tf_body.word_wrap = True
    tf_body.margin_left = tf_body.margin_right = tf_body.margin_top = tf_body.margin_bottom = 0
    p_b = tf_body.paragraphs[0]
    p_b.text = desc
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(9.5)
    p_b.font.color.rgb = hex_to_rgb('#2D3748')
    p_b.alignment = PP_ALIGN.CENTER
    
    # 4. Rotated Diamond Icon Marker on Timeline
    diam_sz = 1.15
    diam_cx = x + card2_w / 2
    diam_cy = bar_y + bar_h / 2
    
    # Diamond Shadow
    add_soft_shadow(s2, MSO_SHAPE.DIAMOND, diam_cx - diam_sz/2, diam_cy - diam_sz/2, diam_sz, diam_sz, 0.03, 0.03)
    
    # Diamond Shape
    diam = s2.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(diam_cx - diam_sz/2), Inches(diam_cy - diam_sz/2), Inches(diam_sz), Inches(diam_sz))
    diam.fill.solid()
    diam.fill.fore_color.rgb = c_rgb
    diam.line.color.rgb = hex_to_rgb('#FFFFFF')
    diam.line.width = Pt(2.0)
    
    # Place Icon
    if os.path.exists(icon_path):
        icon_size = 0.48
        s2.shapes.add_picture(
            icon_path,
            Inches(diam_cx - icon_size/2),
            Inches(diam_cy - icon_size/2),
            width=Inches(icon_size),
            height=Inches(icon_size)
        )
        
    # 5. Year Label below Diamond
    tb_yr = s2.shapes.add_textbox(Inches(diam_cx - 1.0), Inches(diam_cy + diam_sz/2 + 0.22), Inches(2.0), Inches(0.50))
    tf_yr = tb_yr.text_frame
    p_yr = tf_yr.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = FONT_TITLE
    p_yr.font.size = Pt(17)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#2D3748')
    p_yr.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 3: 4 CONCENTRIC TARGET MARKERS WITH ALTERNATING PILLS & TEXT
# ==============================================================================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3, 'Online Timeline Template')

s3_data = [
    ('Business', 'Business refers to the\nactivities involved in\nproducing and selling goods\nor services for profit.', '#C92605', '#FDECE8', True, 0.70),
    ('Process', 'A process is a series of\nsteps or actions that are\ntaken in order to achieve a\nparticular goal or objective.', '#F37022', '#FEF0E9', False, 3.80),
    ('Target', 'A target is a specific goal or\nobjective that a person,\nteam, or organization aims to\nachieve.', '#0089D1', '#E5F3FB', True, 6.90),
    ('Success', 'Success can be defined as\nthe achievement of a\ndesired goal or outcome.', '#0CB14B', '#E7F7ED', False, 10.00),
]

y_center = 3.80

for title, desc, col, col_tint, is_top, x in s3_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. Concentric Target Marker
    target_cx = x + 0.65
    target_cy = y_center
    
    # Outer pale ring
    r1 = 0.52
    add_soft_shadow(s3, MSO_SHAPE.OVAL, target_cx - r1, target_cy - r1, r1*2, r1*2, 0.02, 0.02)
    outer_ring = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(target_cx - r1), Inches(target_cy - r1), Inches(r1*2), Inches(r1*2))
    outer_ring.fill.solid()
    outer_ring.fill.fore_color.rgb = hex_to_rgb(col_tint)
    outer_ring.line.color.rgb = c_rgb
    outer_ring.line.width = Pt(1.5)
    
    # Mid solid circle
    r2 = 0.36
    mid_circle = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(target_cx - r2), Inches(target_cy - r2), Inches(r2*2), Inches(r2*2))
    mid_circle.fill.solid()
    mid_circle.fill.fore_color.rgb = c_rgb
    mid_circle.line.fill.background()
    
    # Inner white circle
    r3 = 0.16
    in_circle = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(target_cx - r3), Inches(target_cy - r3), Inches(r3*2), Inches(r3*2))
    in_circle.fill.solid()
    in_circle.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    in_circle.line.fill.background()
    
    # Center colored dot
    r4 = 0.07
    dot = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(target_cx - r4), Inches(target_cy - r4), Inches(r4*2), Inches(r4*2))
    dot.fill.solid()
    dot.fill.fore_color.rgb = c_rgb
    dot.line.fill.background()
    
    # 2. Pill Card next to Target
    pill_x = target_cx + r1 + 0.30
    pill_y = target_cy - 0.45
    pill_w = 1.60
    pill_h = 0.90
    
    add_soft_shadow(s3, MSO_SHAPE.ROUNDED_RECTANGLE, pill_x, pill_y, pill_w, pill_h, 0.03, 0.03)
    pill = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(pill_x), Inches(pill_y), Inches(pill_w), Inches(pill_h))
    pill.fill.solid()
    pill.fill.fore_color.rgb = c_rgb
    pill.line.fill.background()
    
    p_yr = pill.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = FONT_TITLE
    p_yr.font.size = Pt(17)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_yr.alignment = PP_ALIGN.CENTER
    
    # 3. Connecting Vertical Line & Text Block
    line_x = target_cx
    if is_top:
        # Line from text top down to outer ring
        top_y = 1.80
        ln = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(line_x - 0.01), Inches(top_y), Inches(0.025), Inches(target_cy - r1 - top_y))
        ln.fill.solid()
        ln.fill.fore_color.rgb = c_rgb
        ln.line.fill.background()
        
        # Text Above (left aligned next to line)
        tb = s3.shapes.add_textbox(Inches(line_x + 0.18), Inches(top_y), Inches(2.20), Inches(1.40))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = hex_to_rgb('#2D3748')
        p_t.space_after = Pt(4)
        
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = hex_to_rgb('#555555')
    else:
        # Line from outer ring down to bottom text
        bot_end_y = 6.20
        start_y = target_cy + r1
        ln = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(line_x - 0.01), Inches(start_y), Inches(0.025), Inches(bot_end_y - start_y))
        ln.fill.solid()
        ln.fill.fore_color.rgb = c_rgb
        ln.line.fill.background()
        
        # Text Below (left aligned next to line)
        tb = s3.shapes.add_textbox(Inches(line_x + 0.18), Inches(start_y + 0.35), Inches(2.20), Inches(1.40))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = hex_to_rgb('#2D3748')
        p_t.space_after = Pt(4)
        
        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = hex_to_rgb('#555555')

# ==============================================================================
# SLIDE 4: 3D PERSPECTIVE TIMELINE BEAM WITH WAVING FLAGS & CONCENTRIC BASES
# ==============================================================================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, 'Online Timeline Template')

# 3D Perspective Road Track
# Bottom shadow edge
road_shd = [
    (Inches(0.00), Inches(6.85)),
    (Inches(13.333), Inches(4.25)),
    (Inches(13.333), Inches(4.55)),
    (Inches(0.00), Inches(7.15))
]
fb_rs = s4.shapes.build_freeform(road_shd[0][0], road_shd[0][1])
fb_rs.add_line_segments(road_shd[1:], close=True)
road_s = fb_rs.convert_to_shape()
road_s.fill.solid()
road_s.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
road_s.line.fill.background()

# Top surface
road_pts = [
    (Inches(0.00), Inches(6.75)),
    (Inches(13.333), Inches(4.15)),
    (Inches(13.333), Inches(4.35)),
    (Inches(0.00), Inches(6.95))
]
fb_r = s4.shapes.build_freeform(road_pts[0][0], road_pts[0][1])
fb_r.add_line_segments(road_pts[1:], close=True)
road = fb_r.convert_to_shape()
road.fill.solid()
road.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
road.line.fill.background()

s4_data = [
    ('Business', 'Business refers to the\nactivities involved in\nproducing and selling\ngoods or services for profit.', '#ECA815', 1.50, 6.55, 3.10),
    ('Process', 'A process is a series of\nsteps or actions that are\ntaken in order to achieve a\nparticular goal or objective.', '#DB6119', 4.50, 5.95, 2.85),
    ('Target', 'A target is a specific goal\nor objective that a person,\nteam, or organization aims\nto achieve.', '#AF2200', 7.50, 5.35, 2.60),
    ('Success', 'Success can be defined\nas the achievement of a\ndesired goal or outcome.', '#555193', 10.50, 4.75, 2.35),
]

for title, desc, col, px, py_ground, py_flag in s4_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. Concentric Ellipse Base on Road
    ew1, eh1 = 0.88, 0.34
    add_soft_shadow(s4, MSO_SHAPE.OVAL, px - ew1/2, py_ground - eh1/2, ew1, eh1, 0.02, 0.02)
    
    el_outer = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px - ew1/2), Inches(py_ground - eh1/2), Inches(ew1), Inches(eh1))
    el_outer.fill.solid()
    el_outer.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    el_outer.line.color.rgb = c_rgb
    el_outer.line.width = Pt(2.5)
    
    ew2, eh2 = 0.44, 0.18
    el_inner = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px - ew2/2), Inches(py_ground - eh2/2), Inches(ew2), Inches(eh2))
    el_inner.fill.solid()
    el_inner.fill.fore_color.rgb = c_rgb
    el_inner.line.fill.background()
    
    # 2. Vertical Flagpole
    pole_h = py_ground - py_flag
    pole = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px - 0.02), Inches(py_flag), Inches(0.04), Inches(pole_h))
    pole.fill.solid()
    pole.fill.fore_color.rgb = c_rgb
    pole.line.fill.background()
    
    # 3. Waving Flag Shape (Wave with curved swallowtail or clean wave)
    flag_w = 1.55
    flag_h = 0.78
    
    pts_flag = [
        (Inches(px), Inches(py_flag)),
        (Inches(px + flag_w * 0.40), Inches(py_flag - 0.08)),
        (Inches(px + flag_w * 0.80), Inches(py_flag + 0.04)),
        (Inches(px + flag_w), Inches(py_flag + 0.08)),
        (Inches(px + flag_w), Inches(py_flag + flag_h + 0.08)),
        (Inches(px + flag_w * 0.80), Inches(py_flag + flag_h + 0.04)),
        (Inches(px + flag_w * 0.40), Inches(py_flag + flag_h - 0.08)),
        (Inches(px), Inches(py_flag + flag_h))
    ]
    fb_fl = s4.shapes.build_freeform(pts_flag[0][0], pts_flag[0][1])
    fb_fl.add_line_segments(pts_flag[1:], close=True)
    flag = fb_fl.convert_to_shape()
    flag.fill.solid()
    flag.fill.fore_color.rgb = c_rgb
    flag.line.fill.background()
    
    # Flag Text
    tb_fl = s4.shapes.add_textbox(Inches(px + 0.10), Inches(py_flag + 0.10), Inches(flag_w - 0.20), Inches(flag_h - 0.20))
    tf_fl = tb_fl.text_frame
    p_fl = tf_fl.paragraphs[0]
    p_fl.text = '20XX'
    p_fl.font.name = FONT_TITLE
    p_fl.font.size = Pt(14)
    p_fl.font.bold = True
    p_fl.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_fl.alignment = PP_ALIGN.CENTER
    
    # 4. Text next to pole
    tb_txt = s4.shapes.add_textbox(Inches(px + 0.20), Inches(py_flag + flag_h + 0.35), Inches(2.10), Inches(1.50))
    tf_txt = tb_txt.text_frame
    tf_txt.word_wrap = True
    tf_txt.margin_left = tf_txt.margin_right = tf_txt.margin_top = tf_txt.margin_bottom = 0
    p_t = tf_txt.paragraphs[0]
    p_t.text = title
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(12.5)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#2D3748')
    p_t.space_after = Pt(4)
    
    p_d = tf_txt.add_paragraph()
    p_d.text = desc
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(8.5)
    p_d.font.color.rgb = hex_to_rgb('#555555')

# ==============================================================================
# SLIDE 5: 4 MAP-PIN LOCATION CARDS ON TIMELINE BASELINE
# ==============================================================================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5, 'Online Timeline Template')

# Horizontal Baseline
base_y = 5.25
line5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.50), Inches(base_y), Inches(12.333), Inches(0.03))
line5.fill.solid()
line5.fill.fore_color.rgb = hex_to_rgb('#A2A2A2')
line5.line.fill.background()

s5_data = [
    ('Business', 'Business refers to\nthe activities involved in\nproducing and selling\ngoods or services for\nprofit.', '#FCB712', 1.05),
    ('Process', 'A process is a series of\nsteps or actions that are\ntaken in order to achieve\na particular goal or\nobjective.', '#0CB14B', 4.05),
    ('Target', 'A target is a specific\ngoal or objective that a\nperson, team, or\norganization aims to\nachieve.', '#C92605', 7.05),
    ('Success', 'Success can be defined\nas the achievement of a\ndesired goal or\noutcome.', '#6460AA', 10.05),
]

pin_w = 2.20
pin_body_h = 1.95
pin_tip_y = 5.15
pin_y = 1.90

for title, desc, col, x in s5_data:
    c_rgb = hex_to_rgb(col)
    
    # Smooth curved rounded pin perimeter
    cx = x + pin_w / 2
    r_cor = 0.28
    
    pts_pin = [
        # Top-left corner
        (Inches(x + r_cor), Inches(pin_y)),
        (Inches(x + pin_w - r_cor), Inches(pin_y)),
        (Inches(x + pin_w), Inches(pin_y + r_cor)),
        # Right side down to taper start
        (Inches(x + pin_w), Inches(pin_y + pin_body_h - r_cor)),
        (Inches(x + pin_w - r_cor), Inches(pin_y + pin_body_h)),
        (Inches(cx + 0.15), Inches(pin_y + pin_body_h)),
        # Taper tip
        (Inches(cx), Inches(pin_tip_y)),
        # Taper back to left
        (Inches(cx - 0.15), Inches(pin_y + pin_body_h)),
        (Inches(x + r_cor), Inches(pin_y + pin_body_h)),
        (Inches(x), Inches(pin_y + pin_body_h - r_cor)),
        (Inches(x), Inches(pin_y + r_cor)),
    ]
    
    # Pin Shadow
    pts_shd = [(px + Inches(0.03), py + Inches(0.03)) for px, py in pts_pin]
    fb_ps = s5.shapes.build_freeform(pts_shd[0][0], pts_shd[0][1])
    fb_ps.add_line_segments(pts_shd[1:], close=True)
    pshd = fb_ps.convert_to_shape()
    pshd.fill.solid()
    pshd.fill.fore_color.rgb = hex_to_rgb(Colors.SHADOW_AMBIENT)
    pshd.line.fill.background()
    
    # Pin Main Shape
    fb_p = s5.shapes.build_freeform(pts_pin[0][0], pts_pin[0][1])
    fb_p.add_line_segments(pts_pin[1:], close=True)
    pin = fb_p.convert_to_shape()
    pin.fill.solid()
    pin.fill.fore_color.rgb = c_rgb
    pin.line.fill.background()
    
    # Text inside Pin
    tb = s5.shapes.add_textbox(Inches(x + 0.15), Inches(pin_y + 0.35), Inches(pin_w - 0.30), Inches(1.50))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(6)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(8.5)
    p_d.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_d.alignment = PP_ALIGN.CENTER
    
    # 2. Milestone Node Circle on baseline
    node_diam = 0.38
    node_cx = cx
    node_cy = base_y + 0.015
    
    node = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(node_cx - node_diam/2), Inches(node_cy - node_diam/2), Inches(node_diam), Inches(node_diam))
    node.fill.solid()
    node.fill.fore_color.rgb = c_rgb
    node.line.fill.background()
    
    # 3. Year label below baseline
    tb_yr = s5.shapes.add_textbox(Inches(node_cx - 1.0), Inches(base_y + 0.40), Inches(2.0), Inches(0.50))
    tf_yr = tb_yr.text_frame
    p_yr = tf_yr.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = FONT_TITLE
    p_yr.font.size = Pt(17)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#2D3748')
    p_yr.alignment = PP_ALIGN.CENTER

# ==================== SAVE OUTPUT ====================
final_output_path = os.path.abspath(os.path.join(SOURCE_DIR, 'presentation_95_precision.pptx'))
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print(f"Presentation saved successfully to {final_output_path}")

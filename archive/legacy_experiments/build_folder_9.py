import sys
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, hex_to_rgb, add_slide_header, add_card_text, create_circular_badge, add_soft_shadow
from ppt_engine.geometry import create_chevron, create_downward_banner, create_smooth_path
from ppt_engine.layout import distribute_horizontal

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint_1'
os.makedirs('output/f9_assets', exist_ok=True)

# 1. Crop exact crisp white icons from Slide_01
img1 = Image.open(os.path.join(folder_path, 'Slide_01.png'))
w1, h1 = img1.size

icon_crops_s1 = [
    (0.125, 0.655, 0.165, 0.725),  # Briefcase
    (0.362, 0.655, 0.402, 0.725),  # Process recycle
    (0.598, 0.655, 0.638, 0.725),  # Target
    (0.835, 0.655, 0.875, 0.725),  # Success climbing
]

icon_paths_s1 = []
for idx, box in enumerate(icon_crops_s1):
    c = img1.crop((int(box[0]*w1), int(box[1]*h1), int(box[2]*w1), int(box[3]*h1)))
    p = os.path.abspath(f'output/f9_assets/icon_s1_{idx+1}.png')
    c.save(p)
    icon_paths_s1.append(p)

# Base text content
data_text = [
    ('Business', 'Business refers to the activities involved in producing and selling goods or services for profit.'),
    ('Process', 'A process is a series of steps or actions that are taken in order to achieve a particular goal or objective.'),
    ('Target', 'A target is a specific goal or objective that a person, team, or organization aims to achieve.'),
    ('Success', 'Success can be defined as the achievement of a desired goal or outcome.'),
]

# ==================== SLIDE 1: PERFECT NOTCHED RIBBONS + CROPPED WHITE ICONS ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s1, 'Online Timeline Template', x=0.80, y=0.55, w=11.733)

s1_colors = ['#EB6C24', '#BF2800', '#5855A2', '#149F44']
xs = distribute_horizontal(4, card_width=2.45, slide_width=13.333, margin_x=0.90)

for i in range(4):
    x = xs[i]
    col_hex = s1_colors[i]
    title, desc = data_text[i]
    
    # White Top Card with soft shadow
    add_soft_shadow(s1, MSO_SHAPE.RECTANGLE, x, 1.80, 2.45, 3.20, 0.03, 0.03)
    crd = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.80), Inches(2.45), Inches(3.20))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.fill.background()
    
    # Top Year
    tb_yr = s1.shapes.add_textbox(Inches(x), Inches(2.15), Inches(2.45), Inches(0.40))
    p_yr = tb_yr.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = Fonts.TITLE
    p_yr.font.size = Pt(13)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p_yr.alignment = PP_ALIGN.CENTER
    
    # Body
    tb_b = s1.shapes.add_textbox(Inches(x + 0.20), Inches(2.70), Inches(2.05), Inches(1.80))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.text = desc
    p_b.font.name = Fonts.BODY
    p_b.font.size = Pt(8.8)
    p_b.font.color.rgb = hex_to_rgb(Colors.TEXT_MUTED)
    p_b.alignment = PP_ALIGN.CENTER
    
    # Downward Ribbon with V-cut top and pointed bottom
    rw = 1.75
    rh = 1.65
    rx = x + 0.35
    ry = 4.35
    pts = [
        (Inches(rx), Inches(ry)),
        (Inches(rx + rw/2.0), Inches(ry + 0.30)),
        (Inches(rx + rw), Inches(ry)),
        (Inches(rx + rw), Inches(ry + rh - 0.45)),
        (Inches(rx + rw/2.0), Inches(ry + rh)),
        (Inches(rx), Inches(ry + rh - 0.45))
    ]
    fb = s1.shapes.build_freeform(pts[0][0], pts[0][1])
    fb.add_line_segments(pts[1:], close=True)
    rib = fb.convert_to_shape()
    rib.fill.solid()
    rib.fill.fore_color.rgb = hex_to_rgb(col_hex)
    rib.line.fill.background()
    
    # White Cropped Icon
    s1.shapes.add_picture(icon_paths_s1[i], Inches(rx + (rw - 0.65)/2.0), Inches(ry + 0.55), width=Inches(0.65))
    
    # Vertical connector
    ln = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 1.215), Inches(5.95), Inches(0.02), Inches(0.40))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb(Colors.BORDER_LIGHT)
    ln.line.fill.background()
    
    # Number Badges
    create_circular_badge(s1, x + 0.825, 6.25, 0.80, bg_color='#FFFFFF', border_color=col_hex, text='', has_shadow=True)
    create_circular_badge(s1, x + 0.885, 6.31, 0.68, bg_color=col_hex, text=f"{i+1:02d}", text_color='#FFFFFF', font_size=11, has_shadow=False)

# ==================== SLIDE 2: EXACT COLOR PALETTE & CALLOUT HEADER NOTCHES ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s2, 'Online Timeline Template', x=0.80, y=0.55, w=11.733)

# Exact colors: Orange, Purple, Red, Green
s2_colors = ['#EB6C24', '#5855A2', '#BF2800', '#149F44']

arr = s2.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(0.80), Inches(5.50), Inches(11.733), Inches(0.08))
arr.fill.solid()
arr.fill.fore_color.rgb = hex_to_rgb('#0099FF')
arr.line.fill.background()

for i in range(4):
    x = xs[i]
    col_hex = s2_colors[i]
    title, desc = data_text[i]
    
    # Tall Card
    add_soft_shadow(s2, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.00, 2.45, 3.40, 0.03, 0.03)
    crd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.00), Inches(2.45), Inches(3.40))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    crd.line.fill.background()
    
    # Header with bottom callout pointer notch
    hw = 2.45
    hh = 0.70
    h_pts = [
        (Inches(x), Inches(2.00)),
        (Inches(x + hw), Inches(2.00)),
        (Inches(x + hw), Inches(2.00 + hh)),
        (Inches(x + hw/2.0 + 0.15), Inches(2.00 + hh)),
        (Inches(x + hw/2.0), Inches(2.00 + hh + 0.15)),
        (Inches(x + hw/2.0 - 0.15), Inches(2.00 + hh)),
        (Inches(x), Inches(2.00 + hh))
    ]
    fb_h = s2.shapes.build_freeform(h_pts[0][0], h_pts[0][1])
    fb_h.add_line_segments(h_pts[1:], close=True)
    hdr = fb_h.convert_to_shape()
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = hex_to_rgb(col_hex)
    hdr.line.fill.background()
    
    tb_h = s2.shapes.add_textbox(Inches(x), Inches(2.08), Inches(hw), Inches(0.40))
    p_h = tb_h.text_frame.paragraphs[0]
    p_h.text = title
    p_h.font.name = Fonts.TITLE
    p_h.font.size = Pt(12)
    p_h.font.bold = True
    p_h.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_h.alignment = PP_ALIGN.CENTER
    
    # Body text
    tb_b = s2.shapes.add_textbox(Inches(x + 0.20), Inches(2.95), Inches(2.05), Inches(2.10))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.text = desc
    p_b.font.name = Fonts.BODY
    p_b.font.size = Pt(8.8)
    p_b.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p_b.alignment = PP_ALIGN.CENTER
    
    # Diamond with White Cropped Icon
    dia = s2.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(x + 0.775), Inches(5.05), Inches(0.90), Inches(0.90))
    dia.fill.solid()
    dia.fill.fore_color.rgb = hex_to_rgb(col_hex)
    dia.line.color.rgb = hex_to_rgb('#FFFFFF')
    dia.line.width = Pt(2.0)
    
    s2.shapes.add_picture(icon_paths_s1[i], Inches(x + 0.95), Inches(5.22), width=Inches(0.55))
    
    # Year
    tb_yr = s2.shapes.add_textbox(Inches(x), Inches(6.15), Inches(2.45), Inches(0.40))
    p_yr = tb_yr.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = Fonts.TITLE
    p_yr.font.size = Pt(12)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p_yr.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 3: EXACT COLOR ORDER (RED, ORANGE, BLUE, GREEN) + STADIUM PILLS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s3, 'Online Timeline Template', x=0.80, y=0.55, w=11.733)

s3_colors = ['#C02200', '#EB6C24', '#0084D1', '#149F44']
s3_xs = [1.00, 3.95, 6.90, 9.85]

for i in range(4):
    x = s3_xs[i]
    col_hex = s3_colors[i]
    title, desc = data_text[i]
    is_top = (i % 2 == 0)
    
    # Target concentric badge
    create_circular_badge(s3, x, 3.80, 0.95, bg_color='#FFFFFF', border_color=col_hex, text='', has_shadow=True)
    create_circular_badge(s3, x + 0.15, 3.95, 0.65, bg_color=col_hex, text='◎', text_color='#FFFFFF', font_size=16, has_shadow=False)
    
    # Full stadium pill for 20XX
    pill = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 1.20), Inches(3.85), Inches(1.45), Inches(0.85))
    pill.fill.solid()
    pill.fill.fore_color.rgb = hex_to_rgb(col_hex)
    pill.line.fill.background()
    pill.adjustments[0] = 0.5  # Max roundness = Stadium Pill
    p_p = pill.text_frame.paragraphs[0]
    p_p.text = '20XX'
    p_p.font.name = Fonts.TITLE
    p_p.font.size = Pt(13)
    p_p.font.bold = True
    p_p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_p.alignment = PP_ALIGN.CENTER
    
    # Vertical connecting line
    y_ln_start = 2.40 if is_top else 4.80
    y_ln_h = 1.35
    ln = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.465), Inches(y_ln_start), Inches(0.02), Inches(y_ln_h))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb(col_hex)
    ln.line.fill.background()
    
    # Text Block
    y_tb = 1.70 if is_top else 5.20
    tb = s3.shapes.add_textbox(Inches(x + 0.60), Inches(y_tb), Inches(2.05), Inches(1.50))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = Fonts.TITLE
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = Fonts.BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb(Colors.TEXT_MUTED)
    p2.alignment = PP_ALIGN.LEFT

# ==================== SLIDE 4: EXACT COLORS (GOLD, ORANGE, BRICK, SLATE PURPLE) + PERSPECTIVE ELLIPSES ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s4, 'Online Timeline Template', x=0.80, y=0.55, w=11.733)

s4_colors = ['#F5A623', '#E55B2B', '#A81C07', '#5C5B94']
road_pts = [(Inches(0.20), Inches(6.80)), (Inches(13.10), Inches(4.30))]
fb_r = s4.shapes.build_freeform(road_pts[0][0], road_pts[0][1])
fb_r.add_line_segments(road_pts[1:], close=False)
rd = fb_r.convert_to_shape()
rd.fill.background()
rd.line.color.rgb = hex_to_rgb('#CBD5E1')
rd.line.width = Pt(6.0)

flag_xs = [1.20, 3.80, 6.40, 9.00]
flag_pole_ys = [6.50, 5.85, 5.20, 4.55]

for i in range(4):
    x = flag_xs[i]
    py = flag_pole_ys[i]
    col_hex = s4_colors[i]
    title, desc = data_text[i]
    
    # Perspective Squashed Base Ellipse (h = w * 0.42)
    el_w = 0.85
    el_h = 0.36
    b1 = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.05), Inches(py - el_h/2.0), Inches(el_w), Inches(el_h))
    b1.fill.solid()
    b1.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    b1.line.color.rgb = hex_to_rgb(col_hex)
    b1.line.width = Pt(2.0)
    
    b2 = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.22), Inches(py - (el_h*0.55)/2.0), Inches(el_w*0.60), Inches(el_h*0.55))
    b2.fill.solid()
    b2.fill.fore_color.rgb = hex_to_rgb(col_hex)
    b2.line.fill.background()
    
    # Flag Pole
    pole = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.46), Inches(py - 3.20), Inches(0.03), Inches(3.20))
    pole.fill.solid()
    pole.fill.fore_color.rgb = hex_to_rgb(col_hex)
    pole.line.fill.background()
    
    # Flag Wave
    flg = s4.shapes.add_shape(MSO_SHAPE.WAVE, Inches(x + 0.48), Inches(py - 3.20), Inches(1.35), Inches(0.75))
    flg.fill.solid()
    flg.fill.fore_color.rgb = hex_to_rgb(col_hex)
    flg.line.fill.background()
    p_f = flg.text_frame.paragraphs[0]
    p_f.text = '20XX'
    p_f.font.name = Fonts.TITLE
    p_f.font.size = Pt(11)
    p_f.font.bold = True
    p_f.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_f.alignment = PP_ALIGN.CENTER
    
    # Card text
    tb = s4.shapes.add_textbox(Inches(x + 0.60), Inches(py - 2.20), Inches(2.05), Inches(1.80))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = Fonts.TITLE
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = Fonts.BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb(Colors.TEXT_MUTED)

# ==================== SLIDE 5: EXACT COLORS (YELLOW, GREEN, RED, PURPLE) + ROUNDED BALLOON PINS ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s5, 'Online Timeline Template', x=0.80, y=0.55, w=11.733)

s5_colors = ['#F5A623', '#00A84D', '#BF1E0F', '#5E5A9B']

line5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(5.60), Inches(11.733), Inches(0.02))
line5.fill.solid()
line5.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
line5.line.fill.background()

for i in range(4):
    x = xs[i]
    col_hex = s5_colors[i]
    title, desc = data_text[i]
    
    # Rounded Card Top + Tapered Base
    pw = 2.45
    ph = 2.00
    pin_pts = [
        (Inches(x + 0.20), Inches(2.30)),
        (Inches(x + pw - 0.20), Inches(2.30)),
        (Inches(x + pw), Inches(2.50)),
        (Inches(x + pw), Inches(4.30)),
        (Inches(x + pw/2.0), Inches(5.45)),
        (Inches(x), Inches(4.30)),
        (Inches(x), Inches(2.50))
    ]
    fb_p = s5.shapes.build_freeform(pin_pts[0][0], pin_pts[0][1])
    fb_p.add_line_segments(pin_pts[1:], close=True)
    pin = fb_p.convert_to_shape()
    pin.fill.solid()
    pin.fill.fore_color.rgb = hex_to_rgb(col_hex)
    pin.line.fill.background()
    
    tb = s5.shapes.add_textbox(Inches(x + 0.20), Inches(2.60), Inches(2.05), Inches(1.80))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = Fonts.TITLE
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = Fonts.BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#F7FAFC')
    p2.alignment = PP_ALIGN.CENTER
    
    # Milestone Bead
    create_circular_badge(s5, x + 1.075, 5.45, 0.30, bg_color=col_hex, has_shadow=False)
    
    # 20XX
    tb_yr = s5.shapes.add_textbox(Inches(x), Inches(5.95), Inches(2.45), Inches(0.40))
    p_yr = tb_yr.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = Fonts.TITLE
    p_yr.font.size = Pt(12)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p_yr.alignment = PP_ALIGN.CENTER

out_f9_fixed = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint_1\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Fixed presentation for Folder 9 saved to:', out_f9_fixed)

# Re-export previews
exported = VisionQA.export_slides_to_png(out_f9_fixed, output_dir='output/folder9_perfect_previews')
for exp in exported:
    print('Exported fixed preview:', exp)

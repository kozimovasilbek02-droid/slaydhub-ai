import sys
import os
import win32com.client
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def hex_to_rgb(hex_code):
    h = hex_code.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

def add_header(slide, title_text, subtitle_text=None, align=PP_ALIGN.CENTER, x=0.8, y=0.50, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1.10))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = 'Montserrat'
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#222222')
    p.alignment = align
    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(13)
        p2.font.color.rgb = hex_to_rgb('#666666')
        p2.alignment = align

# ==================== SLIDE 1: 5 DOWNWARD-POINTED CARDS ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, '100 Days Action Plan', 'In 5 Phases')

s1_data = [
    ('First 20\nDays', ['Understand the company culture', 'Outline current year commitment plans', 'Create yearly agenda with stakeholders'], '#8CC63F', 0.65),
    ('20-40\nDays', ['Meetings with internal teams', 'Build sales funnel'], '#29B6D8', 3.10),
    ('40-60\nDays', ['Executing plan', 'Formal meeting with clients', 'Dialogue with suitable prospects'], '#FFAA00', 5.55),
    ('60-80\nDays', ['Full engagement with client & team on existing opportunities'], '#EA4335', 8.00),
    ('80-100\nDays', ['Working towards mile stones'], '#AF26FA', 10.45),
]

cw = 2.25
for title, bullets, col, x in s1_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. Base White Card with Border
    card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.00), Inches(cw), Inches(4.70))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.color.rgb = hex_to_rgb('#E2E8F0')
    card.line.width = Pt(1.5)
    
    # 2. Downward Pointed Header Banner (Pentagon pointing down)
    pts_banner = [
        (Inches(x), Inches(2.00)),
        (Inches(x + cw), Inches(2.00)),
        (Inches(x + cw), Inches(3.30)),
        (Inches(x + cw/2), Inches(3.70)),
        (Inches(x), Inches(3.30))
    ]
    fb = s1.shapes.build_freeform(pts_banner[0][0], pts_banner[0][1])
    fb.add_line_segments(pts_banner[1:], close=True)
    ban = fb.convert_to_shape()
    ban.fill.solid()
    ban.fill.fore_color.rgb = c_rgb
    ban.line.fill.background()
    
    # Banner Text
    tb_b = s1.shapes.add_textbox(Inches(x), Inches(2.15), Inches(cw), Inches(1.10))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    lines = title.split('\n')
    p1 = tf_b.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(14)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf_b.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER
        
    # Card Bullet Text
    tb_t = s1.shapes.add_textbox(Inches(x + 0.15), Inches(3.90), Inches(cw - 0.30), Inches(2.60))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    for i, b in enumerate(bullets):
        p = tf_t.paragraphs[0] if i == 0 else tf_t.add_paragraph()
        p.text = '✓  ' + b
        p.font.name = 'Segoe UI'
        p.font.size = Pt(10)
        p.font.color.rgb = hex_to_rgb('#444444')
        p.space_after = Pt(8)

# ==================== SLIDE 2: 4 OVERLAPPING CIRCULAR CARDS ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, '100 Days Action Plan', 'In 4 Phases')

s2_data = [
    ('First 25 Days', ['Observe for understanding', 'Start opportunity assessment'], '#8CC63F', '01', False, 0.70),
    ('25-50 Days', ['Build month sales calendar', 'Sales reporting'], '#29B6D8', '02', True, 3.75),
    ('50-75 Days', ['Forecasting', 'Report assessment findings', 'Introduce changes in roles, responsibilities', 'Constructive feedback'], '#FFAA00', '03', False, 6.80),
    ('75-100 Days', ['Introduce changes in roles, responsibilities'], '#EA4335', '04', True, 9.85),
]

diam = 2.90
for title, bullets, col, num, is_top_tab, x in s2_data:
    c_rgb = hex_to_rgb(col)
    y_circle = 2.30
    
    # Colored Outer Ring / Crescent Arc
    ring = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - 0.12), Inches(y_circle - 0.12), Inches(diam + 0.24), Inches(diam + 0.24))
    ring.fill.solid()
    ring.fill.fore_color.rgb = c_rgb
    ring.line.fill.background()
    
    # White Inner Circle
    circle = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y_circle), Inches(diam), Inches(diam))
    circle.fill.solid()
    circle.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    circle.line.color.rgb = hex_to_rgb('#E2E8F0')
    circle.line.width = Pt(1.5)
    
    # Semi-circular Number Tab (Top or Bottom)
    y_tab = y_circle - 0.40 if is_top_tab else y_circle + diam - 0.35
    tab = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + diam/2 - 0.45), Inches(y_tab), Inches(0.90), Inches(0.75))
    tab.fill.solid()
    tab.fill.fore_color.rgb = c_rgb
    tab.line.fill.background()
    p_num = tab.text_frame.paragraphs[0]
    p_num.text = num
    p_num.font.name = 'Montserrat'
    p_num.font.size = Pt(12)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER
    
    # Text inside Circle
    tb_c = s2.shapes.add_textbox(Inches(x + 0.25), Inches(y_circle + 0.40), Inches(diam - 0.50), Inches(diam - 0.80))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = 0
    p_t = tf_c.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(12.5)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#222222')
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(6)
    
    for b in bullets:
        p_b = tf_c.add_paragraph()
        p_b.text = '• ' + b
        p_b.font.name = 'Segoe UI'
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = hex_to_rgb('#555555')
        p_b.alignment = PP_ALIGN.LEFT
        p_b.space_after = Pt(3)

# ==================== SLIDE 3: 3 CHEVRON BARS WITH LARGE LEFT DISCS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3, '100 Days\nAction Plan', align=PP_ALIGN.CENTER, x=8.0, y=2.8, w=4.8)

s3_data = [
    ('Days\n1-30', ['Understand the firm\'s infrastructure and ecosystem', 'Yearly agenda with stakeholders', 'Setup sales funnel for target achievement'], '#8CC63F', 0.80),
    ('Days\n30-60', ['Actionable, specific and implementation on existing plan', 'Dialogues with emerging prospects'], '#29B6D8', 2.80),
    ('Days\n60-100', ['Full engagement with clients and team', 'Work towards mutually accepted milestones'], '#FFAA00', 4.80),
]

for title, bullets, col, y in s3_data:
    c_rgb = hex_to_rgb(col)
    
    # 1. 3D Dark Shadow Fold under disc
    fld_pts = [(Inches(2.40), Inches(y + 0.10)), (Inches(3.30), Inches(y)), (Inches(3.30), Inches(y + 1.60)), (Inches(2.40), Inches(y + 1.50))]
    fb_fld = s3.shapes.build_freeform(fld_pts[0][0], fld_pts[0][1])
    fb_fld.add_line_segments(fld_pts[1:], close=True)
    fld = fb_fld.convert_to_shape()
    fld.fill.solid()
    fld.fill.fore_color.rgb = hex_to_rgb('#1A3300' if '8CC' in col else ('#0D3B4C' if '29B' in col else '#4D2A00'))
    fld.line.fill.background()
    
    # 2. Right-pointing Chevron Body
    chv = s3.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(3.20), Inches(y), Inches(4.70), Inches(1.60))
    chv.fill.solid()
    chv.fill.fore_color.rgb = c_rgb
    chv.line.fill.background()
    
    tf_chv = chv.text_frame
    tf_chv.word_wrap = True
    tf_chv.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_chv.margin_left = Inches(0.40)
    tf_chv.margin_right = Inches(0.70)
    for i, b in enumerate(bullets):
        p = tf_chv.paragraphs[0] if i == 0 else tf_chv.add_paragraph()
        p.text = '•  ' + b
        p.font.name = 'Segoe UI'
        p.font.size = Pt(10.5)
        p.font.color.rgb = hex_to_rgb('#FFFFFF')
        p.space_after = Pt(4)
        
    # 3. Large Left Circular Disc (Layered on top)
    disc_outer = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.90), Inches(y - 0.10), Inches(1.80), Inches(1.80))
    disc_outer.fill.solid()
    disc_outer.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc_outer.line.fill.background()
    
    disc = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.00), Inches(y), Inches(1.60), Inches(1.60))
    disc.fill.solid()
    disc.fill.fore_color.rgb = c_rgb
    disc.line.fill.background()
    
    tf_d = disc.text_frame
    tf_d.word_wrap = True
    tf_d.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = title.split('\n')
    p1 = tf_d.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf_d.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 4: 5 INTERSECTING VENN CIRCLES ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, '100 Days Action Plan')

s4_data = [
    ('0-20 Days', 'Understand the\ncompany culture', '#8CC63F', '01', 0.60),
    ('21-40 Days', 'Meetings with\ninternal teams', '#29B6D8', '02', 2.85),
    ('41-60 Days', 'Executing plan', '#FFAA00', '03', 5.10),
    ('61-80 Days', 'Full engagement\nwith client & team\non existing\nopportunities', '#EA4335', '04', 7.35),
    ('81-100 Days', 'Working towards\nmile stones', '#AF26FA', '05', 9.60),
]

diam_v = 2.85
for title, desc, col, num, x in s4_data:
    c_rgb = hex_to_rgb(col)
    y_v = 2.50
    
    # Outer colored border ring
    ring = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y_v), Inches(diam_v), Inches(diam_v))
    ring.fill.solid()
    ring.fill.fore_color.rgb = c_rgb
    ring.line.fill.background()
    
    # Inner white disc
    inner = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.12), Inches(y_v + 0.12), Inches(diam_v - 0.24), Inches(diam_v - 0.24))
    inner.fill.solid()
    inner.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    inner.line.fill.background()
    
    # Right Crescent Moon Intersection Badge
    cres = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + diam_v - 0.75), Inches(y_v + diam_v/2 - 0.55), Inches(0.90), Inches(1.10))
    cres.fill.solid()
    cres.fill.fore_color.rgb = c_rgb
    cres.line.color.rgb = hex_to_rgb('#FFFFFF')
    cres.line.width = Pt(2.0)
    p_num = cres.text_frame.paragraphs[0]
    p_num.text = num
    p_num.font.name = 'Montserrat'
    p_num.font.size = Pt(13)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER
    
    # Text inside disc
    tb = s4.shapes.add_textbox(Inches(x + 0.25), Inches(y_v + 0.65), Inches(diam_v - 0.95), Inches(1.50))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#222222')
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(4)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = 'Segoe UI'
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = hex_to_rgb('#666666')
    p_d.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 5: 5 LEAF CAPSULE CARDS + VERTICAL MULTI-COLOR AXIS ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5, '100 Days\nAction Plan', align=PP_ALIGN.CENTER, x=8.0, y=2.8, w=4.8)

# Vertical multi-colored timeline bar
colors = ['#8CC63F', '#29B6D8', '#FFAA00', '#EA4335', '#AF26FA']
for idx, col in enumerate(colors):
    seg = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.40), Inches(0.80 + idx * 1.15), Inches(0.08), Inches(1.15))
    seg.fill.solid()
    seg.fill.fore_color.rgb = hex_to_rgb(col)
    seg.line.fill.background()

# Curved bracket on left
arc = s5.shapes.add_shape(MSO_SHAPE.ARC, Inches(1.10), Inches(2.20), Inches(2.60), Inches(2.60))
arc.line.color.rgb = hex_to_rgb('#29B6D8')
arc.line.width = Pt(1.5)

tb_arc = s5.shapes.add_textbox(Inches(1.20), Inches(3.05), Inches(1.10), Inches(0.90))
tf_arc = tb_arc.text_frame
tf_arc.word_wrap = True
p_a1 = tf_arc.paragraphs[0]
p_a1.text = '100'
p_a1.font.name = 'Montserrat'
p_a1.font.size = Pt(13)
p_a1.font.bold = True
p_a1.font.color.rgb = hex_to_rgb('#222222')
p_a1.alignment = PP_ALIGN.CENTER
p_a2 = tf_arc.add_paragraph()
p_a2.text = 'Days'
p_a2.font.name = 'Montserrat'
p_a2.font.size = Pt(13)
p_a2.font.bold = True
p_a2.font.color.rgb = hex_to_rgb('#222222')
p_a2.alignment = PP_ALIGN.CENTER

s5_data = [
    ('First 20 Days', 'Understand the company culture', '#8CC63F', '01', 0.85),
    ('20-40 Days', 'Meetings with internal teams', '#29B6D8', '02', 2.00),
    ('40-60 Days', 'Executing plan', '#FFAA00', '03', 3.15),
    ('60-80 Days', 'Full engagement with client & team on\nexisting opportunities', '#EA4335', '04', 4.30),
    ('80-100 Days', 'Working towards mile stones', '#AF26FA', '05', 5.45),
]

for title, desc, col, num, y in s5_data:
    c_rgb = hex_to_rgb(col)
    
    # White rounded capsule base
    base = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.00), Inches(y), Inches(4.50), Inches(0.90))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    base.line.color.rgb = hex_to_rgb('#E2E8F0')
    base.line.width = Pt(1.5)
    
    # Left leaf-shaped colored tab
    leaf = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.00), Inches(y), Inches(1.10), Inches(0.90))
    leaf.fill.solid()
    leaf.fill.fore_color.rgb = c_rgb
    leaf.line.fill.background()
    p_num = leaf.text_frame.paragraphs[0]
    p_num.text = num
    p_num.font.name = 'Montserrat'
    p_num.font.size = Pt(14)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER
    
    # Axis Milestone circular node
    node = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.20), Inches(y + 0.20), Inches(0.48), Inches(0.48))
    node.fill.solid()
    node.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    node.line.color.rgb = hex_to_rgb('#E2E8F0')
    node.line.width = Pt(1.5)
    
    # Text inside capsule
    tb = s5.shapes.add_textbox(Inches(4.25), Inches(y + 0.12), Inches(3.10), Inches(0.70))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#222222')
    p_t.space_after = Pt(2)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = 'Segoe UI'
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = hex_to_rgb('#666666')

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\100_Days_Action_Plan\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Saved Folder 2 presentation to:', out_file)

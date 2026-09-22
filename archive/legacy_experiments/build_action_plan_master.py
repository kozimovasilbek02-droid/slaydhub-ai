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

def add_header(slide, title_text, subtitle_text=None, align=PP_ALIGN.CENTER, x=0.8, y=0.70, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = 'Montserrat'
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#2D3748')
    p.alignment = align
    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(13)
        p2.font.color.rgb = hex_to_rgb('#718096')
        p2.alignment = align
        p2.space_before = Pt(4)

# ==================== SLIDE 1: 5 DOWNWARD-POINTED CARDS ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, '100 Days Action Plan', 'In 5 Phases')

s1_data = [
    ('First 20\nDays', ['Understand the company culture', 'Outline current year commitment plans', 'Create yearly agenda with stakeholders'], '#8CC63F', 0.77),
    ('20-40\nDays', ['Meetings with internal teams', 'Build sales funnel'], '#29B6D8', 3.17),
    ('40-60\nDays', ['Executing plan', 'Formal meeting with clients', 'Dialogue with suitable prospects'], '#FFAA00', 5.57),
    ('60-80\nDays', ['Full engagement with client & team on existing opportunities'], '#EA4335', 7.97),
    ('80-100\nDays', ['Working towards mile stones'], '#AF26FA', 10.37),
]

cw = 2.18
ch = 3.90
y_top = 2.45

for title, bullets, col, x in s1_data:
    c_rgb = hex_to_rgb(col)
    
    # Soft Drop Shadow Base
    shadow = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.03), Inches(y_top + 0.04), Inches(cw), Inches(ch))
    shadow.fill.solid()
    shadow.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shadow.line.fill.background()
    
    # Main White Card
    card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y_top), Inches(cw), Inches(ch))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.color.rgb = hex_to_rgb('#EDF2F7')
    card.line.width = Pt(1.0)
    
    # Downward Pointed Top Banner
    pts_banner = [
        (Inches(x), Inches(y_top)),
        (Inches(x + cw), Inches(y_top)),
        (Inches(x + cw), Inches(y_top + 1.15)),
        (Inches(x + cw/2), Inches(y_top + 1.45)),
        (Inches(x), Inches(y_top + 1.15))
    ]
    fb = s1.shapes.build_freeform(pts_banner[0][0], pts_banner[0][1])
    fb.add_line_segments(pts_banner[1:], close=True)
    ban = fb.convert_to_shape()
    ban.fill.solid()
    ban.fill.fore_color.rgb = c_rgb
    ban.line.fill.background()
    
    # Banner Text
    tb_b = s1.shapes.add_textbox(Inches(x), Inches(y_top + 0.12), Inches(cw), Inches(0.95))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    lines = title.split('\n')
    p1 = tf_b.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf_b.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(13.5)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER
        
    # Bullets
    tb_t = s1.shapes.add_textbox(Inches(x + 0.16), Inches(y_top + 1.65), Inches(cw - 0.32), Inches(ch - 1.80))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    for i, b in enumerate(bullets):
        p = tf_t.paragraphs[0] if i == 0 else tf_t.add_paragraph()
        run_chk = p.add_run()
        run_chk.text = '✓  '
        run_chk.font.name = 'Segoe UI'
        run_chk.font.size = Pt(10)
        run_chk.font.bold = True
        run_chk.font.color.rgb = c_rgb
        run_txt = p.add_run()
        run_txt.text = b
        run_txt.font.name = 'Segoe UI'
        run_txt.font.size = Pt(10)
        run_txt.font.color.rgb = hex_to_rgb('#374151')
        p.space_after = Pt(8)

# ==================== SLIDE 2: 4 OVERLAPPING CIRCLES WITH NUMBER TABS ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, '100 Days Action Plan', 'In 4 Phases')

s2_data = [
    ('First 25 Days', ['Observe for understanding', 'Start opportunity assessment'], '#8CC63F', '01', False, 0.75),
    ('25-50 Days', ['Build month sales calendar', 'Sales reporting'], '#29B6D8', '02', True, 3.75),
    ('50-75 Days', ['Forecasting', 'Report assessment findings', 'Introduce changes in roles, responsibilities', 'Constructive feedback'], '#FFAA00', '03', False, 6.75),
    ('75-100 Days', ['Introduce changes in roles, responsibilities'], '#EA4335', '04', True, 9.75),
]

diam = 2.85
y_c = 2.50

for title, bullets, col, num, is_top_tab, x in s2_data:
    c_rgb = hex_to_rgb(col)
    
    # Soft Drop Shadow
    shd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.04), Inches(y_c + 0.04), Inches(diam), Inches(diam))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Outer Colored Border Ring
    ring = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y_c), Inches(diam), Inches(diam))
    ring.fill.solid()
    ring.fill.fore_color.rgb = c_rgb
    ring.line.fill.background()
    
    # Inner White Disc
    inner = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.08), Inches(y_c + 0.08), Inches(diam - 0.16), Inches(diam - 0.16))
    inner.fill.solid()
    inner.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    inner.line.fill.background()
    
    # Semicircular Number Tab (Layered on top of circumference)
    y_tab = y_c - 0.38 if is_top_tab else y_c + diam - 0.32
    tab = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + diam/2 - 0.35), Inches(y_tab), Inches(0.70), Inches(0.70))
    tab.fill.solid()
    tab.fill.fore_color.rgb = c_rgb
    tab.line.color.rgb = hex_to_rgb('#FFFFFF')
    tab.line.width = Pt(2.0)
    p_num = tab.text_frame.paragraphs[0]
    p_num.text = num
    p_num.font.name = 'Montserrat'
    p_num.font.size = Pt(12)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER
    
    # Text inside Circle
    tb = s2.shapes.add_textbox(Inches(x + 0.22), Inches(y_c + 0.45), Inches(diam - 0.44), Inches(diam - 0.90))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#2D3748')
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(6)
    
    for b in bullets:
        p_b = tf.add_paragraph()
        p_b.text = '• ' + b
        p_b.font.name = 'Segoe UI'
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = hex_to_rgb('#4A5568')
        p_b.alignment = PP_ALIGN.LEFT
        p_b.space_after = Pt(2)

# ==================== SLIDE 3: 3 CHEVRON BARS WITH LARGE LEFT DISCS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])

# Right Bold Title
tb_r = s3.shapes.add_textbox(Inches(8.00), Inches(2.60), Inches(4.80), Inches(2.00))
tf_r = tb_r.text_frame
tf_r.word_wrap = True
p_r1 = tf_r.paragraphs[0]
p_r1.text = '100 Days'
p_r1.font.name = 'Montserrat'
p_r1.font.size = Pt(34)
p_r1.font.bold = True
p_r1.font.color.rgb = hex_to_rgb('#2D3748')
p_r2 = tf_r.add_paragraph()
p_r2.text = 'Action Plan'
p_r2.font.name = 'Montserrat'
p_r2.font.size = Pt(34)
p_r2.font.bold = True
p_r2.font.color.rgb = hex_to_rgb('#2D3748')

s3_data = [
    ('Days\n1-30', ['Understand the firm\'s infrastructure and ecosystem', 'Yearly agenda with stakeholders', 'Setup sales funnel for target achievement'], '#8CC63F', 0.85),
    ('Days\n30-60', ['Actionable, specific and implementation on existing plan', 'Dialogues with emerging prospects'], '#29B6D8', 2.85),
    ('Days\n60-100', ['Full engagement with clients and team', 'Work towards mutually accepted milestones'], '#FFAA00', 4.85),
]

for title, bullets, col, y in s3_data:
    c_rgb = hex_to_rgb(col)
    
    # 3D Shadow Fold behind circle
    fld_pts = [(Inches(2.00), Inches(y + 0.15)), (Inches(2.90), Inches(y)), (Inches(2.90), Inches(y + 1.55)), (Inches(2.00), Inches(y + 1.40))]
    fb_fld = s3.shapes.build_freeform(fld_pts[0][0], fld_pts[0][1])
    fb_fld.add_line_segments(fld_pts[1:], close=True)
    fld = fb_fld.convert_to_shape()
    fld.fill.solid()
    fld.fill.fore_color.rgb = hex_to_rgb('#2E5309' if '8CC' in col else ('#0C4C5E' if '29B' in col else '#663C00'))
    fld.line.fill.background()
    
    # Right-pointing Chevron Body with Soft Drop Shadow
    shd_chv = s3.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(2.23), Inches(y + 0.04), Inches(5.60), Inches(1.55))
    shd_chv.fill.solid()
    shd_chv.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd_chv.line.fill.background()
    
    chv = s3.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(2.20), Inches(y), Inches(5.60), Inches(1.55))
    chv.fill.solid()
    chv.fill.fore_color.rgb = c_rgb
    chv.line.fill.background()
    
    tf_chv = chv.text_frame
    tf_chv.word_wrap = True
    tf_chv.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_chv.margin_left = Inches(1.10)
    tf_chv.margin_right = Inches(0.60)
    for i, b in enumerate(bullets):
        p = tf_chv.paragraphs[0] if i == 0 else tf_chv.add_paragraph()
        p.text = '•  ' + b
        p.font.name = 'Segoe UI'
        p.font.size = Pt(10.5)
        p.font.color.rgb = hex_to_rgb('#FFFFFF')
        p.space_after = Pt(4)
        
    # Large Left Disc (Outer White + Inner Color with Drop Shadow)
    shd_d = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.88), Inches(y - 0.06), Inches(1.70), Inches(1.70))
    shd_d.fill.solid()
    shd_d.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    shd_d.line.fill.background()
    
    disc_outer = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.85), Inches(y - 0.08), Inches(1.70), Inches(1.70))
    disc_outer.fill.solid()
    disc_outer.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc_outer.line.fill.background()
    
    disc = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.95), Inches(y + 0.02), Inches(1.50), Inches(1.50))
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
y_v = 2.50

# Pass 1: Draw all 5 circles and white inner discs
for title, desc, col, num, x in s4_data:
    c_rgb = hex_to_rgb(col)
    
    # Shadow
    shd = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.04), Inches(y_v + 0.04), Inches(diam_v), Inches(diam_v))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Outer colored border ring
    ring = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y_v), Inches(diam_v), Inches(diam_v))
    ring.fill.solid()
    ring.fill.fore_color.rgb = c_rgb
    ring.line.fill.background()
    
    # Inner white disc
    inner = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.10), Inches(y_v + 0.10), Inches(diam_v - 0.20), Inches(diam_v - 0.20))
    inner.fill.solid()
    inner.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    inner.line.fill.background()
    
    tb = s4.shapes.add_textbox(Inches(x + 0.20), Inches(y_v + 0.65), Inches(diam_v - 0.90), Inches(1.50))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#2D3748')
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(4)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = 'Segoe UI'
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = hex_to_rgb('#718096')
    p_d.alignment = PP_ALIGN.CENTER

# Pass 2: Draw all 5 crescent intersection badges ON TOP
for title, desc, col, num, x in s4_data:
    c_rgb = hex_to_rgb(col)
    cres = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + diam_v - 0.68), Inches(y_v + diam_v/2 - 0.48), Inches(0.80), Inches(0.96))
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

# ==================== SLIDE 5: 5 LEAF CAPSULE CARDS + VERTICAL AXIS ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])

# Right Bold Title
tb_r5 = s5.shapes.add_textbox(Inches(8.00), Inches(2.60), Inches(4.80), Inches(2.00))
tf_r5 = tb_r5.text_frame
tf_r5.word_wrap = True
p_r51 = tf_r5.paragraphs[0]
p_r51.text = '100 Days'
p_r51.font.name = 'Montserrat'
p_r51.font.size = Pt(34)
p_r51.font.bold = True
p_r51.font.color.rgb = hex_to_rgb('#2D3748')
p_r52 = tf_r5.add_paragraph()
p_r52.text = 'Action Plan'
p_r52.font.name = 'Montserrat'
p_r52.font.size = Pt(34)
p_r52.font.bold = True
p_r52.font.color.rgb = hex_to_rgb('#2D3748')

# Vertical multi-colored timeline bar
colors = ['#8CC63F', '#29B6D8', '#FFAA00', '#EA4335', '#AF26FA']
for idx, col in enumerate(colors):
    seg = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.40), Inches(0.85 + idx * 1.15), Inches(0.08), Inches(1.15))
    seg.fill.solid()
    seg.fill.fore_color.rgb = hex_to_rgb(col)
    seg.line.fill.background()

# Thin curved arc (open freeform curve)
arc_pts = [(Inches(2.30), Inches(1.70)), (Inches(1.00), Inches(3.50)), (Inches(2.30), Inches(5.30))]
fb_arc = s5.shapes.build_freeform(arc_pts[0][0], arc_pts[0][1])
fb_arc.add_line_segments(arc_pts[1:], close=False)
arc = fb_arc.convert_to_shape()
arc.fill.background()
arc.line.color.rgb = hex_to_rgb('#29B6D8')
arc.line.width = Pt(2.0)

tb_arc = s5.shapes.add_textbox(Inches(1.15), Inches(3.10), Inches(1.15), Inches(0.90))
tf_arc = tb_arc.text_frame
tf_arc.word_wrap = True
p_a1 = tf_arc.paragraphs[0]
p_a1.text = '100'
p_a1.font.name = 'Montserrat'
p_a1.font.size = Pt(13)
p_a1.font.bold = True
p_a1.font.color.rgb = hex_to_rgb('#2D3748')
p_a1.alignment = PP_ALIGN.CENTER
p_a2 = tf_arc.add_paragraph()
p_a2.text = 'Days'
p_a2.font.name = 'Montserrat'
p_a2.font.size = Pt(13)
p_a2.font.bold = True
p_a2.font.color.rgb = hex_to_rgb('#2D3748')
p_a2.alignment = PP_ALIGN.CENTER

s5_data = [
    ('First 20 Days', 'Understand the company culture', '#8CC63F', '01', 0.85),
    ('20-40 Days', 'Meetings with internal teams', '#29B6D8', '02', 2.00),
    ('40-60 Days', 'Executing plan', '#FFAA00', '03', 3.15),
    ('60-80 Days', 'Full engagement with client & team on existing opportunities', '#EA4335', '04', 4.30),
    ('80-100 Days', 'Working towards mile stones', '#AF26FA', '05', 5.45),
]

for title, desc, col, num, y in s5_data:
    c_rgb = hex_to_rgb(col)
    
    # Soft Drop Shadow
    shd = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.03), Inches(y + 0.04), Inches(4.50), Inches(0.90))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # White rounded capsule base
    base = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.00), Inches(y), Inches(4.50), Inches(0.90))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    base.line.color.rgb = hex_to_rgb('#EDF2F7')
    base.line.width = Pt(1.0)
    
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
    
    # Axis Milestone circular node with drop shadow
    node_shd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.22), Inches(y + 0.22), Inches(0.44), Inches(0.44))
    node_shd.fill.solid()
    node_shd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    node_shd.line.fill.background()
    
    node = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.20), Inches(y + 0.20), Inches(0.44), Inches(0.44))
    node.fill.solid()
    node.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    node.line.color.rgb = hex_to_rgb('#EDF2F7')
    node.line.width = Pt(1.5)
    
    # Text inside capsule
    tb = s5.shapes.add_textbox(Inches(4.25), Inches(y + 0.12), Inches(3.15), Inches(0.70))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#2D3748')
    p_t.space_after = Pt(2)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = 'Segoe UI'
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = hex_to_rgb('#718096')

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\100_Days_Action_Plan\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation saved to:', out_file)

import sys
import os
import math
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

def add_header(slide, title_text='10Vs Of Big Data', align=PP_ALIGN.CENTER, x=0.8, y=0.50, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.85))
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

v_colors_10 = [
    '#2D3748', # Value
    '#E53E3E', # Virality
    '#DD6B20', # Venue
    '#00A3C4', # Variability
    '#48BB78', # Volume
    '#805AD5', # Viscosity
    '#002B99', # Veracity
    '#C05621', # Visualization
    '#0987A0', # Velocity
    '#742A2A', # Variety
]

v_data_10 = [
    ('Value', 'Can you find the information you are looking for?', v_colors_10[0], '💎'),
    ('Virality', "Portable 'Aha' message for presentation or Instagram.", v_colors_10[1], '🕸'),
    ('Venue', 'Distributed Heterogeneous Data from multiple Platforms', v_colors_10[2], '📍'),
    ('Variability', 'Dynamic, Evolving Behavior in Data Source', v_colors_10[3], '📊'),
    ('Volume', 'Can you find it when you most need it?', v_colors_10[4], '🗄'),
    ('Viscosity', 'Does it stick with you? Does it call for action?', v_colors_10[5], '🧪'),
    ('Veracity', 'Are you dealing with information or disinformation?', v_colors_10[6], '🎯'),
    ('Visualization', 'Can you make sense at a glance? Does it trigger a decision?', v_colors_10[7], '👁'),
    ('Velocity', "Info accelerates, crises and chances unfold. Today's outlook?", v_colors_10[8], '⏱'),
    ('Variety', 'Image impact, multilingual context, information equilibrium check.', v_colors_10[9], '🗂'),
]

# ==================== SLIDE 1: 10 HORIZONTAL ROWS WITH LEFT TITLE ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Bold Title
tb_t1 = s1.shapes.add_textbox(Inches(0.80), Inches(2.60), Inches(4.20), Inches(2.00))
tf_t1 = tb_t1.text_frame
tf_t1.word_wrap = True
p1_1 = tf_t1.paragraphs[0]
p1_1.text = '10Vs Of'
p1_1.font.name = 'Montserrat'
p1_1.font.size = Pt(36)
p1_1.font.bold = True
p1_1.font.color.rgb = hex_to_rgb('#2D3748')
p1_2 = tf_t1.add_paragraph()
p1_2.text = 'Big Data'
p1_2.font.name = 'Montserrat'
p1_2.font.size = Pt(36)
p1_2.font.bold = True
p1_2.font.color.rgb = hex_to_rgb('#2D3748')

y_start = 0.50
row_h = 0.58
for i, (name, desc, col, icon) in enumerate(v_data_10):
    y = y_start + i * (row_h + 0.08)
    c_rgb = hex_to_rgb(col)
    
    # Shadow
    shd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.42), Inches(y + 0.02), Inches(7.20), Inches(row_h))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Base capsule
    base = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.40), Inches(y), Inches(7.20), Inches(row_h))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    base.line.color.rgb = c_rgb
    base.line.width = Pt(1.5)
    
    # Left Icon Badge
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.40), Inches(y), Inches(0.80), Inches(row_h))
    badge.fill.solid()
    badge.fill.fore_color.rgb = c_rgb
    badge.line.fill.background()
    p_ic = badge.text_frame.paragraphs[0]
    p_ic.text = icon
    p_ic.font.size = Pt(12)
    p_ic.alignment = PP_ALIGN.CENTER
    
    # Text
    tb_r = s1.shapes.add_textbox(Inches(6.35), Inches(y + 0.05), Inches(6.15), Inches(row_h - 0.10))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = tf_r.margin_bottom = 0
    p_nm = tf_r.paragraphs[0]
    p_nm.text = name
    p_nm.font.name = 'Montserrat'
    p_nm.font.size = Pt(10.5)
    p_nm.font.bold = True
    p_nm.font.color.rgb = hex_to_rgb('#2D3748')
    p_dc = tf_r.add_paragraph()
    p_dc.text = desc
    p_dc.font.name = 'Segoe UI'
    p_dc.font.size = Pt(9)
    p_dc.font.color.rgb = hex_to_rgb('#718096')

# ==================== SLIDE 2: CENTRAL ZIGZAG DIAMOND COLUMN ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, '10Vs Of Big Data')

cx_s2 = 6.66
y_start_s2 = 1.60
row_h_s2 = 0.52

for i, (name, desc, col, icon) in enumerate(v_data_10):
    y = y_start_s2 + i * (row_h_s2 + 0.04)
    c_rgb = hex_to_rgb(col)
    is_left = (i % 2 == 0)
    
    # Diamond in Center
    dia = s2.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(cx_s2 - 0.35), Inches(y - 0.02), Inches(0.70), Inches(0.56))
    dia.fill.solid()
    dia.fill.fore_color.rgb = c_rgb
    dia.line.color.rgb = hex_to_rgb('#FFFFFF')
    dia.line.width = Pt(1.5)
    p_d = dia.text_frame.paragraphs[0]
    p_d.text = icon
    p_d.font.size = Pt(10)
    p_d.alignment = PP_ALIGN.CENTER
    
    # Connecting Card (Left or Right)
    card_x = cx_s2 - 5.50 if is_left else cx_s2 + 0.30
    
    shd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(card_x + 0.02), Inches(y + 0.02), Inches(5.20), Inches(row_h_s2))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(card_x), Inches(y), Inches(5.20), Inches(row_h_s2))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    tb = s2.shapes.add_textbox(Inches(card_x + 0.20), Inches(y + 0.04), Inches(4.80), Inches(row_h_s2 - 0.08))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.RIGHT if is_left else PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.RIGHT if is_left else PP_ALIGN.LEFT

# ==================== SLIDE 3: CENTRAL HEXAGON HUB WITH 10 RADIAL SPOKES ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3, '10Vs Of Big Data')

left_y = [1.65, 2.65, 3.65, 4.65, 5.65]
right_y = [1.65, 2.65, 3.65, 4.65, 5.65]

# Center Hexagon Hub
hx_hub = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(5.56), Inches(3.05), Inches(2.20), Inches(2.00))
hx_hub.fill.solid()
hx_hub.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
hx_hub.line.color.rgb = hex_to_rgb('#FFFFFF')
hx_hub.line.width = Pt(2.0)
tf_hub = hx_hub.text_frame
tf_hub.vertical_anchor = MSO_ANCHOR.MIDDLE
p_h1 = tf_hub.paragraphs[0]
p_h1.text = '10Vs Of'
p_h1.font.name = 'Montserrat'
p_h1.font.size = Pt(14)
p_h1.font.bold = True
p_h1.font.color.rgb = hex_to_rgb('#2D3748')
p_h1.alignment = PP_ALIGN.CENTER
p_h2 = tf_hub.add_paragraph()
p_h2.text = 'Big Data'
p_h2.font.name = 'Montserrat'
p_h2.font.size = Pt(14)
p_h2.font.bold = True
p_h2.font.color.rgb = hex_to_rgb('#2D3748')
p_h2.alignment = PP_ALIGN.CENTER

for i in range(5):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = left_y[i]
    
    spk = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.30), Inches(y + 0.28), Inches(1.50), Inches(0.08))
    spk.fill.solid()
    spk.fill.fore_color.rgb = c_rgb
    spk.line.fill.background()
    
    hx = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(3.70), Inches(y), Inches(0.90), Inches(0.70))
    hx.fill.solid()
    hx.fill.fore_color.rgb = c_rgb
    hx.line.fill.background()
    p_i = hx.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(11)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s3.shapes.add_textbox(Inches(0.40), Inches(y), Inches(3.20), Inches(0.70))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.RIGHT

for i in range(5, 10):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = right_y[i - 5]
    
    spk = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.50), Inches(y + 0.28), Inches(1.50), Inches(0.08))
    spk.fill.solid()
    spk.fill.fore_color.rgb = c_rgb
    spk.line.fill.background()
    
    hx = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(8.70), Inches(y), Inches(0.90), Inches(0.70))
    hx.fill.solid()
    hx.fill.fore_color.rgb = c_rgb
    hx.line.fill.background()
    p_i = hx.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(11)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s3.shapes.add_textbox(Inches(9.70), Inches(y), Inches(3.20), Inches(0.70))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.LEFT

# ==================== SLIDE 4: CENTRAL DONUT RING + 10 CAPSULE CARDS + ELBOW LINES ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, '10Vs Of Big Data')

# Central Ring with Segments
ring_c = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.46), Inches(2.80), Inches(2.40), Inches(2.40))
ring_c.fill.solid()
ring_c.fill.fore_color.rgb = hex_to_rgb('#EDF2F7')
ring_c.line.color.rgb = hex_to_rgb('#CBD5E1')
ring_c.line.width = Pt(1.5)

tb_cnt = s4.shapes.add_textbox(Inches(5.66), Inches(3.40), Inches(2.00), Inches(1.00))
tf_cnt = tb_cnt.text_frame
tf_cnt.word_wrap = True
p_c1 = tf_cnt.paragraphs[0]
p_c1.text = '10Vs Of'
p_c1.font.name = 'Montserrat'
p_c1.font.size = Pt(14)
p_c1.font.bold = True
p_c1.font.color.rgb = hex_to_rgb('#2D3748')
p_c1.alignment = PP_ALIGN.CENTER
p_c2 = tf_cnt.add_paragraph()
p_c2.text = 'Big Data'
p_c2.font.name = 'Montserrat'
p_c2.font.size = Pt(14)
p_c2.font.bold = True
p_c2.font.color.rgb = hex_to_rgb('#2D3748')
p_c2.alignment = PP_ALIGN.CENTER

for i in range(5):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = left_y[i]
    
    # Connecting Elbow Line
    ln = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.70), Inches(y + 0.35), Inches(0.90), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(3.80), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(3.80), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = hex_to_rgb('#EDF2F7')
    crd.line.width = Pt(1.0)
    
    ic_b = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.00), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    ic_b.fill.solid()
    ic_b.fill.fore_color.rgb = c_rgb
    ic_b.line.color.rgb = hex_to_rgb('#FFFFFF')
    ic_b.line.width = Pt(1.5)
    p_i = ic_b.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(10)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s4.shapes.add_textbox(Inches(0.95), Inches(y + 0.08), Inches(2.95), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.RIGHT

for i in range(5, 10):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = right_y[i - 5]
    
    ln = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.70), Inches(y + 0.35), Inches(0.90), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.72), Inches(y + 0.02), Inches(3.80), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.70), Inches(y), Inches(3.80), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = hex_to_rgb('#EDF2F7')
    crd.line.width = Pt(1.0)
    
    ic_b = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.70), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    ic_b.fill.solid()
    ic_b.fill.fore_color.rgb = c_rgb
    ic_b.line.color.rgb = hex_to_rgb('#FFFFFF')
    ic_b.line.width = Pt(1.5)
    p_i = ic_b.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(10)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s4.shapes.add_textbox(Inches(9.40), Inches(y + 0.08), Inches(2.95), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.LEFT

# ==================== SLIDE 5: CENTRAL CIRCULAR HUB WITH 10 PILL SPOKES ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5, '10Vs Of Big Data')

# Center Circular Hub
hub5 = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.66), Inches(2.90), Inches(2.00), Inches(2.00))
hub5.fill.solid()
hub5.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
hub5.line.color.rgb = hex_to_rgb('#FFFFFF')
hub5.line.width = Pt(3.0)
tf_h5 = hub5.text_frame
tf_h5.vertical_anchor = MSO_ANCHOR.MIDDLE
p5_1 = tf_h5.paragraphs[0]
p5_1.text = '10Vs Of'
p5_1.font.name = 'Montserrat'
p5_1.font.size = Pt(14)
p5_1.font.bold = True
p5_1.font.color.rgb = hex_to_rgb('#2D3748')
p5_1.alignment = PP_ALIGN.CENTER
p5_2 = tf_h5.add_paragraph()
p5_2.text = 'Big Data'
p5_2.font.name = 'Montserrat'
p5_2.font.size = Pt(14)
p5_2.font.bold = True
p5_2.font.color.rgb = hex_to_rgb('#2D3748')
p5_2.alignment = PP_ALIGN.CENTER

for i in range(5):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = left_y[i]
    
    # Spoke connecting to hub
    spk = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.70), Inches(y + 0.32), Inches(1.10), Inches(0.06))
    spk.fill.solid()
    spk.fill.fore_color.rgb = c_rgb
    spk.line.fill.background()
    
    shd = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(4.00), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(4.00), Inches(0.70))
    pill.fill.solid()
    pill.fill.fore_color.rgb = c_rgb
    pill.line.fill.background()
    
    wd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.15), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    wd.fill.solid()
    wd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    wd.line.fill.background()
    p_i = wd.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(10)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s5.shapes.add_textbox(Inches(0.95), Inches(y + 0.08), Inches(3.10), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#FFFFFF')
    p2.alignment = PP_ALIGN.RIGHT

for i in range(5, 10):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = right_y[i - 5]
    
    spk = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.50), Inches(y + 0.32), Inches(1.10), Inches(0.06))
    spk.fill.solid()
    spk.fill.fore_color.rgb = c_rgb
    spk.line.fill.background()
    
    shd = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.52), Inches(y + 0.02), Inches(4.00), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.50), Inches(y), Inches(4.00), Inches(0.70))
    pill.fill.solid()
    pill.fill.fore_color.rgb = c_rgb
    pill.line.fill.background()
    
    wd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.55), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    wd.fill.solid()
    wd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    wd.line.fill.background()
    p_i = wd.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(10)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s5.shapes.add_textbox(Inches(9.25), Inches(y + 0.08), Inches(3.10), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#FFFFFF')
    p2.alignment = PP_ALIGN.LEFT

# ==================== SLIDE 6: 10 ROWS WITH TIMELINE NODES & FAR-RIGHT NUMBER PILLS ====================
s6 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Bold Title
tb_t6 = s6.shapes.add_textbox(Inches(0.80), Inches(2.60), Inches(4.00), Inches(2.00))
tf_t6 = tb_t6.text_frame
tf_t6.word_wrap = True
p6_1 = tf_t6.paragraphs[0]
p6_1.text = '10Vs Of'
p6_1.font.name = 'Montserrat'
p6_1.font.size = Pt(36)
p6_1.font.bold = True
p6_1.font.color.rgb = hex_to_rgb('#2D3748')
p6_2 = tf_t6.add_paragraph()
p6_2.text = 'Big Data'
p6_2.font.name = 'Montserrat'
p6_2.font.size = Pt(36)
p6_2.font.bold = True
p6_2.font.color.rgb = hex_to_rgb('#2D3748')

# Timeline Vertical Axis line
ax6 = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.10), Inches(0.60), Inches(0.02), Inches(6.20))
ax6.fill.solid()
ax6.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
ax6.line.fill.background()

for i, (name, desc, col, icon) in enumerate(v_data_10):
    y = y_start + i * (row_h + 0.08)
    c_rgb = hex_to_rgb(col)
    num_str = f"{i+1:02d}"
    
    # Node on Axis
    nd = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.96), Inches(y + 0.15), Inches(0.30), Inches(0.30))
    nd.fill.solid()
    nd.fill.fore_color.rgb = c_rgb
    nd.line.fill.background()
    
    shd = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.42), Inches(y + 0.02), Inches(7.20), Inches(row_h))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.40), Inches(y), Inches(7.20), Inches(row_h))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = hex_to_rgb('#EDF2F7')
    crd.line.width = Pt(1.0)
    
    ic_d = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.50), Inches(y + 0.06), Inches(0.46), Inches(0.46))
    ic_d.fill.solid()
    ic_d.fill.fore_color.rgb = c_rgb
    ic_d.line.fill.background()
    p_i = ic_d.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(9)
    p_i.alignment = PP_ALIGN.CENTER
    
    np = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.90), Inches(y), Inches(0.70), Inches(row_h))
    np.fill.solid()
    np.fill.fore_color.rgb = c_rgb
    np.line.fill.background()
    p_n = np.text_frame.paragraphs[0]
    p_n.text = num_str
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(11)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER
    
    tb = s6.shapes.add_textbox(Inches(6.05), Inches(y + 0.05), Inches(5.75), Inches(row_h - 0.10))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9)
    p2.font.color.rgb = hex_to_rgb('#718096')

# ==================== SLIDE 7: DONUT HUB WITH 10 OUTWARD ARROWS & FLOATING CARDS ====================
s7 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s7, '10Vs Of Big Data')

# Center Donut Hub
hub7 = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.66), Inches(2.90), Inches(2.00), Inches(2.00))
hub7.fill.solid()
hub7.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
hub7.line.color.rgb = hex_to_rgb('#CBD5E1')
hub7.line.width = Pt(3.0)
tf_h7 = hub7.text_frame
tf_h7.vertical_anchor = MSO_ANCHOR.MIDDLE
p7_1 = tf_h7.paragraphs[0]
p7_1.text = '10Vs Of'
p7_1.font.name = 'Montserrat'
p7_1.font.size = Pt(14)
p7_1.font.bold = True
p7_1.font.color.rgb = hex_to_rgb('#2D3748')
p7_1.alignment = PP_ALIGN.CENTER
p7_2 = tf_h7.add_paragraph()
p7_2.text = 'Big Data'
p7_2.font.name = 'Montserrat'
p7_2.font.size = Pt(14)
p7_2.font.bold = True
p7_2.font.color.rgb = hex_to_rgb('#2D3748')
p7_2.alignment = PP_ALIGN.CENTER

for i in range(5):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = left_y[i]
    
    # Connecting Line to Hub
    arr = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.70), Inches(y + 0.35), Inches(1.00), Inches(0.02))
    arr.fill.solid()
    arr.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    arr.line.fill.background()
    
    shd = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(3.80), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(3.80), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = hex_to_rgb('#EDF2F7')
    crd.line.width = Pt(1.0)
    
    tab = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.00), Inches(y), Inches(0.60), Inches(0.70))
    tab.fill.solid()
    tab.fill.fore_color.rgb = c_rgb
    tab.line.fill.background()
    
    ic_d = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.70), Inches(y + 0.08), Inches(0.54), Inches(0.54))
    ic_d.fill.solid()
    ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    ic_d.line.color.rgb = c_rgb
    ic_d.line.width = Pt(1.5)
    p_i = ic_d.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(9)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s7.shapes.add_textbox(Inches(0.95), Inches(y + 0.08), Inches(2.65), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.RIGHT

for i in range(5, 10):
    name, desc, col, icon = v_data_10[i]
    c_rgb = hex_to_rgb(col)
    y = right_y[i - 5]
    
    arr = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.60), Inches(y + 0.35), Inches(1.10), Inches(0.02))
    arr.fill.solid()
    arr.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    arr.line.fill.background()
    
    shd = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.72), Inches(y + 0.02), Inches(3.80), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.70), Inches(y), Inches(3.80), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = hex_to_rgb('#EDF2F7')
    crd.line.width = Pt(1.0)
    
    tab = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.70), Inches(y), Inches(0.60), Inches(0.70))
    tab.fill.solid()
    tab.fill.fore_color.rgb = c_rgb
    tab.line.fill.background()
    
    ic_d = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.05), Inches(y + 0.08), Inches(0.54), Inches(0.54))
    ic_d.fill.solid()
    ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    ic_d.line.color.rgb = c_rgb
    ic_d.line.width = Pt(1.5)
    p_i = ic_d.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(9)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s7.shapes.add_textbox(Inches(9.70), Inches(y + 0.08), Inches(2.65), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.LEFT

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10Vs_Of_Big_Data\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 3 saved to:', out_file)

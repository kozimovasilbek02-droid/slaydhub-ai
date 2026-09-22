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

def add_header(slide, title_text='10 Step Business Process', align=PP_ALIGN.CENTER, x=0.8, y=0.55, w=11.733):
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

# ==================== SLIDE 1: WINDING ROAD WITH 10 NODES ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1)

# Curved dashed road line across the slide
road_pts = [
    (Inches(1.0), Inches(2.8)),
    (Inches(3.3), Inches(2.6)),
    (Inches(5.6), Inches(2.6)),
    (Inches(7.9), Inches(2.6)),
    (Inches(10.2), Inches(2.8)),
    (Inches(12.2), Inches(3.6)),
    (Inches(10.2), Inches(4.5)),
    (Inches(7.9), Inches(4.7)),
    (Inches(5.6), Inches(4.7)),
    (Inches(3.3), Inches(4.7)),
    (Inches(1.0), Inches(5.4))
]
fb_rd = s1.shapes.build_freeform(road_pts[0][0], road_pts[0][1])
fb_rd.add_line_segments(road_pts[1:], close=False)
rd = fb_rd.convert_to_shape()
rd.fill.background()
rd.line.color.rgb = hex_to_rgb('#CBD5E1')
rd.line.width = Pt(2.5)

s1_steps_top = [
    ('Define\nObjectives', '#00A3C4', '🎯', 1.30, 2.50, False),
    ('Map\nStakeholders', '#2B6CB0', '👥', 3.60, 2.30, True),
    ('Plan\nResources', '#805AD5', '⚙️', 5.90, 2.30, False),
    ('Prepare\nData', '#E53E3E', '📊', 8.20, 2.30, True),
    ('Transfer\nInformation', '#DD6B20', '📄', 10.50, 2.50, False),
]

s1_steps_bot = [
    ('Connect\nPlatforms', '#DD6B20', '💻', 1.30, 4.70, False),
    ('Launch\nSystem', '#E53E3E', '🚀', 3.60, 4.70, True),
    ('Validate\nPerformance', '#D69E2E', '📋', 5.90, 4.70, False),
    ('Enable\nTeams', '#2B6CB0', '🤝', 8.20, 4.70, True),
    ('Drive\nAdoption', '#00A3C4', '💡', 10.50, 4.70, False),
]

for label, col, icon, x, y, is_text_above in s1_steps_top:
    c_rgb = hex_to_rgb(col)
    
    # Shadow
    shd = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.04), Inches(y + 0.04), Inches(0.90), Inches(0.90))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    shd.line.fill.background()
    
    node = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.90), Inches(0.90))
    node.fill.solid()
    node.fill.fore_color.rgb = c_rgb
    node.line.fill.background()
    p_i = node.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(13)
    p_i.alignment = PP_ALIGN.CENTER
    
    # Label text
    y_lbl = y - 0.90 if is_text_above else y + 0.95
    tb = s1.shapes.add_textbox(Inches(x - 0.45), Inches(y_lbl), Inches(1.80), Inches(0.80))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = label.split('\n')
    p1 = tf.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(10.5)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#2D3748')
        p2.alignment = PP_ALIGN.CENTER

for label, col, icon, x, y, is_text_above in s1_steps_bot:
    c_rgb = hex_to_rgb(col)
    
    shd = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.04), Inches(y + 0.04), Inches(0.90), Inches(0.90))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    shd.line.fill.background()
    
    node = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.90), Inches(0.90))
    node.fill.solid()
    node.fill.fore_color.rgb = c_rgb
    node.line.fill.background()
    p_i = node.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(13)
    p_i.alignment = PP_ALIGN.CENTER
    
    y_lbl = y - 0.90 if is_text_above else y + 0.95
    tb = s1.shapes.add_textbox(Inches(x - 0.45), Inches(y_lbl), Inches(1.80), Inches(0.80))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = label.split('\n')
    p1 = tf.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(10.5)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#2D3748')
        p2.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 2: CENTRAL CIRCULAR HUB + 10 CAPSULES ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2)

# Central Double-Ring Hub
hub2_shd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.82), Inches(2.17), Inches(3.70), Inches(3.70))
hub2_shd.fill.solid()
hub2_shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
hub2_shd.line.fill.background()

hub2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.80), Inches(2.15), Inches(3.70), Inches(3.70))
hub2.fill.solid()
hub2.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
hub2.line.color.rgb = hex_to_rgb('#CBD5E1')
hub2.line.width = Pt(3.0)

tb_h2 = s2.shapes.add_textbox(Inches(5.15), Inches(3.20), Inches(3.00), Inches(1.50))
tf_h2 = tb_h2.text_frame
tf_h2.word_wrap = True
p2_1 = tf_h2.paragraphs[0]
p2_1.text = 'Business'
p2_1.font.name = 'Montserrat'
p2_1.font.size = Pt(20)
p2_1.font.bold = True
p2_1.font.color.rgb = hex_to_rgb('#2D3748')
p2_1.alignment = PP_ALIGN.CENTER
p2_2 = tf_h2.add_paragraph()
p2_2.text = 'Process'
p2_2.font.name = 'Montserrat'
p2_2.font.size = Pt(20)
p2_2.font.bold = True
p2_2.font.color.rgb = hex_to_rgb('#2D3748')
p2_2.alignment = PP_ALIGN.CENTER

s2_left = [
    ('Market Analysis', '#00A3C4', '01'),
    ('Lead Generation', '#2B6CB0', '03'),
    ('Pitch Development', '#805AD5', '05'),
    ('Lead Nurturing', '#E53E3E', '07'),
    ('Objection Handling', '#DD6B20', '09'),
]

s2_right = [
    ('Audience Identification', '#DD6B20', '02'),
    ('Sales Insights', '#E53E3E', '04'),
    ('Sales Outreach', '#805AD5', '06'),
    ('Prospect Evaluation', '#2B6CB0', '08'),
    ('Deal Closure', '#00A3C4', '10'),
]

left_y_s2 = [1.50, 2.55, 3.60, 4.65, 5.70]
right_y_s2 = [1.50, 2.55, 3.60, 4.65, 5.70]

for i, (name, col, num) in enumerate(s2_left):
    c_rgb = hex_to_rgb(col)
    y = left_y_s2[i]
    
    ln = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.00), Inches(y + 0.35), Inches(0.90), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(3.40), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(3.40), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    # Left capsule cap
    cap = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(0.40), Inches(0.70))
    cap.fill.solid()
    cap.fill.fore_color.rgb = c_rgb
    cap.line.fill.background()
    
    # Right Number Disc
    nd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.55), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    nd.fill.solid()
    nd.fill.fore_color.rgb = c_rgb
    nd.line.color.rgb = hex_to_rgb('#FFFFFF')
    nd.line.width = Pt(1.5)
    p_n = nd.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(11)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER
    
    tb = s2.shapes.add_textbox(Inches(1.30), Inches(y + 0.12), Inches(2.15), Inches(0.50))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = name
    p.font.name = 'Montserrat'
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#2D3748')
    p.alignment = PP_ALIGN.RIGHT

for i, (name, col, num) in enumerate(s2_right):
    c_rgb = hex_to_rgb(col)
    y = right_y_s2[i]
    
    ln = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.40), Inches(y + 0.35), Inches(0.90), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.12), Inches(y + 0.02), Inches(3.40), Inches(0.70))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.10), Inches(y), Inches(3.40), Inches(0.70))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    cap = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(12.10), Inches(y), Inches(0.40), Inches(0.70))
    cap.fill.solid()
    cap.fill.fore_color.rgb = c_rgb
    cap.line.fill.background()
    
    nd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.15), Inches(y + 0.05), Inches(0.60), Inches(0.60))
    nd.fill.solid()
    nd.fill.fore_color.rgb = c_rgb
    nd.line.color.rgb = hex_to_rgb('#FFFFFF')
    nd.line.width = Pt(1.5)
    p_n = nd.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(11)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER
    
    tb = s2.shapes.add_textbox(Inches(9.85), Inches(y + 0.12), Inches(2.15), Inches(0.50))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = name
    p.font.name = 'Montserrat'
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#2D3748')
    p.alignment = PP_ALIGN.LEFT

# ==================== SLIDE 3: 10 ROUNDED BLOCK CARDS IN 2 ROWS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3)

# Looping connector ribbon behind cards
ribbon_pts = [
    (Inches(0.80), Inches(3.65)),
    (Inches(12.50), Inches(3.65))
]
fb_rb = s3.shapes.build_freeform(ribbon_pts[0][0], ribbon_pts[0][1])
fb_rb.add_line_segments(ribbon_pts[1:], close=False)
rb = fb_rb.convert_to_shape()
rb.fill.background()
rb.line.color.rgb = hex_to_rgb('#E2E8F0')
rb.line.width = Pt(12.0)

s3_row1 = [
    ('Market\nAnalysis', '#00A3C4', '#007A93', '01', 1.20),
    ('Audience\nIdentification', '#2B6CB0', '#1A4971', '02', 3.50),
    ('Lead\nGeneration', '#805AD5', '#553C9A', '03', 5.80),
    ('Sales\nInsights', '#E53E3E', '#9B2C2C', '04', 8.10),
    ('Pitch\nDevelopment', '#DD6B20', '#9C4221', '05', 10.40),
]

s3_row2 = [
    ('Sales\nOutreach', '#C05621', '#7B341E', '06', 1.20),
    ('Lead\nNurturing', '#E53E3E', '#9B2C2C', '07', 3.50),
    ('Prospect\nEvaluation', '#D69E2E', '#975A16', '08', 5.80),
    ('Objection\nHandling', '#2B6CB0', '#1A4971', '09', 8.10),
    ('Deal\nClosure', '#00A3C4', '#007A93', '10', 10.40),
]

# Row 1 (Top)
for title, col_top, col_bot, num, x in s3_row1:
    c_top = hex_to_rgb(col_top)
    c_bot = hex_to_rgb(col_bot)
    
    # Shadow
    shd = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.03), Inches(1.85 + 0.03), Inches(1.80), Inches(2.00))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Top Card Half (Rounded top)
    card_top = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.85), Inches(1.80), Inches(1.10))
    card_top.fill.solid()
    card_top.fill.fore_color.rgb = c_top
    card_top.line.fill.background()
    
    # Text in Top Half
    tb = s3.shapes.add_textbox(Inches(x + 0.10), Inches(2.00), Inches(1.60), Inches(0.80))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = title.split('\n')
    p1 = tf.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER
        
    # Bottom Half (Number)
    card_bot = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.95), Inches(1.80), Inches(0.90))
    card_bot.fill.solid()
    card_bot.fill.fore_color.rgb = c_bot
    card_bot.line.fill.background()
    p_n = card_bot.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(18)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER

# Row 2 (Bottom)
for title, col_top, col_bot, num, x in s3_row2:
    c_top = hex_to_rgb(col_top)
    c_bot = hex_to_rgb(col_bot)
    
    shd = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.03), Inches(4.30 + 0.03), Inches(1.80), Inches(2.00))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Top Half (Number)
    card_top = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(4.30), Inches(1.80), Inches(0.90))
    card_top.fill.solid()
    card_top.fill.fore_color.rgb = c_top
    card_top.line.fill.background()
    p_n = card_top.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(18)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER
    
    # Bottom Half (Text)
    card_bot = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(5.20), Inches(1.80), Inches(1.10))
    card_bot.fill.solid()
    card_bot.fill.fore_color.rgb = c_bot
    card_bot.line.fill.background()
    
    tb = s3.shapes.add_textbox(Inches(x + 0.10), Inches(5.35), Inches(1.60), Inches(0.80))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = title.split('\n')
    p1 = tf.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Step_Business_Process_Ppt_Editable_Free_Download\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 6 saved to:', out_file)

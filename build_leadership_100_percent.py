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

def add_header(slide, title_text='100 Day New Leadership Position Plan', align=PP_ALIGN.CENTER, x=0.8, y=0.50, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.85))
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

# ==================== SLIDE 1: 5 FOLDED TIERED CARDS ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, '100 Day New Leadership Position Plan')

# Layer 1: Horizontal Axis Line
ax = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(1.95), Inches(11.733), Inches(0.025))
ax.fill.solid()
ax.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
ax.line.fill.background()

s1_data = [
    ('First 20\nDays', ['Understand the company culture', 'Outline current year commitment plans', 'Create yearly agenda with stakeholders'], '#A80000', '#BA1B1B', 0.95),
    ('20-40\nDays', ['Meeting with internal teams', 'Build sales funnel'], '#881A4E', '#B32D68', 3.35),
    ('40-60\nDays', ['Executing plan', 'Formal meeting with a client', 'Dialog with the suitable prospect'], '#B38800', '#E8AC00', 5.75),
    ('60-80\nDays', ['First engagement with client and team on existing opportunities'], '#558B2F', '#7CB342', 8.15),
    ('80-100\nDays', ['Working towards milestones'], '#00695C', '#009688', 10.55),
]

cw = 2.15
for title, bullets, col_head, col_body, x in s1_data:
    # Layer 2: Dashed/Solid Stem Connector
    c_line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + cw/2 - 0.015), Inches(2.05), Inches(0.03), Inches(0.60))
    c_line.fill.solid()
    c_line.fill.fore_color.rgb = hex_to_rgb(col_body)
    c_line.line.fill.background()
    
    # Layer 3: Card Body (Lighter, with rounded bottom)
    body = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(3.20), Inches(cw), Inches(3.40))
    body.fill.solid()
    body.fill.fore_color.rgb = hex_to_rgb(col_body)
    body.line.fill.background()
    
    tf_b = body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = Inches(0.18)
    tf_b.margin_right = Inches(0.18)
    tf_b.margin_top = Inches(0.35)
    for i, b_text in enumerate(bullets):
        p_b = tf_b.paragraphs[0] if i == 0 else tf_b.add_paragraph()
        p_b.text = '•  ' + b_text
        p_b.font.name = 'Segoe UI'
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_b.space_after = Pt(8)
        
    # Layer 4: Top Folded Tab (Darker, sitting overlapping on top of body)
    tab = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.45), Inches(cw), Inches(0.95))
    tab.fill.solid()
    tab.fill.fore_color.rgb = hex_to_rgb(col_head)
    tab.line.fill.background()
    
    tf_tab = tab.text_frame
    tf_tab.word_wrap = True
    tf_tab.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_tab.paragraphs[0]
    p_t.text = title.replace('\n', ' ')
    p_t.font.name = 'Montserrat'
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_t.alignment = PP_ALIGN.CENTER
    
    # Layer 5: Axis Milestone Dot on Top
    dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + cw/2 - 0.12), Inches(1.83), Inches(0.24), Inches(0.24))
    dot.fill.solid()
    dot.fill.fore_color.rgb = hex_to_rgb(col_body)
    dot.line.fill.background()

# ==================== SLIDE 2: 8 TILES + SEESAW FULCRUM ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, '100 Day New Leadership Position Plan')

# Top Chevron Banner
top_arr = s2.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(1.00), Inches(2.15), Inches(11.333), Inches(0.70))
top_arr.fill.solid()
top_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
top_arr.line.fill.background()
p_ta = top_arr.text_frame.paragraphs[0]
p_ta.text = "The New Leader's 100- Day Action Plan"
p_ta.font.name = 'Montserrat'
p_ta.font.size = Pt(14)
p_ta.font.bold = True
p_ta.font.color.rgb = hex_to_rgb('#222222')
p_ta.alignment = PP_ALIGN.CENTER

# Bottom Chevron Banner (Layer underneath see-saw)
bot_arr = s2.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0.60), Inches(5.75), Inches(12.133), Inches(0.70))
bot_arr.fill.solid()
bot_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
bot_arr.line.fill.background()
p_ba = bot_arr.text_frame.paragraphs[0]
p_ba.text = 'Activate Ongoing Communication'
p_ba.font.name = 'Montserrat'
p_ba.font.size = Pt(14)
p_ba.font.bold = True
p_ba.font.color.rgb = hex_to_rgb('#222222')
p_ba.alignment = PP_ALIGN.CENTER

# Center Fulcrum Triangle (Layered above bottom chevron banner)
tri_pts = [(Inches(6.66), Inches(4.90)), (Inches(7.30), Inches(6.00)), (Inches(6.02), Inches(6.00))]
tb_t = s2.shapes.build_freeform(tri_pts[0][0], tri_pts[0][1])
tb_t.add_line_segments(tri_pts[1:], close=True)
tri = tb_t.convert_to_shape()
tri.fill.solid()
tri.fill.fore_color.rgb = hex_to_rgb('#888888')
tri.line.fill.background()

s2_tiles = [
    ('Position\nyourself for\nsuccess', '#BA1B1B', 0.90),
    ('Leverage\nthe Fuzzy\nfront end', '#B32D68', 2.25),
    ('Take\ncontrol\nof Day one', '#FFC107', 3.60),
    ('Co-create\nBurning\nImperative', '#7CB342', 4.95),
    ('Embed\nMilestones', '#009688', 6.30),
    ('Jumpstart\nEarly wins', '#00A3E0', 7.65),
    ('Complete\norganization\nRole sort', '#0055B8', 9.00),
    ('Evolve\nleadership,\npractices,\nculture', '#4B53A0', 10.35),
]

for title, col, x in s2_tiles:
    pts = [(Inches(x + 0.35), Inches(3.00)), (Inches(x + 1.65), Inches(3.00)), (Inches(x + 1.30), Inches(5.30)), (Inches(x), Inches(5.30))]
    fb = s2.shapes.build_freeform(pts[0][0], pts[0][1])
    fb.add_line_segments(pts[1:], close=True)
    poly = fb.convert_to_shape()
    poly.fill.solid()
    poly.fill.fore_color.rgb = hex_to_rgb(col)
    poly.line.color.rgb = hex_to_rgb('#FFFFFF')
    poly.line.width = Pt(1.5)
    
    tb = s2.shapes.add_textbox(Inches(x + 0.05), Inches(3.15), Inches(1.50), Inches(2.00))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = 'Segoe UI'
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 3: 6 CAPSULES WITH CURVED SWOOP & VECTOR ICONS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Upward Arrow (Action Plan)
l_arr = s3.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(0.50), Inches(0.80), Inches(0.55), Inches(6.00))
l_arr.fill.solid()
l_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
l_arr.line.fill.background()

# Middle Downward Arrow (Activate and direct ongoing communication)
m_arr = s3.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(7.00), Inches(0.80), Inches(0.55), Inches(6.00))
m_arr.fill.solid()
m_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
m_arr.line.fill.background()

add_header(s3, '100 Day New\nLeadership\nPosition Plan', align=PP_ALIGN.CENTER, x=7.8, y=2.6, w=5.0)

s3_cards = [
    ('Day\n01', 'Take control of day one.', '#BA1B1B', 0.60, 'CALENDAR'),
    ('Day\n30', 'Embed burning imperative', '#B32D68', 1.65, 'CODE'),
    ('Day\n45', 'Exploit key milestones', '#FFC107', 2.70, 'FLAG'),
    ('Day\n60', 'Invest in early wins', '#7CB342', 3.75, 'TROPHY'),
    ('Day\n70', 'Secure skillful employees in the right roles', '#009688', 4.80, 'TEAM'),
    ('Day\n100', 'Expand people, plans, and practices', '#00A3E0', 5.85, 'REPORT'),
]

for day_txt, desc, col, y, icon_type in s3_cards:
    c_rgb = hex_to_rgb(col)
    
    # 1. Outer White Capsule Base with Drop Shadow
    base = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.45), Inches(y), Inches(5.15), Inches(0.90))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    base.line.color.rgb = hex_to_rgb('#E2E8F0')
    base.line.width = Pt(1.5)
    
    # 2. Left Curved Wave Swoop (Integrated seamless colored tab)
    pts_tab = [
        (Inches(1.45), Inches(y)),
        (Inches(2.40), Inches(y)),
        (Inches(2.15), Inches(y + 0.90)),
        (Inches(1.45), Inches(y + 0.90))
    ]
    fb_tab = s3.shapes.build_freeform(pts_tab[0][0], pts_tab[0][1])
    fb_tab.add_line_segments(pts_tab[1:], close=True)
    tab_shape = fb_tab.convert_to_shape()
    tab_shape.fill.solid()
    tab_shape.fill.fore_color.rgb = c_rgb
    tab_shape.line.fill.background()
    
    # Left Rounded Edge to cover corner
    l_cap = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.45), Inches(y), Inches(0.50), Inches(0.90))
    l_cap.fill.solid()
    l_cap.fill.fore_color.rgb = c_rgb
    l_cap.line.fill.background()
    
    tb_d = s3.shapes.add_textbox(Inches(1.45), Inches(y + 0.10), Inches(0.85), Inches(0.70))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    tf_d.margin_left = tf_d.margin_right = tf_d.margin_top = tf_d.margin_bottom = 0
    lines = day_txt.split('\n')
    p1 = tf_d.paragraphs[0]
    p1.text = lines[0]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    if len(lines) > 1:
        p2 = tf_d.add_paragraph()
        p2.text = lines[1]
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER
        
    # Text Description
    tb = s3.shapes.add_textbox(Inches(2.55), Inches(y + 0.15), Inches(2.95), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = desc
    p.font.name = 'Segoe UI'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#222222')
    
    # Right Circular Icon Badge with Pure Vector Icons
    ic = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.70), Inches(y + 0.08), Inches(0.74), Inches(0.74))
    ic.fill.solid()
    ic.fill.fore_color.rgb = c_rgb
    ic.line.color.rgb = hex_to_rgb('#FFFFFF')
    ic.line.width = Pt(2.5)
    
    # Vector Icon drawing inside badge
    cx = 6.07
    cy = y + 0.45
    if icon_type == 'CALENDAR':
        cal = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - 0.15), Inches(cy - 0.15), Inches(0.30), Inches(0.30))
        cal.line.color.rgb = hex_to_rgb('#FFFFFF')
        cal.line.width = Pt(1.2)
        cal.fill.fill_type = 0
    elif icon_type == 'CODE':
        wnd = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - 0.16), Inches(cy - 0.14), Inches(0.32), Inches(0.28))
        wnd.line.color.rgb = hex_to_rgb('#FFFFFF')
        wnd.line.width = Pt(1.2)
        wnd.fill.fill_type = 0
    elif icon_type == 'FLAG':
        flg = s3.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(cx - 0.12), Inches(cy - 0.16), Inches(0.24), Inches(0.18))
        flg.fill.solid()
        flg.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        flg.line.fill.background()
    elif icon_type == 'TROPHY':
        trp = s3.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(cx - 0.14), Inches(cy - 0.16), Inches(0.28), Inches(0.22))
        trp.fill.solid()
        trp.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        trp.line.fill.background()
    elif icon_type == 'TEAM':
        dot = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.08), Inches(cy - 0.14), Inches(0.16), Inches(0.16))
        dot.fill.solid()
        dot.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        dot.line.fill.background()
    elif icon_type == 'REPORT':
        doc = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.12), Inches(cy - 0.16), Inches(0.24), Inches(0.32))
        doc.line.color.rgb = hex_to_rgb('#FFFFFF')
        doc.line.width = Pt(1.2)
        doc.fill.fill_type = 0

# ==================== SLIDE 4: 6 SEAMLESS TOUCHING INTERLOCKING CARDS ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, '100 Day New Leadership Position Plan')

s4_cards = [
    ('Before day 0', 'Due Diligence', 'Choose a\nwinnable\ngame', '#BA1B1B'),
    ('Day 0', 'Develop', 'Prepare for\ntakeoff', '#B32D68'),
    ('Day 1', 'Demonstrate', 'Be the\nmanifesto', '#7CB342'),
    ('Days 2-30', 'Discover', 'Find the\nrallying cry', '#009688'),
    ('Days 31-60', 'Decide', 'Make your\nmove', '#00A3E0'),
    ('Days 61-100', 'Deliver', 'Keep it up', '#0055B8'),
]

# Total width available: 11.80 inches. Each card width = 1.96 inches, seamlessly side by side!
start_x = 0.75
card_w = 1.96
card_h = 1.75
tab_w = 0.22
tab_h = 0.46
y_card = 2.45

for i, (d_txt, sub_d, sub_b, col) in enumerate(s4_cards):
    x = start_x + i * card_w
    c_rgb = hex_to_rgb(col)
    
    # 1. Main Rounded Card Body
    card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y_card), Inches(card_w - 0.02), Inches(card_h))
    card.fill.solid()
    card.fill.fore_color.rgb = c_rgb
    card.line.fill.background()
    
    # 2. Right Arrow Tab (Interlocking into next card)
    tab_pts = [
        (Inches(x + card_w - 0.03), Inches(y_card + card_h/2 - tab_h/2)),
        (Inches(x + card_w + tab_w), Inches(y_card + card_h/2)),
        (Inches(x + card_w - 0.03), Inches(y_card + card_h/2 + tab_h/2))
    ]
    tb_tab = s4.shapes.build_freeform(tab_pts[0][0], tab_pts[0][1])
    tb_tab.add_line_segments(tab_pts[1:], close=True)
    tab_shape = tb_tab.convert_to_shape()
    tab_shape.fill.solid()
    tab_shape.fill.fore_color.rgb = c_rgb
    tab_shape.line.fill.background()
    
    # Text inside Card
    tf_c = card.text_frame
    tf_c.word_wrap = True
    tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_c.margin_left = Inches(0.10 if i == 0 else 0.22)
    tf_c.margin_right = Inches(0.10)
    p1 = tf_c.paragraphs[0]
    p1.text = d_txt
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf_c.add_paragraph()
    p2.text = sub_d
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = hex_to_rgb('#FFFFFF')
    p2.alignment = PP_ALIGN.CENTER
    
    # 3. Vertical Connector Stem
    c_stem = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + card_w/2 - 0.05), Inches(y_card + card_h), Inches(0.10), Inches(0.40))
    c_stem.fill.solid()
    c_stem.fill.fore_color.rgb = c_rgb
    c_stem.line.fill.background()
    
    # 4. Subtitle Text Box
    tb_sub = s4.shapes.add_textbox(Inches(x - 0.05), Inches(y_card + card_h + 0.45), Inches(card_w + 0.10), Inches(1.10))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = tf_sub.margin_right = tf_sub.margin_top = tf_sub.margin_bottom = 0
    p_s = tf_sub.paragraphs[0]
    p_s.text = sub_b.replace('\n', ' ')
    p_s.font.name = 'Segoe UI'
    p_s.font.size = Pt(10.5)
    p_s.font.color.rgb = hex_to_rgb('#222222')
    p_s.alignment = PP_ALIGN.CENTER
    
    # 5. Bottom Accent Line
    b_bar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(6.15), Inches(card_w - 0.02), Inches(0.14))
    b_bar.fill.solid()
    b_bar.fill.fore_color.rgb = c_rgb
    b_bar.line.fill.background()

# ==================== SLIDE 5: 6 HEXAGON MARKERS ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5, '100 Day New Leadership Position Plan')

s5_hexagons = [
    ('Before day 0', 'Due Diligence', 'Choose a\nwinnable\ngame', '#BA1B1B', 0.60),
    ('Day 0', 'Develop', 'Prepare for\ntakeoff', '#B32D68', 2.65),
    ('Day 1', 'Demonstrate', 'Be the\nmanifesto', '#FFC107', 4.70),
    ('Days 2-30', 'Discover', 'Find the\nrallying cry', '#7CB342', 6.75),
    ('Days 31-60', 'Decide', 'Make your\nmove', '#009688', 8.80),
    ('Days 61-100', 'Deliver', 'Keep it up', '#00A3E0', 10.85),
]

for d_txt, sub_txt, bot_txt, col, x in s5_hexagons:
    pts_hx = [
        (Inches(x + 0.40), Inches(2.20)),
        (Inches(x + 1.40), Inches(2.20)),
        (Inches(x + 1.80), Inches(3.05)),
        (Inches(x + 0.90), Inches(4.30)),
        (Inches(x), Inches(3.05))
    ]
    fb = s5.shapes.build_freeform(pts_hx[0][0], pts_hx[0][1])
    fb.add_line_segments(pts_hx[1:], close=True)
    hx = fb.convert_to_shape()
    hx.fill.solid()
    hx.fill.fore_color.rgb = hex_to_rgb(col)
    hx.line.fill.background()
    
    iw = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.20), Inches(2.35), Inches(1.40), Inches(1.40))
    iw.fill.solid()
    iw.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    iw.line.fill.background()
    
    tf_h = iw.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf_h.paragraphs[0]
    p1.text = d_txt
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#222222')
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf_h.add_paragraph()
    p2.text = sub_txt
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = hex_to_rgb('#666666')
    p2.alignment = PP_ALIGN.CENTER
    
    arch_pts = [(Inches(x + 0.45), Inches(4.30)), (Inches(x + 0.90), Inches(4.55)), (Inches(x + 1.35), Inches(4.30))]
    fb_a = s5.shapes.build_freeform(arch_pts[0][0], arch_pts[0][1])
    fb_a.add_line_segments(arch_pts[1:], close=False)
    arch = fb_a.convert_to_shape()
    arch.fill.background()
    arch.line.color.rgb = hex_to_rgb(col)
    arch.line.width = Pt(2.0)
    
    brk = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.10), Inches(4.80), Inches(1.60), Inches(1.20))
    brk.fill.solid()
    brk.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    brk.line.color.rgb = hex_to_rgb(col)
    brk.line.width = Pt(1.5)
    
    tf_b = brk.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_b.paragraphs[0]
    p_b.text = bot_txt.replace('\n', ' ')
    p_b.font.name = 'Segoe UI'
    p_b.font.size = Pt(10.5)
    p_b.font.color.rgb = hex_to_rgb('#222222')
    p_b.alignment = PP_ALIGN.CENTER

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\100_Day_New_Leadership_Position_Plan\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation cleanly written to:', out_file)

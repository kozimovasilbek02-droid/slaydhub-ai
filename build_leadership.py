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

def add_header(slide, title_text='100 Day New Leadership Position Plan', align=PP_ALIGN.CENTER, x=0.8, y=0.55, w=11.733):
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

# ==================== SLIDE 1: 5 TIERED TIMELINE CARDS ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, '100 Day New Leadership Position Plan')

ax = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.90), Inches(1.95), Inches(11.533), Inches(0.03))
ax.fill.solid()
ax.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
ax.line.fill.background()

s1_data = [
    ('First 20\nDays', ['Understand the company culture', 'Outline current year commitment plans', 'Create yearly agenda with stakeholders'], '#A80000', '#BA1B1B', 1.05),
    ('20-40\nDays', ['Meeting with internal teams', 'Build sales funnel'], '#881A4E', '#B32D68', 3.45),
    ('40-60\nDays', ['Executing plan', 'Formal meeting with a client', 'Dialog with the suitable prospect'], '#B38800', '#E8AC00', 5.85),
    ('60-80\nDays', ['First engagement with client and team on existing opportunities'], '#558B2F', '#7CB342', 8.25),
    ('80-100\nDays', ['Working towards milestones'], '#00695C', '#009688', 10.65),
]

for title, bullets, col_head, col_body, x in s1_data:
    dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.90), Inches(1.85), Inches(0.24), Inches(0.24))
    dot.fill.solid()
    dot.fill.fore_color.rgb = hex_to_rgb(col_body)
    dot.line.fill.background()
    
    c_line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 1.01), Inches(2.09), Inches(0.02), Inches(0.60))
    c_line.fill.solid()
    c_line.fill.fore_color.rgb = hex_to_rgb(col_body)
    c_line.line.fill.background()
    
    tab = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.65), Inches(2.05), Inches(0.95))
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
    
    body = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(3.40), Inches(2.05), Inches(3.20))
    body.fill.solid()
    body.fill.fore_color.rgb = hex_to_rgb(col_body)
    body.line.fill.background()
    
    tf_b = body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = Inches(0.15)
    tf_b.margin_right = Inches(0.15)
    tf_b.margin_top = Inches(0.20)
    for i, b_text in enumerate(bullets):
        p_b = tf_b.paragraphs[0] if i == 0 else tf_b.add_paragraph()
        p_b.text = '•  ' + b_text
        p_b.font.name = 'Segoe UI'
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_b.space_after = Pt(8)

# ==================== SLIDE 2: 8 ANGLED PARALLELOGRAM TILES ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2, '100 Day New Leadership Position Plan')

top_arr = s2.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(1.30), Inches(2.25), Inches(10.733), Inches(0.65))
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

s2_tiles = [
    ('Position\nyourself for\nsuccess', '#BA1B1B', 1.30),
    ('Leverage\nthe Fuzzy\nfront end', '#B32D68', 2.62),
    ('Take\ncontrol\nof Day one', '#FFC107', 3.94),
    ('Co-create\nBurning\nImperative', '#7CB342', 5.26),
    ('Embed\nMilestones', '#009688', 6.58),
    ('Jumpstart\nEarly wins', '#00A3E0', 7.90),
    ('Complete\norganization\nRole sort', '#0055B8', 9.22),
    ('Evolve\nleadership,\npractices,\nculture', '#4B53A0', 10.54),
]

for title, col, x in s2_tiles:
    pts = [(Inches(x + 0.35), Inches(3.05)), (Inches(x + 1.60), Inches(3.05)), (Inches(x + 1.25), Inches(5.35)), (Inches(x), Inches(5.35))]
    fb = s2.shapes.build_freeform(pts[0][0], pts[0][1])
    fb.add_line_segments(pts[1:], close=True)
    poly = fb.convert_to_shape()
    poly.fill.solid()
    poly.fill.fore_color.rgb = hex_to_rgb(col)
    poly.line.color.rgb = hex_to_rgb('#FFFFFF')
    poly.line.width = Pt(1.5)
    
    tb = s2.shapes.add_textbox(Inches(x + 0.05), Inches(3.20), Inches(1.45), Inches(2.00))
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

tri_pts = [(Inches(5.80), Inches(5.50)), (Inches(6.40), Inches(6.40)), (Inches(5.20), Inches(6.40))]
tb_t = s2.shapes.build_freeform(tri_pts[0][0], tri_pts[0][1])
tb_t.add_line_segments(tri_pts[1:], close=True)
tri = tb_t.convert_to_shape()
tri.fill.solid()
tri.fill.fore_color.rgb = hex_to_rgb('#A6ACAF')
tri.line.fill.background()

bot_arr = s2.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0.90), Inches(5.65), Inches(11.533), Inches(0.65))
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

# ==================== SLIDE 3: 6 HORIZONTAL CAPSULE CARDS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Upward Arrow with text
l_arr = s3.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(0.50), Inches(0.80), Inches(0.55), Inches(6.00))
l_arr.fill.solid()
l_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
l_arr.line.fill.background()

# Middle Downward Arrow
m_arr = s3.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.90), Inches(0.80), Inches(0.55), Inches(6.00))
m_arr.fill.solid()
m_arr.fill.fore_color.rgb = hex_to_rgb('#BDC3C7')
m_arr.line.fill.background()

add_header(s3, '100 Day New\nLeadership\nPosition Plan', align=PP_ALIGN.CENTER, x=7.8, y=2.7, w=5.0)

s3_cards = [
    ('Day 01', 'Take control of day one.', '#BA1B1B', 0.60, '📅'),
    ('Day 30', 'Embed burning imperative', '#B32D68', 1.65, '💻'),
    ('Day 45', 'Exploit key milestones', '#FFC107', 2.70, '🚩'),
    ('Day 60', 'Invest in early wins', '#7CB342', 3.75, '🏆'),
    ('Day 70', 'Secure skillful employees in the right roles', '#009688', 4.80, '👥'),
    ('Day 100', 'Expand people, plans, and practices', '#00A3E0', 5.85, '📋'),
]

for day, desc, col, y, icon in s3_cards:
    base = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.50), Inches(y), Inches(5.00), Inches(0.90))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    base.line.color.rgb = hex_to_rgb('#E2E8F0')
    base.line.width = Pt(1.5)
    
    bdg = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.50), Inches(y), Inches(1.15), Inches(0.90))
    bdg.fill.solid()
    bdg.fill.fore_color.rgb = hex_to_rgb(col)
    bdg.line.fill.background()
    tf_d = bdg.text_frame
    tf_d.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_d = tf_d.paragraphs[0]
    p_d.text = day
    p_d.font.name = 'Montserrat'
    p_d.font.size = Pt(12)
    p_d.font.bold = True
    p_d.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_d.alignment = PP_ALIGN.CENTER
    
    tb = s3.shapes.add_textbox(Inches(2.80), Inches(y + 0.15), Inches(2.70), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.name = 'Segoe UI'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#222222')
    
    ic = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.65), Inches(y + 0.08), Inches(0.74), Inches(0.74))
    ic.fill.solid()
    ic.fill.fore_color.rgb = hex_to_rgb(col)
    ic.line.color.rgb = hex_to_rgb('#FFFFFF')
    ic.line.width = Pt(2.0)
    p_i = ic.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(14)
    p_i.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 4: 6 INTERCONNECTED CARDS WITH ARROW HEADS ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, '100 Day New Leadership Position Plan')

s4_cards = [
    ('Before day 0', 'Due Diligence', 'Choose a\nwinnable\ngame', '#BA1B1B', 0.55),
    ('Day 0', 'Develop', 'Prepare for\ntakeoff', '#B32D68', 2.60),
    ('Day 1', 'Demonstrate', 'Be the\nmanifesto', '#7CB342', 4.65),
    ('Days 2-30', 'Discover', 'Find the\nrallying cry', '#009688', 6.70),
    ('Days 31-60', 'Decide', 'Make your\nmove', '#00A3E0', 8.75),
    ('Days 61-100', 'Deliver', 'Keep it up', '#0055B8', 10.80),
]

card_w = 1.80
card_h = 1.70
tab_w = 0.22
tab_h = 0.44

for i, (d_txt, sub_d, sub_b, col, x) in enumerate(s4_cards):
    y = 2.45
    c_rgb = hex_to_rgb(col)
    
    card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(card_w), Inches(card_h))
    card.fill.solid()
    card.fill.fore_color.rgb = c_rgb
    card.line.fill.background()
    
    tab_pts = [
        (Inches(x + card_w - 0.02), Inches(y + card_h/2 - tab_h/2)),
        (Inches(x + card_w + tab_w), Inches(y + card_h/2)),
        (Inches(x + card_w - 0.02), Inches(y + card_h/2 + tab_h/2))
    ]
    tb_tab = s4.shapes.build_freeform(tab_pts[0][0], tab_pts[0][1])
    tb_tab.add_line_segments(tab_pts[1:], close=True)
    tab_shape = tb_tab.convert_to_shape()
    tab_shape.fill.solid()
    tab_shape.fill.fore_color.rgb = c_rgb
    tab_shape.line.fill.background()
    
    if i > 0:
        notch_pts = [
            (Inches(x - 0.01), Inches(y + card_h/2 - tab_h/2)),
            (Inches(x + tab_w), Inches(y + card_h/2)),
            (Inches(x - 0.01), Inches(y + card_h/2 + tab_h/2))
        ]
        tb_notch = s4.shapes.build_freeform(notch_pts[0][0], notch_pts[0][1])
        tb_notch.add_line_segments(notch_pts[1:], close=True)
        notch_shape = tb_notch.convert_to_shape()
        notch_shape.fill.solid()
        notch_shape.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        notch_shape.line.fill.background()
        
    tf_c = card.text_frame
    tf_c.word_wrap = True
    tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_c.margin_left = Inches(0.15 if i == 0 else 0.25)
    tf_c.margin_right = Inches(0.15)
    p1 = tf_c.paragraphs[0]
    p1.text = d_txt
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf_c.add_paragraph()
    p2.text = sub_d
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(10)
    p2.font.color.rgb = hex_to_rgb('#FFFFFF')
    p2.alignment = PP_ALIGN.CENTER
    
    c_stem = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + card_w/2 - 0.05), Inches(y + card_h), Inches(0.10), Inches(0.42))
    c_stem.fill.solid()
    c_stem.fill.fore_color.rgb = c_rgb
    c_stem.line.fill.background()
    
    tb_sub = s4.shapes.add_textbox(Inches(x - 0.10), Inches(y + card_h + 0.50), Inches(card_w + 0.20), Inches(1.10))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_s = tf_sub.paragraphs[0]
    p_s.text = sub_b.replace('\n', ' ')
    p_s.font.name = 'Segoe UI'
    p_s.font.size = Pt(11)
    p_s.font.color.rgb = hex_to_rgb('#222222')
    p_s.alignment = PP_ALIGN.CENTER
    
    b_bar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(6.15), Inches(card_w), Inches(0.14))
    b_bar.fill.solid()
    b_bar.fill.fore_color.rgb = c_rgb
    b_bar.line.fill.background()

# ==================== SLIDE 5: 6 HEXAGON PIN MARKERS ====================
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
    # Hexagon with sharp bottom V-tip
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
    
    # Inner White Circle
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
    
    # Curved Arch Connector
    arch_pts = [(Inches(x + 0.45), Inches(4.30)), (Inches(x + 0.90), Inches(4.55)), (Inches(x + 1.35), Inches(4.30))]
    fb_a = s5.shapes.build_freeform(arch_pts[0][0], arch_pts[0][1])
    fb_a.add_line_segments(arch_pts[1:], close=False)
    arch = fb_a.convert_to_shape()
    arch.fill.background()
    arch.line.color.rgb = hex_to_rgb(col)
    arch.line.width = Pt(2.0)
    
    # Outlined Bottom Card
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
print('Saved final 5 content slides to:', out_file)

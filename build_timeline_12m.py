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

def add_header(slide, title_text='Timeline Slide', align=PP_ALIGN.CENTER, x=0.8, y=0.55, w=11.733):
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

# ==================== SLIDE 1: 4 CALLOUT BOXES OVER BASELINE ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1)

# Baseline
line1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(5.30), Inches(11.733), Inches(0.02))
line1.fill.solid()
line1.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
line1.line.fill.background()

s1_data = [
    ('Begin', 'Start the project and set the main direction.', '#E53E3E', '💡', 1.00),
    ('Plan', 'Organize tasks and prepare the workflow for the year.', '#2B6CB0', '📋', 3.95),
    ('Execute', 'Move forward with key actions and major activities.', '#48BB78', '⚙️', 6.90),
    ('Conclude', 'Wrap up tasks and complete the overall goal.', '#DD6B20', '📑', 9.85),
]

for title, desc, col, icon, x in s1_data:
    c_rgb = hex_to_rgb(col)
    
    # Shadow
    shd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.03), Inches(1.83), Inches(2.45), Inches(3.00))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # White Card Box
    crd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.80), Inches(2.45), Inches(3.00))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    # Pointer notch down
    notch_pts = [(Inches(x + 1.05), Inches(4.80)), (Inches(x + 1.225), Inches(5.05)), (Inches(x + 1.40), Inches(4.80))]
    fb_n = s1.shapes.build_freeform(notch_pts[0][0], notch_pts[0][1])
    fb_n.add_line_segments(notch_pts[1:], close=True)
    n = fb_n.convert_to_shape()
    n.fill.solid()
    n.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    n.line.color.rgb = c_rgb
    n.line.width = Pt(1.5)
    
    # Icon
    tb_i = s1.shapes.add_textbox(Inches(x), Inches(2.10), Inches(2.45), Inches(0.50))
    tf_i = tb_i.text_frame
    p_i = tf_i.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(22)
    p_i.alignment = PP_ALIGN.CENTER
    
    # Text
    tb = s1.shapes.add_textbox(Inches(x + 0.15), Inches(2.80), Inches(2.15), Inches(1.80))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(6)
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.CENTER
    
    # Milestone Bead
    bead = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1.075), Inches(5.15), Inches(0.30), Inches(0.30))
    bead.fill.solid()
    bead.fill.fore_color.rgb = c_rgb
    bead.line.fill.background()
    
    # 20XX Label
    tb_yr = s1.shapes.add_textbox(Inches(x + 0.225), Inches(5.65), Inches(2.00), Inches(0.50))
    tf_yr = tb_yr.text_frame
    p_yr = tf_yr.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = 'Montserrat'
    p_yr.font.size = Pt(11)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#2D3748')
    p_yr.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 2: CONNECTED CHEVRON CHAIN ====================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s2)

s2_data = [
    ('Start', 'Begin the initiative and set the main direction.', '#E53E3E', '01', True, 0.80),
    ('Plan', 'Define yearly targets and organize the required tasks.', '#2B6CB0', '02', False, 3.15),
    ('Build', 'Develop core activities and push the project forward.', '#48BB78', '03', True, 5.50),
    ('Improve', 'Review outcomes and enhance areas needing attention.', '#DD6B20', '04', False, 7.85),
    ('Strengthen', 'Expand progress and reinforce successful actions.', '#00A3C4', '05', True, 10.20),
]

y_chv = 3.65
chv_w = 2.45
chv_h = 0.65

for title, desc, col, num, is_top, x in s2_data:
    c_rgb = hex_to_rgb(col)
    
    # Chevron Segment
    chv = s2.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y_chv), Inches(chv_w), Inches(chv_h))
    chv.fill.solid()
    chv.fill.fore_color.rgb = c_rgb
    chv.line.fill.background()
    p_yr = chv.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = 'Montserrat'
    p_yr.font.size = Pt(11)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_yr.alignment = PP_ALIGN.CENTER
    
    # Disc Position
    y_disc = 2.00 if is_top else 4.75
    
    # Dashed line connector
    ln = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.60), Inches(2.95 if is_top else 4.30), Inches(0.02), Inches(0.70))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    # Circular Disc
    shd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.08), Inches(y_disc + 0.03), Inches(1.10), Inches(1.10))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    disc = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.05), Inches(y_disc), Inches(1.10), Inches(1.10))
    disc.fill.solid()
    disc.fill.fore_color.rgb = c_rgb
    disc.line.fill.background()
    p_n = disc.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(16)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER
    
    # Text next to Disc
    tb = s2.shapes.add_textbox(Inches(x + 1.25), Inches(y_disc + 0.10), Inches(1.55), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.space_after = Pt(2)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')

# ==================== SLIDE 3: 5 OUTLINED CARD BOXES WITH TOP NUMBER DISCS ====================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3)

line3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(2.20), Inches(11.733), Inches(0.02))
line3.fill.solid()
line3.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
line3.line.fill.background()

s3_data = [
    ('Start Phase', 'Initial planning and goal setting.', '#E53E3E', '01', 0.80),
    ('Development', 'Design and prototype creation.', '#2B6CB0', '02', 3.25),
    ('Testing', 'Trial runs and adjustments.', '#48BB78', '03', 5.70),
    ('Launch', 'Official release or implementation.', '#DD6B20', '04', 8.15),
    ('Review', 'Performance evaluation and feedback collection.', '#E53E3E', '05', 10.60),
]

for title, desc, col, num, x in s3_data:
    c_rgb = hex_to_rgb(col)
    
    # Top Year
    tb_yr = s3.shapes.add_textbox(Inches(x + 0.10), Inches(1.50), Inches(1.80), Inches(0.40))
    tf_yr = tb_yr.text_frame
    p_yr = tf_yr.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = 'Montserrat'
    p_yr.font.size = Pt(11.5)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#2D3748')
    p_yr.alignment = PP_ALIGN.CENTER
    
    # Bead on Axis
    bead = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.85), Inches(2.06), Inches(0.28), Inches(0.28))
    bead.fill.solid()
    bead.fill.fore_color.rgb = c_rgb
    bead.line.fill.background()
    
    # Dashed line down
    ln = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.98), Inches(2.35), Inches(0.02), Inches(0.65))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    # Box Shadow
    shd = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.03), Inches(3.53), Inches(2.00), Inches(2.60))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    # Box
    box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(3.50), Inches(2.00), Inches(2.60))
    box.fill.solid()
    box.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    box.line.color.rgb = c_rgb
    box.line.width = Pt(1.5)
    
    # Top Number Disc
    disc = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.50), Inches(3.00), Inches(1.00), Inches(1.00))
    disc.fill.solid()
    disc.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc.line.color.rgb = c_rgb
    disc.line.width = Pt(1.5)
    p_n = disc.text_frame.paragraphs[0]
    p_n.text = num
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(13)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#2D3748')
    p_n.alignment = PP_ALIGN.CENTER
    
    # Box Text
    tb = s3.shapes.add_textbox(Inches(x + 0.15), Inches(4.30), Inches(1.70), Inches(1.60))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(6)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 4: 6 DOWNWARD BANNERS WITH ICONS ====================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4)

line4 = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(4.60), Inches(11.733), Inches(0.02))
line4.fill.solid()
line4.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
line4.line.fill.background()

s4_data = [
    ('Initiate', 'Start the project and define the core purpose.', '#E53E3E', '⚙️', 0.80),
    ('Outline', 'Set yearly goals and arrange the primary action steps.', '#2B6CB0', '📝', 2.80),
    ('Develop', 'Work on major tasks and move plans into execution.', '#48BB78', '🧠', 4.80),
    ('Refine', 'Review progress and improve areas that need adjustment.', '#DD6B20', '🎛️', 6.80),
    ('Advance', 'Strengthen results and expand on successful actions.', '#E53E3E', '🏆', 8.80),
    ('Complete', 'Finalize all work and deliver the full project outcome.', '#00A3C4', '✅', 10.80),
]

for title, desc, col, icon, x in s4_data:
    c_rgb = hex_to_rgb(col)
    
    # Downward Banner
    ban_pts = [
        (Inches(x), Inches(2.20)),
        (Inches(x + 1.70), Inches(2.20)),
        (Inches(x + 1.70), Inches(3.60)),
        (Inches(x + 0.85), Inches(4.00)),
        (Inches(x), Inches(3.60))
    ]
    fb_b = s4.shapes.build_freeform(ban_pts[0][0], ban_pts[0][1])
    fb_b.add_line_segments(ban_pts[1:], close=True)
    ban = fb_b.convert_to_shape()
    ban.fill.solid()
    ban.fill.fore_color.rgb = c_rgb
    ban.line.fill.background()
    
    # Banner 20XX Text
    tb_yr = s4.shapes.add_textbox(Inches(x), Inches(3.05), Inches(1.70), Inches(0.50))
    tf_yr = tb_yr.text_frame
    p_yr = tf_yr.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = 'Montserrat'
    p_yr.font.size = Pt(11)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_yr.alignment = PP_ALIGN.CENTER
    
    # White Circular Icon Disc
    disc = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.35), Inches(1.65), Inches(1.00), Inches(1.00))
    disc.fill.solid()
    disc.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc.line.color.rgb = c_rgb
    disc.line.width = Pt(1.5)
    p_i = disc.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(16)
    p_i.alignment = PP_ALIGN.CENTER
    
    # Bead on Axis
    bead = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.725), Inches(4.475), Inches(0.25), Inches(0.25))
    bead.fill.solid()
    bead.fill.fore_color.rgb = c_rgb
    bead.line.fill.background()
    
    # Text below Axis
    tb = s4.shapes.add_textbox(Inches(x - 0.15), Inches(5.15), Inches(2.00), Inches(1.60))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(4)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 5: 5 CIRCULAR BALLOONS OVER BASELINE ====================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5)

line5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.80), Inches(5.00), Inches(11.733), Inches(0.02))
line5.fill.solid()
line5.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
line5.line.fill.background()

s5_data = [
    ('Planning', 'Define goals and strategy.', '#E53E3E', 1.00),
    ('Design', 'Create and develop.', '#2B6CB0', 3.45),
    ('Testing', 'Conduct trials and refine.', '#48BB78', 5.90),
    ('Launch', 'Release and implement.', '#DD6B20', 8.35),
    ('Evaluate', 'Assess performance and feedback.', '#E53E3E', 10.80),
]

for title, desc, col, x in s5_data:
    c_rgb = hex_to_rgb(col)
    
    # Dashed line
    ln = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.69), Inches(3.30), Inches(0.02), Inches(1.60))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    # Circular Balloon
    shd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.03), Inches(1.93), Inches(1.40), Inches(1.40))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    bal = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(1.90), Inches(1.40), Inches(1.40))
    bal.fill.solid()
    bal.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bal.line.color.rgb = c_rgb
    bal.line.width = Pt(1.5)
    p_yr = bal.text_frame.paragraphs[0]
    p_yr.text = '20XX'
    p_yr.font.name = 'Montserrat'
    p_yr.font.size = Pt(13)
    p_yr.font.bold = True
    p_yr.font.color.rgb = hex_to_rgb('#2D3748')
    p_yr.alignment = PP_ALIGN.CENTER
    
    # Bead
    bead = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.56), Inches(4.86), Inches(0.28), Inches(0.28))
    bead.fill.solid()
    bead.fill.fore_color.rgb = c_rgb
    bead.line.fill.background()
    
    # Text below
    tb = s5.shapes.add_textbox(Inches(x - 0.20), Inches(5.40), Inches(1.80), Inches(1.40))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(4)
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.CENTER

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Timeline_Template_Powerpoint\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 8 saved to:', out_file)

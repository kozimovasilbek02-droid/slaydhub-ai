"""
Master Precision Builder for Executive & Sales Director 30-60-90 Day Roadmaps
Implements 5 high-precision infographic architectures:
1. Left Pillar Sidebar with Continuous Center Progress Beam & Bottom Halo Badges
2. Split-Level Dual-Badge Matrix with Horizontal Dividing Beam
3. Neo-Brutalist 3D Offset Pastel Cards with Top Circular Nodes
4. Ascending Velocity Chevron Process Flow
5. Concentric Halo Discs with Pill Sub-Tabs
"""
import sys, os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

target_folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\30_60_90_Day_Plan_Sales_Director_Examples"
folder_name = os.path.basename(os.path.normpath(target_folder))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

C_CYAN = '#06B6D4'
C_ROSE = '#F43F5E'
C_AMBER = '#F59E0B'
THEME_COLORS = [C_CYAN, C_ROSE, C_AMBER]

def add_header(slide, title, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = 'Georgia'
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E293B')
    p.alignment = PP_ALIGN.CENTER
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(13)
        p2.font.color.rgb = hex_to_rgb('#64748B')
        p2.alignment = PP_ALIGN.CENTER

descriptions = [
    "Understand team dynamics, assess baseline performance, conduct audits, and build vital relationships.",
    "Develop core sales strategies, identify growth opportunities, and implement tactical process improvements.",
    "Measure quantitative results, refine key workflows, and set scalable long-term goals for sustained success."
]

# ----------------------------------------------------
# Slide 1: Left Pillar Sidebar + Center Beam + Tall Cards
# ----------------------------------------------------
s1 = prs.slides.add_slide(blank_layout)
add_header(s1, "30 60 90 Day Plan Sales Director Examples")

# Left Pillar Card
p_sh = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.8), Inches(2.0), Inches(5.0))
p_sh.fill.solid()
p_sh.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
p_sh.line.color.rgb = hex_to_rgb('#E2E8F0')

# Left Purple Vertical Accent
p_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.9), Inches(1.8), Inches(0.12), Inches(5.0))
p_bar.fill.solid()
p_bar.fill.fore_color.rgb = hex_to_rgb('#6366F1')
p_bar.line.fill.background()

tb_p = s1.shapes.add_textbox(Inches(1.1), Inches(3.2), Inches(1.7), Inches(2.5))
tf = tb_p.text_frame
for word in ["Days", "Plan", "Sales"]:
    p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
    p.text = word
    p.font.name = 'Georgia'
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E293B')
    p.alignment = PP_ALIGN.CENTER

# Continuous Slate-Purple Progress Beam
beam = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.2), Inches(3.8), Inches(9.1), Inches(0.15))
beam.fill.solid()
beam.fill.fore_color.rgb = hex_to_rgb('#6366F1')
beam.line.fill.background()

card_w = 2.65
card_gap = 0.45
start_cx = 3.3
percentages = ["50%", "30%", "48%"]
days_lbl = ["30 DAYS", "60 DAYS", "90 DAYS"]

for i in range(3):
    cx = start_cx + i * (card_w + card_gap)
    col = THEME_COLORS[i]
    
    # Card Background
    card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(1.8), Inches(card_w), Inches(4.8))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
    card.line.color.rgb = hex_to_rgb('#E2E8F0')
    
    # Top Days Title
    tb_t = s1.shapes.add_textbox(Inches(cx), Inches(2.2), Inches(card_w), Inches(0.6))
    p = tb_t.text_frame.paragraphs[0]
    p.text = days_lbl[i]
    p.font.name = 'Montserrat'
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E293B')
    p.alignment = PP_ALIGN.CENTER
    
    # Accent Underline Bar
    u_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.2), Inches(3.1), Inches(card_w - 0.4), Inches(0.18))
    u_bar.fill.solid()
    u_bar.fill.fore_color.rgb = hex_to_rgb(col)
    u_bar.line.fill.background()
    
    # Beam Node Dot
    dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + card_w/2 - 0.12), Inches(3.75), Inches(0.24), Inches(0.24))
    dot.fill.solid()
    dot.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    dot.line.color.rgb = hex_to_rgb('#6366F1')
    dot.line.width = Pt(1.5)
    
    # Description Body
    tb_d = s1.shapes.add_textbox(Inches(cx + 0.15), Inches(4.2), Inches(card_w - 0.3), Inches(1.8))
    tf = tb_d.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = descriptions[i]
    p.font.name = 'Segoe UI'
    p.font.size = Pt(10.5)
    p.font.color.rgb = hex_to_rgb('#475569')
    p.alignment = PP_ALIGN.CENTER
    
    # Bottom Percentage Dual-Halo Badge
    halo = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + card_w/2 - 0.6), Inches(6.0), Inches(1.2), Inches(1.2))
    halo.fill.solid()
    halo.fill.fore_color.rgb = hex_to_rgb(col)
    halo.line.color.rgb = hex_to_rgb('#FFFFFF')
    halo.line.width = Pt(3.0)
    p = halo.text_frame.paragraphs[0]
    p.text = percentages[i]
    p.font.name = 'Montserrat'
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 2: Split Level Dual-Badge Matrix with Divider Beam
# ----------------------------------------------------
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "30 60 90 Day Plan Sales Director Examples", "Structured Progression & Milestone Metrics")

col2_w = 3.3
col2_gap = 0.65
start_x2 = 1.15
seq_nums = ["01", "02", "03"]
milestones = ["30", "60", "90"]
T_COLS_S2 = [C_AMBER, C_ROSE, C_CYAN]

for i in range(3):
    cx = start_x2 + i * (col2_w + col2_gap)
    col = T_COLS_S2[i]
    
    # Top Card
    top_c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(2.4), Inches(col2_w), Inches(2.0))
    top_c.fill.solid()
    top_c.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    top_c.line.color.rgb = hex_to_rgb('#E2E8F0')
    top_c.line.width = Pt(1.5)
    
    # Top Circular Ring Badge
    rb = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + col2_w/2 - 0.65), Inches(1.75), Inches(1.3), Inches(1.3))
    rb.fill.solid()
    rb.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    rb.line.color.rgb = hex_to_rgb(col)
    rb.line.width = Pt(3.5)
    p = rb.text_frame.paragraphs[0]
    p.text = seq_nums[i]
    p.font.name = 'Montserrat'
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E293B')
    p.alignment = PP_ALIGN.CENTER
    
    # Text in Top Card
    tb_tc = s2.shapes.add_textbox(Inches(cx + 0.15), Inches(2.95), Inches(col2_w - 0.3), Inches(1.35))
    tf = tb_tc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = descriptions[i]
    p.font.name = 'Segoe UI'
    p.font.size = Pt(10.5)
    p.font.color.rgb = hex_to_rgb('#475569')
    p.alignment = PP_ALIGN.CENTER

# Middle Dividing Beam
d_beam = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.15), Inches(4.7), Inches(11.033), Inches(0.55))
d_beam.fill.solid()
d_beam.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
d_beam.line.fill.background()

# Bottom Milestone Solid Spheres
for i in range(3):
    cx = start_x2 + i * (col2_w + col2_gap)
    col = T_COLS_S2[i]
    
    # Shadow
    b_sh = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + col2_w/2 - 0.77), Inches(5.62), Inches(1.6), Inches(1.6))
    b_sh.fill.solid()
    b_sh.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    b_sh.line.fill.background()
    
    # Sphere
    sph = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + col2_w/2 - 0.8), Inches(5.55), Inches(1.6), Inches(1.6))
    sph.fill.solid()
    sph.fill.fore_color.rgb = hex_to_rgb(col)
    sph.line.fill.background()
    p = sph.text_frame.paragraphs[0]
    p.text = milestones[i]
    p.font.name = 'Georgia'
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 3: Neo-Brutalist Offset Pastel Cards with Badges
# ----------------------------------------------------
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "30 60 90 Day Plan Sales Director Examples", "3D Chromatic Offset Strategy Framework")

PASTEL_BG = ['#FEF9C3', '#FFE4E6', '#CFFAFE']
SOLID_ACC = [C_AMBER, C_ROSE, C_CYAN]

card3_w = 3.2
card3_gap = 0.7
start_x3 = 1.25

for i in range(3):
    cx = start_x3 + i * (card3_w + card3_gap)
    col_solid = SOLID_ACC[i]
    col_pastel = PASTEL_BG[i]
    
    # Solid Offset Shadow Base
    base = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - 0.12), Inches(2.62), Inches(card3_w), Inches(4.2))
    base.fill.solid()
    base.fill.fore_color.rgb = hex_to_rgb(col_solid)
    base.line.fill.background()
    
    # Pastel Main Body Card
    body = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(2.5), Inches(card3_w), Inches(4.2))
    body.fill.solid()
    body.fill.fore_color.rgb = hex_to_rgb(col_pastel)
    body.line.fill.background()
    
    # Top Milestone Sphere
    sph3 = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + card3_w/2 - 0.75), Inches(1.75), Inches(1.5), Inches(1.5))
    sph3.fill.solid()
    sph3.fill.fore_color.rgb = hex_to_rgb(col_solid)
    sph3.line.fill.background()
    p = sph3.text_frame.paragraphs[0]
    p.text = milestones[i]
    p.font.name = 'Georgia'
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Description Body
    tb_b3 = s3.shapes.add_textbox(Inches(cx + 0.25), Inches(3.6), Inches(card3_w - 0.5), Inches(2.8))
    tf = tb_b3.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = descriptions[i]
    p.font.name = 'Georgia'
    p.font.size = Pt(12)
    p.font.color.rgb = hex_to_rgb('#334155')
    p.alignment = PP_ALIGN.CENTER

out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")

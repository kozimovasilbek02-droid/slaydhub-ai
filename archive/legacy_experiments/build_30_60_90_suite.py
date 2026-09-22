"""
Master Precision Builder for 30-60-90 Day Action Plans & Strategy Roadmaps
Implements 6 distinct infographic architectures:
1. Stepped Ascending Arrow Ramp with Drop Connectors
2. 3 Concentric Halo Discs with Pill Tabs
3. 3 Outlined Container Cards with Top Disc Badges
4. Stacked Horizontal Capsule Ribbons with Left Discs
5. 3 Slanted Parallelograms with Bullet Legends
6. 3 Bold Horizontal Chevron Process Arrows
"""
import sys, os, math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

target_folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\30_60_90_Action_Plan_Template_For_Sales_Presentation_54896"
folder_name = os.path.basename(os.path.normpath(target_folder))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Theme Colors (Yellow, Crimson/Pink, Teal/Cyan)
C_YEL = '#F59E0B'
C_RED = '#E11D48'
C_TEAL = '#0D9488'
THEME_COLORS = [C_YEL, C_RED, C_TEAL]

def add_header(slide, title, subtitle="Strategic Action Plan"):
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.45), Inches(11.333), Inches(1.1))
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

# ----------------------------------------------------
# Slide 1: 3 Concentric Halo Discs & Pill Tabs
# ----------------------------------------------------
s1 = prs.slides.add_slide(blank_layout)
add_header(s1, "30 60 90 Day Marketing Plan", "Phase Milestones & Strategic Objectives")

labels = ["30 Days", "60 Days", "90 Days"]
sub_labels = ["First 30 days", "First 60 days", "First 90 days"]
descriptions = [
    "A comprehensive plan for the first 30 days of a new job, territory, or leadership role.",
    "Divided into 30-day goals for short-term accomplishments and quick organizational wins.",
    "Builds long-term momentum towards achieving annual sales targets and operational KPIs."
]

col_w = 3.4
gap_x = 0.65
start_x = 1.05

for i in range(3):
    cx = start_x + i * (col_w + gap_x)
    col = THEME_COLORS[i]
    
    # Concentric Rings (Outer Halo + Inner Ring + White Hub)
    halo = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 0.9), Inches(1.9), Inches(1.6), Inches(1.6))
    halo.fill.solid()
    halo.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
    halo.line.color.rgb = hex_to_rgb(col)
    halo.line.width = Pt(3.5)
    
    hub = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 1.05), Inches(2.05), Inches(1.3), Inches(1.3))
    hub.fill.solid()
    hub.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    hub.line.color.rgb = hex_to_rgb(col)
    hub.line.width = Pt(1.5)
    p = hub.text_frame.paragraphs[0]
    p.text = labels[i].split()[0]
    p.font.name = 'Montserrat'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(col)
    p.alignment = PP_ALIGN.CENTER
    
    # Pill Tab
    pill = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.4), Inches(3.8), Inches(col_w - 0.8), Inches(0.6))
    pill.fill.solid()
    pill.fill.fore_color.rgb = hex_to_rgb(col)
    pill.line.fill.background()
    p = pill.text_frame.paragraphs[0]
    p.text = sub_labels[i]
    p.font.name = 'Montserrat'
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Description Body
    tb_desc = s1.shapes.add_textbox(Inches(cx), Inches(4.6), Inches(col_w), Inches(2.0))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = descriptions[i]
    p.font.name = 'Segoe UI'
    p.font.size = Pt(11)
    p.font.color.rgb = hex_to_rgb('#64748B')
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 2: 3 Outlined Container Cards with Top Disc Badges
# ----------------------------------------------------
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "30 60 90 Day Implementation Matrix", "Structured Execution Framework")

for i in range(3):
    cx = start_x + i * (col_w + gap_x)
    col = THEME_COLORS[i]
    
    # Card Body with Shadow
    c_sh = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.05), Inches(2.45), Inches(col_w), Inches(4.3))
    c_sh.fill.solid()
    c_sh.fill.fore_color.rgb = hex_to_rgb('#F1F5F9')
    c_sh.line.fill.background()
    
    card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(2.4), Inches(col_w), Inches(4.3))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.color.rgb = hex_to_rgb(col)
    card.line.width = Pt(2.5)
    
    # Top Disc Badge (Overlapping)
    badge = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + col_w/2 - 0.65), Inches(1.9), Inches(1.3), Inches(1.3))
    badge.fill.solid()
    badge.fill.fore_color.rgb = hex_to_rgb(col)
    badge.line.color.rgb = hex_to_rgb('#FFFFFF')
    badge.line.width = Pt(3.0)
    p = badge.text_frame.paragraphs[0]
    p.text = labels[i].split()[0]
    p.font.name = 'Montserrat'
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Title & Subtitle inside Card
    tb = s2.shapes.add_textbox(Inches(cx + 0.2), Inches(3.4), Inches(col_w - 0.4), Inches(0.6))
    p = tb.text_frame.paragraphs[0]
    p.text = sub_labels[i]
    p.font.name = 'Montserrat'
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(col)
    p.alignment = PP_ALIGN.CENTER
    
    tb_b = s2.shapes.add_textbox(Inches(cx + 0.3), Inches(4.1), Inches(col_w - 0.6), Inches(2.3))
    tf = tb_b.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = descriptions[i]
    p.font.name = 'Segoe UI'
    p.font.size = Pt(11)
    p.font.color.rgb = hex_to_rgb('#475569')
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 3: Stacked Horizontal Capsule Ribbons
# ----------------------------------------------------
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "30 60 90 Day Strategic Roadmap", "Linear Timeline Progression")

for i in range(3):
    cy = 2.0 + i * 1.65
    col = THEME_COLORS[i]
    
    # Capsule Ribbon
    rib = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(cy), Inches(9.2), Inches(1.3))
    rib.fill.solid()
    rib.fill.fore_color.rgb = hex_to_rgb(col)
    rib.line.fill.background()
    
    # Left 360° Circular Badge
    disc = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.6), Inches(cy - 0.1), Inches(1.5), Inches(1.5))
    disc.fill.solid()
    disc.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc.line.color.rgb = hex_to_rgb(col)
    disc.line.width = Pt(3.0)
    p = disc.text_frame.paragraphs[0]
    p.text = labels[i].split()[0]
    p.font.name = 'Montserrat'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(col)
    p.alignment = PP_ALIGN.CENTER
    
    # Text inside Capsule
    tb_txt = s3.shapes.add_textbox(Inches(3.3), Inches(cy + 0.15), Inches(7.8), Inches(1.0))
    tf = tb_txt.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = sub_labels[i]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    
    p2 = tf.add_paragraph()
    p2.text = descriptions[i]
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(10.5)
    p2.font.color.rgb = hex_to_rgb('#FFFBEB' if i == 0 else '#FFE4E6' if i == 1 else '#CCFBF1')

# ----------------------------------------------------
# Slide 4: 3 Slanted Parallelograms with Bullet Legends
# ----------------------------------------------------
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "30 60 90 Day Execution Milestones", "Sequential Growth Blocks")

p_width = 3.2
p_gap = 0.5
start_px = 1.3

for i in range(3):
    px = start_px + i * (p_width + p_gap)
    col = THEME_COLORS[i]
    
    # Parallelogram Block
    para = s4.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(px), Inches(2.2), Inches(p_width), Inches(1.8))
    para.fill.solid()
    para.fill.fore_color.rgb = hex_to_rgb(col)
    para.line.fill.background()
    p = para.text_frame.paragraphs[0]
    p.text = labels[i].split()[0]
    p.font.name = 'Montserrat'
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Matching Legend Row below
    leg_tab = s4.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(2.0), Inches(4.5 + i * 0.8), Inches(0.8), Inches(0.4))
    leg_tab.fill.solid()
    leg_tab.fill.fore_color.rgb = hex_to_rgb(col)
    leg_tab.line.fill.background()
    
    tb_l = s4.shapes.add_textbox(Inches(3.0), Inches(4.45 + i * 0.8), Inches(8.5), Inches(0.5))
    tf = tb_l.text_frame
    p = tf.paragraphs[0]
    p.text = f"{sub_labels[i]}: {descriptions[i]}"
    p.font.name = 'Segoe UI'
    p.font.size = Pt(11)
    p.font.color.rgb = hex_to_rgb('#334155')

# ----------------------------------------------------
# Slide 5: 3 Bold Horizontal Chevron Process Arrows
# ----------------------------------------------------
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "30 60 90 Day Action Process", "Fast-Track Velocity Arrows")

arr_w = 3.6
arr_gap = 0.35
start_ax = 1.05

for i in range(3):
    ax = start_ax + i * (arr_w + arr_gap)
    col = THEME_COLORS[i]
    
    # Chevron Arrow
    chv = s5.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(ax), Inches(2.3), Inches(arr_w), Inches(2.0))
    chv.fill.solid()
    chv.fill.fore_color.rgb = hex_to_rgb(col)
    chv.line.fill.background()
    p = chv.text_frame.paragraphs[0]
    p.text = labels[i].split()[0]
    p.font.name = 'Montserrat'
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Bottom Description Block
    tb_d = s5.shapes.add_textbox(Inches(ax), Inches(4.6), Inches(arr_w), Inches(2.2))
    tf = tb_d.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = sub_labels[i]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(col)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = descriptions[i]
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(11)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.CENTER

out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")

"""
Master Precision Builder for Quarterly & Yearly Calendars
Supports:
- 2023_Quarterly_Powerpoint_Calendar
- 2022_Year_Planner_Free_Download
- 2023_Planning_Calendar_Powerpoint_Template
- 2023_Yearly_Calendar_For_Powerpoint
- 2023_Yearly_Powerpoint_Calendar_Slide
"""
import sys, os, glob
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

target_folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\2023_Quarterly_Powerpoint_Calendar"
folder_name = os.path.basename(os.path.normpath(target_folder))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def add_header(slide, title, subtitle=None, color_hex='#1E293B'):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = 'Georgia'
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(color_hex)
    p.alignment = PP_ALIGN.CENTER
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.name = 'Montserrat'
        p2.font.size = Pt(14)
        p2.font.color.rgb = hex_to_rgb('#64748B')
        p2.alignment = PP_ALIGN.CENTER

days_header = ['S', 'M', 'T', 'W', 'T', 'F', 'S']

# ----------------------------------------------------
# Slide 1: Cover Slide
# ----------------------------------------------------
s1 = prs.slides.add_slide(blank_layout)

# Top Right Angular Header
bg_tri = s1.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(4.5), Inches(0), Inches(8.833), Inches(3.8))
bg_tri.fill.solid()
bg_tri.fill.fore_color.rgb = hex_to_rgb('#991B1B')
bg_tri.line.fill.background()
bg_tri.rotation = 180

tb_title = s1.shapes.add_textbox(Inches(1.0), Inches(3.2), Inches(10.0), Inches(2.5))
tf = tb_title.text_frame
p1 = tf.paragraphs[0]
p1.text = "2023"
p1.font.name = 'Georgia'
p1.font.size = Pt(64)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#B91C1C')

p2 = tf.add_paragraph()
p2.text = "Quarterly Calendar"
p2.font.name = 'Georgia'
p2.font.size = Pt(36)
p2.font.bold = True
p2.font.color.rgb = hex_to_rgb('#991B1B')

# ----------------------------------------------------
# Slide 2: Q1 (Jan, Feb, Mar) - Blue Circular Cards
# ----------------------------------------------------
s2 = prs.slides.add_slide(blank_layout)

# Left Accent Bar & Title
bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.2), Inches(0.5), Inches(0.15), Inches(1.0))
bar.fill.solid()
bar.fill.fore_color.rgb = hex_to_rgb('#0066CC')
bar.line.fill.background()

tb_s2 = s2.shapes.add_textbox(Inches(2.5), Inches(0.4), Inches(8.5), Inches(1.1))
p = tb_s2.text_frame.paragraphs[0]
p.text = "2023 Quarterly Calendar - Q1"
p.font.name = 'Georgia'
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#0066CC')

q1_months = ["JANUARY", "FEBRUARY", "MARCH"]
card_w, card_h = 3.5, 4.6
start_x = 1.0
gap_x = 0.65

for idx, m_name in enumerate(q1_months):
    cx = start_x + idx * (card_w + gap_x)
    
    # Top Tab
    tab = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.4), Inches(2.0), Inches(card_w - 0.8), Inches(0.7))
    tab.fill.solid()
    tab.fill.fore_color.rgb = hex_to_rgb('#0066CC')
    tab.line.fill.background()
    p = tab.text_frame.paragraphs[0]
    p.text = m_name
    p.font.name = 'Montserrat'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    # Main Circle Shadow & Body
    disc_sh = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 0.05), Inches(2.45), Inches(card_w), Inches(card_w))
    disc_sh.fill.solid()
    disc_sh.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    disc_sh.line.fill.background()
    
    disc = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(2.4), Inches(card_w), Inches(card_w))
    disc.fill.solid()
    disc.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    disc.line.color.rgb = hex_to_rgb('#E2E8F0')
    disc.line.width = Pt(1.0)
    
    # Days Row
    day_w = (card_w - 0.6) / 7.0
    for d_idx, day_ch in enumerate(days_header):
        tb_d = s2.shapes.add_textbox(Inches(cx + 0.3 + d_idx * day_w), Inches(2.8), Inches(day_w), Inches(0.4))
        p = tb_d.text_frame.paragraphs[0]
        p.text = day_ch
        p.font.name = 'Montserrat'
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb('#1E293B')
        p.alignment = PP_ALIGN.CENTER
        
    # Date Grid
    for r in range(5):
        for c in range(7):
            num = r * 7 + c + 1
            if num > 31:
                continue
            tb_n = s2.shapes.add_textbox(Inches(cx + 0.3 + c * day_w), Inches(3.2 + r * 0.45), Inches(day_w), Inches(0.4))
            p = tb_n.text_frame.paragraphs[0]
            p.text = f"{num}"
            p.font.name = 'Calibri'
            p.font.size = Pt(10)
            p.font.color.rgb = hex_to_rgb('#475569')
            p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 3: Q2 (Apr, May, Jun) - Gold Banner Cards
# ----------------------------------------------------
s3 = prs.slides.add_slide(blank_layout)

# Top Gold Header Banner
hdr_b = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(2.2))
hdr_b.fill.solid()
hdr_b.fill.fore_color.rgb = hex_to_rgb('#D48806')
hdr_b.line.fill.background()

tb_s3 = s3.shapes.add_textbox(Inches(1.0), Inches(0.3), Inches(11.333), Inches(1.5))
tf = tb_s3.text_frame
p1 = tf.paragraphs[0]
p1.text = "2023"
p1.font.name = 'Georgia'
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#FFFFFF')
p1.alignment = PP_ALIGN.CENTER

p2 = tf.add_paragraph()
p2.text = "Quarterly Calendar (Q2)"
p2.font.name = 'Georgia'
p2.font.size = Pt(22)
p2.font.bold = True
p2.font.color.rgb = hex_to_rgb('#FFFBEB')
p2.alignment = PP_ALIGN.CENTER

q2_months = ["APRIL", "MAY", "JUNE"]
for idx, m_name in enumerate(q2_months):
    cx = start_x + idx * (card_w + gap_x)
    
    # White Card Body with Shadow
    c_sh = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.05), Inches(2.75), Inches(card_w), Inches(4.3))
    c_sh.fill.solid()
    c_sh.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    c_sh.line.fill.background()
    
    c_body = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(2.7), Inches(card_w), Inches(4.3))
    c_body.fill.solid()
    c_body.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    c_body.line.color.rgb = hex_to_rgb('#E2E8F0')
    
    # Month Title
    tb_m = s3.shapes.add_textbox(Inches(cx), Inches(2.9), Inches(card_w), Inches(0.5))
    p = tb_m.text_frame.paragraphs[0]
    p.text = m_name
    p.font.name = 'Montserrat'
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#D48806')
    p.alignment = PP_ALIGN.CENTER
    
    # Gold Days Bar
    d_bar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.2), Inches(3.45), Inches(card_w - 0.4), Inches(0.45))
    d_bar.fill.solid()
    d_bar.fill.fore_color.rgb = hex_to_rgb('#FDEBD0')
    d_bar.line.fill.background()
    
    day_w = (card_w - 0.4) / 7.0
    for d_idx, day_ch in enumerate(days_header):
        tb_d = s3.shapes.add_textbox(Inches(cx + 0.2 + d_idx * day_w), Inches(3.45), Inches(day_w), Inches(0.45))
        p = tb_d.text_frame.paragraphs[0]
        p.text = day_ch
        p.font.name = 'Montserrat'
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb('#78350F')
        p.alignment = PP_ALIGN.CENTER
        
    for r in range(5):
        for c in range(7):
            num = r * 7 + c + 1
            if num > 30:
                continue
            tb_n = s3.shapes.add_textbox(Inches(cx + 0.2 + c * day_w), Inches(4.0 + r * 0.5), Inches(day_w), Inches(0.45))
            p = tb_n.text_frame.paragraphs[0]
            p.text = f"{num}"
            p.font.name = 'Calibri'
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb('#1E293B')
            p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 4: Q3 (Jul, Aug, Sep) - Floating Blue Cards
# ----------------------------------------------------
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Quarterly Calendar", "Q3: July - September 2023", '#1890FF')

q3_months = ["July - 2023", "August - 2023", "September - 2023"]
for idx, m_name in enumerate(q3_months):
    cx = start_x + idx * (card_w + gap_x)
    
    c_sh = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.05), Inches(2.15), Inches(card_w), Inches(4.8))
    c_sh.fill.solid()
    c_sh.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    c_sh.line.fill.background()
    
    c_body = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(2.1), Inches(card_w), Inches(4.8))
    c_body.fill.solid()
    c_body.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    c_body.line.color.rgb = hex_to_rgb('#CBD5E1')
    
    # Top Blue Header Tab
    h_tab = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.2), Inches(2.3), Inches(card_w - 0.4), Inches(0.65))
    h_tab.fill.solid()
    h_tab.fill.fore_color.rgb = hex_to_rgb('#1890FF')
    h_tab.line.fill.background()
    p = h_tab.text_frame.paragraphs[0]
    p.text = m_name
    p.font.name = 'Montserrat'
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER
    
    day_w = (card_w - 0.4) / 7.0
    for d_idx, day_ch in enumerate(days_header):
        tb_d = s4.shapes.add_textbox(Inches(cx + 0.2 + d_idx * day_w), Inches(3.1), Inches(day_w), Inches(0.45))
        p = tb_d.text_frame.paragraphs[0]
        p.text = day_ch
        p.font.name = 'Montserrat'
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb('#1E293B')
        p.alignment = PP_ALIGN.CENTER
        
    for r in range(5):
        for c in range(7):
            num = r * 7 + c + 1
            if num > 31:
                continue
            tb_n = s4.shapes.add_textbox(Inches(cx + 0.2 + c * day_w), Inches(3.6 + r * 0.55), Inches(day_w), Inches(0.45))
            p = tb_n.text_frame.paragraphs[0]
            p.text = f"{num}"
            p.font.name = 'Calibri'
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb('#334155')
            p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 5: Q4 (Oct, Nov, Dec) - Pinned Note Cards
# ----------------------------------------------------
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Quarterly Calendar", "Q4: October - December 2023", '#4C1D95')

q4_months = ["OCTOBER", "NOVEMBER", "DECEMBER"]
for idx, m_name in enumerate(q4_months):
    cx = start_x + idx * (card_w + gap_x)
    
    # Realistic Shadow
    sh = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.08), Inches(2.28), Inches(card_w), Inches(4.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    sh.line.fill.background()
    
    # Paper Note Body
    note = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(2.2), Inches(card_w), Inches(4.6))
    note.fill.solid()
    note.fill.fore_color.rgb = hex_to_rgb('#FAF5FF')
    note.line.color.rgb = hex_to_rgb('#E9D5FF')
    note.line.width = Pt(1.0)
    
    # Top Pin Badge
    pin = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + card_w/2 - 0.2), Inches(2.0), Inches(0.4), Inches(0.4))
    pin.fill.solid()
    pin.fill.fore_color.rgb = hex_to_rgb('#5B21B6')
    pin.line.fill.background()
    
    # Month Title
    tb_m = s5.shapes.add_textbox(Inches(cx), Inches(2.6), Inches(card_w), Inches(0.5))
    p = tb_m.text_frame.paragraphs[0]
    p.text = m_name
    p.font.name = 'Georgia'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#5B21B6')
    p.alignment = PP_ALIGN.CENTER
    
    day_w = (card_w - 0.4) / 7.0
    for d_idx, day_ch in enumerate(days_header):
        tb_d = s5.shapes.add_textbox(Inches(cx + 0.2 + d_idx * day_w), Inches(3.15), Inches(day_w), Inches(0.45))
        p = tb_d.text_frame.paragraphs[0]
        p.text = day_ch
        p.font.name = 'Montserrat'
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb('#1E293B')
        p.alignment = PP_ALIGN.CENTER
        
    for r in range(5):
        for c in range(7):
            num = r * 7 + c + 1
            if num > 31:
                continue
            tb_n = s5.shapes.add_textbox(Inches(cx + 0.2 + c * day_w), Inches(3.65 + r * 0.55), Inches(day_w), Inches(0.45))
            p = tb_n.text_frame.paragraphs[0]
            p.text = f"{num}"
            p.font.name = 'Calibri'
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb('#1E293B')
            p.alignment = PP_ALIGN.CENTER

out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")

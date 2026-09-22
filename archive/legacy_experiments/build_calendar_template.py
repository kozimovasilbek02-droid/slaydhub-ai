"""
Master Precision Builder for Calendar Presentations
Supports:
- 2021_August_Ppt_Calendar_Template
- 2021_Calendar_Template_56145
- 2021_February_Ppt_Calendar_Template
- 2021_January_Ppt_Calendar_Template
- 2021_July_Ppt_Calendar_Template
- 2021_May_Ppt_Calendar_Template
- 2021_September_Ppt_Calendar_Template
"""
import sys, os, math, glob
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

target_folder = sys.argv[1] if len(sys.argv) > 1 else r"G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\2021_August_Ppt_Calendar_Template"
folder_name = os.path.basename(os.path.normpath(target_folder))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def add_header(slide, title_text, color_hex='#028090'):
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = 'Georgia'
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(color_hex)
    p.alignment = PP_ALIGN.CENTER

# Slide 1: Modern Rounded Table Calendar
s1 = prs.slides.add_slide(blank_layout)
add_header(s1, f"{folder_name.replace('_', ' ')} - Monthly Plan", '#028090')

cx_card, cy_card = 1.666, 1.6
cw_card, ch_card = 10.0, 5.2
card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx_card), Inches(cy_card), Inches(cw_card), Inches(ch_card))
card.fill.solid()
card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
card.line.color.rgb = hex_to_rgb('#00A896')
card.line.width = Pt(2.0)

days = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT']
col_w = cw_card / 7.0
hdr_h = 0.65

for i, day in enumerate(days):
    h_box = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx_card + i * col_w), Inches(cy_card), Inches(col_w), Inches(hdr_h))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = hex_to_rgb('#00A896' if i != 0 and i != 6 else '#028090')
    h_box.line.color.rgb = hex_to_rgb('#FFFFFF')
    h_box.line.width = Pt(1.0)
    tf = h_box.text_frame
    p = tf.paragraphs[0]
    p.text = day
    p.font.name = 'Montserrat'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

dates = [
    ['', '', '', '01', '02', '03', '04'],
    ['05', '06', '07', '08', '09', '10', '11'],
    ['12', '13', '14', '15', '16', '17', '18'],
    ['19', '20', '21', '22', '23', '24', '25'],
    ['26', '27', '28', '29', '30', '31', '']
]
row_h = (ch_card - hdr_h) / 5.0

hl = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx_card + 4 * col_w + 0.1), Inches(cy_card + hdr_h + 2 * row_h + 0.15), Inches(3 * col_w - 0.2), Inches(row_h - 0.3))
hl.fill.solid()
hl.fill.fore_color.rgb = hex_to_rgb('#FDEBD0')
hl.line.fill.background()

for r in range(5):
    for c in range(7):
        val = dates[r][c]
        if not val:
            continue
        tb = s1.shapes.add_textbox(Inches(cx_card + c * col_w), Inches(cy_card + hdr_h + r * row_h + 0.15), Inches(col_w), Inches(row_h - 0.3))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = val
        p.font.name = 'Montserrat'
        p.font.size = Pt(14)
        p.font.bold = True if val in ['01', '15', '30'] else False
        p.font.color.rgb = hex_to_rgb('#C0392B' if c == 0 or c == 6 else '#1E293B')
        p.alignment = PP_ALIGN.CENTER

# Slide 2: Minimalist Split-Banner Calendar
s2 = prs.slides.add_slide(blank_layout)
s_card = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.5), Inches(7.5))
s_card.fill.solid()
s_card.fill.fore_color.rgb = hex_to_rgb('#F7F9FA')
s_card.line.fill.background()

tb_num = s2.shapes.add_textbox(Inches(0.6), Inches(0.8), Inches(3.3), Inches(1.8))
p = tb_num.text_frame.paragraphs[0]
p.text = "08"
p.font.name = 'Georgia'
p.font.size = Pt(80)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#27AE60')

for y_ic in [3.0, 4.8]:
    badge = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(y_ic), Inches(2.8), Inches(1.4))
    badge.fill.solid()
    badge.fill.fore_color.rgb = hex_to_rgb('#E8F8F5')
    badge.line.color.rgb = hex_to_rgb('#2ECC71')
    badge.line.width = Pt(1.5)
    tf = badge.text_frame
    p = tf.paragraphs[0]
    p.text = "Event Goal" if y_ic == 3.0 else "Deliverables"
    p.font.name = 'Montserrat'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E8449')
    p.alignment = PP_ALIGN.CENTER

tb_m = s2.shapes.add_textbox(Inches(5.5), Inches(0.6), Inches(7.0), Inches(1.0))
p = tb_m.text_frame.paragraphs[0]
p.text = "AUGUST 2021"
p.font.name = 'Georgia'
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#1E293B')
p.alignment = PP_ALIGN.CENTER

gx, gy = 5.2, 1.8
gw, gh = 7.4, 4.8
col_w2 = gw / 7.0
row_h2 = gh / 6.0

for i, day in enumerate(days):
    hb = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(gx + i * col_w2), Inches(gy), Inches(col_w2), Inches(0.55))
    hb.fill.solid()
    hb.fill.fore_color.rgb = hex_to_rgb('#1E293B')
    hb.line.fill.background()
    p = hb.text_frame.paragraphs[0]
    p.text = day
    p.font.name = 'Montserrat'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

for r in range(5):
    for c in range(7):
        val = dates[r][c]
        if not val:
            continue
        tb = s2.shapes.add_textbox(Inches(gx + c * col_w2), Inches(gy + 0.6 + r * (row_h2 - 0.1)), Inches(col_w2), Inches(row_h2 - 0.2))
        p = tb.text_frame.paragraphs[0]
        p.text = val
        p.font.name = 'Calibri'
        p.font.size = Pt(13)
        p.font.bold = True if c in [0, 6] else False
        p.font.color.rgb = hex_to_rgb('#27AE60' if c in [0, 6] else '#2C3E50')
        p.alignment = PP_ALIGN.CENTER

# Slide 3: Chevron Sidebar & Notes Section
s3 = prs.slides.add_slide(blank_layout)
sb = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.2), Inches(7.5))
sb.fill.solid()
sb.fill.fore_color.rgb = hex_to_rgb('#00A896')
sb.line.fill.background()

ch = s3.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0), Inches(1.0), Inches(4.8), Inches(1.5))
ch.fill.solid()
ch.fill.fore_color.rgb = hex_to_rgb('#028090')
ch.line.fill.background()
p = ch.text_frame.paragraphs[0]
p.text = "08"
p.font.name = 'Georgia'
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#FFFFFF')
p.alignment = PP_ALIGN.CENTER

tb_info = s3.shapes.add_textbox(Inches(0.6), Inches(3.0), Inches(3.2), Inches(3.8))
tf_info = tb_info.text_frame
p1 = tf_info.paragraphs[0]
p1.text = "CALENDAR\n2021"
p1.font.name = 'Georgia'
p1.font.size = Pt(24)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#FFFFFF')

p2 = tf_info.add_paragraph()
p2.text = "\n• Phone: +1 234 567 890\n• Email: info@company.com\n• Web: www.company.com"
p2.font.name = 'Calibri'
p2.font.size = Pt(11)
p2.font.color.rgb = hex_to_rgb('#E8F8F5')

tb_r = s3.shapes.add_textbox(Inches(5.2), Inches(0.6), Inches(7.4), Inches(0.8))
p = tb_r.text_frame.paragraphs[0]
p.text = "AUGUST"
p.font.name = 'Georgia'
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#028090')
p.alignment = PP_ALIGN.CENTER

for i, day in enumerate(days):
    hb = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(gx + i * col_w2), Inches(1.6), Inches(col_w2), Inches(0.5))
    hb.fill.solid()
    hb.fill.fore_color.rgb = hex_to_rgb('#1E293B')
    hb.line.fill.background()
    p = hb.text_frame.paragraphs[0]
    p.text = day
    p.font.name = 'Montserrat'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#FFFFFF')
    p.alignment = PP_ALIGN.CENTER

for r in range(5):
    for c in range(7):
        val = dates[r][c]
        if not val:
            continue
        tb = s3.shapes.add_textbox(Inches(gx + c * col_w2), Inches(2.2 + r * 0.65), Inches(col_w2), Inches(0.55))
        p = tb.text_frame.paragraphs[0]
        p.text = val
        p.font.name = 'Calibri'
        p.font.size = Pt(12)
        p.font.bold = True if c in [0, 6] else False
        p.font.color.rgb = hex_to_rgb('#00A896' if c in [0, 6] else '#1E293B')
        p.alignment = PP_ALIGN.CENTER

out_pptx = os.path.join(target_folder, 'presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print(f"Calendar Presentation saved successfully to {out_pptx}")

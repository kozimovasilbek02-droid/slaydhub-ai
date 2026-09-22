"""
Master Precision Builder for Pillars, 3C/4P Strategy Frameworks & Pinwheel Infographics
Supports:
- 3_Pillars_Of_Sustainability
- 3_Pillars_Powerpoint_Template
- 4Ps_Of_Marketing
- 4Ps_Of_Corporate_Governance_Ppt
- 4C_Marketing
- 3C_Analysis_Slide_Design_Powerpoint_Ppt_53792
- 3_Step_Process_Ppt_Template
"""
import sys, os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

target_folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\3_Pillars_Of_Sustainability"
folder_name = os.path.basename(os.path.normpath(target_folder))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ----------------------------------------------------
# Slide 1: 3 Classical Pillars Architecture (Greek Temple)
# ----------------------------------------------------
s1 = prs.slides.add_slide(blank_layout)

# Title
tb_t1 = s1.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.333), Inches(0.8))
p = tb_t1.text_frame.paragraphs[0]
p.text = "3 Pillars Of Sustainability"
p.font.name = 'Georgia'
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#1E293B')
p.alignment = PP_ALIGN.CENTER

# Temple Roof Pediment
ped = s1.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1.5), Inches(1.5), Inches(10.333), Inches(1.8))
ped.fill.solid()
ped.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
ped.line.color.rgb = hex_to_rgb('#CBD5E1')

# Central Roof Medallion
med = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(2.0), Inches(1.3), Inches(1.3))
med.fill.solid()
med.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
med.line.fill.background()

# Stepped Podium Foundation (Base)
base_b = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.8), Inches(6.4), Inches(9.733), Inches(0.25))
base_b.fill.solid()
base_b.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
base_b.line.color.rgb = hex_to_rgb('#CBD5E1')

base_t = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.1), Inches(6.15), Inches(9.133), Inches(0.25))
base_t.fill.solid()
base_t.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
base_t.line.fill.background()

pillar_cols = ['#A855F7', '#06B6D4', '#EF4444']
pillar_titles = ["Economic", "Social", "Environment"]
pillar_descs = [
    "Promotes sustainable growth, high job creation, resource efficiency, and financial system stability.",
    "Ensures social equity, community inclusion, poverty reduction, and safeguarding of human rights.",
    "Focuses on biodiversity protection, zero pollution control, and renewable resource management."
]

p_w = 2.4
p_gap = 0.95
start_px = 2.35

for i in range(3):
    px = start_px + i * (p_w + p_gap)
    col = pillar_cols[i]
    
    # Capital (Top)
    cap = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px - 0.2), Inches(3.2), Inches(p_w + 0.4), Inches(0.35))
    cap.fill.solid()
    cap.fill.fore_color.rgb = hex_to_rgb(col)
    cap.line.fill.background()
    
    # Shaft (Body)
    shaft = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(3.55), Inches(p_w), Inches(2.25))
    shaft.fill.solid()
    shaft.fill.fore_color.rgb = hex_to_rgb(col)
    shaft.line.fill.background()
    
    # Plinth (Bottom)
    plinth = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px - 0.2), Inches(5.8), Inches(p_w + 0.4), Inches(0.35))
    plinth.fill.solid()
    plinth.fill.fore_color.rgb = hex_to_rgb(col)
    plinth.line.fill.background()
    
    # Icon Hub
    hub = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px + p_w/2 - 0.4), Inches(3.7), Inches(0.8), Inches(0.8))
    hub.fill.solid()
    hub.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    hub.line.color.rgb = hex_to_rgb(col)
    
    # Text inside Shaft
    tb_s = s1.shapes.add_textbox(Inches(px + 0.1), Inches(4.55), Inches(p_w - 0.2), Inches(1.2))
    tf = tb_s.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = pillar_titles[i]
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = pillar_descs[i]
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = hex_to_rgb('#F8FAFC')
    p2.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# Slide 2: 4Ps / 4Cs Diamond Quadrant Flower
# ----------------------------------------------------
s2 = prs.slides.add_slide(blank_layout)

# Left Title Box
tb_left = s2.shapes.add_textbox(Inches(0.8), Inches(2.8), Inches(4.2), Inches(2.0))
tf = tb_left.text_frame
p1 = tf.paragraphs[0]
p1.text = "4 P's Of"
p1.font.name = 'Georgia'
p1.font.size = Pt(36)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#1E293B')

p2 = tf.add_paragraph()
p2.text = "Marketing"
p2.font.name = 'Georgia'
p2.font.size = Pt(36)
p2.font.bold = True
p2.font.color.rgb = hex_to_rgb('#0EA5E9')

# Center Pinwheel (Right side center: cx=8.5, cy=3.75)
center_x = 8.5
center_y = 3.75
offsets = [
    (0, -1.6, "Product", '#00A896', '#028090'),
    (1.6, 0, "Price", '#E65100', '#D84315'),
    (0, 1.6, "Place", '#1565C0', '#0D47A1'),
    (-1.6, 0, "Promotion", '#F59E0B', '#D97706'),
]

for dx, dy, p_name, col_main, col_dark in offsets:
    # Outer Diamond Card
    d_out = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(center_x + dx - 1.0), Inches(center_y + dy - 1.0), Inches(2.0), Inches(2.0))
    d_out.rotation = 45
    d_out.fill.solid()
    d_out.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    d_out.line.color.rgb = hex_to_rgb('#E2E8F0')
    d_out.line.width = Pt(1.5)
    
    # Text on Outer Card
    tb_p = s2.shapes.add_textbox(Inches(center_x + dx - 1.0), Inches(center_y + dy - 0.3), Inches(2.0), Inches(0.6))
    p = tb_p.text_frame.paragraphs[0]
    p.text = p_name
    p.font.name = 'Montserrat'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#1E293B')
    p.alignment = PP_ALIGN.CENTER
    
    # Inner Colored Diamond
    d_in = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(center_x + dx*0.45 - 0.5), Inches(center_y + dy*0.45 - 0.5), Inches(1.0), Inches(1.0))
    d_in.rotation = 45
    d_in.fill.solid()
    d_in.fill.fore_color.rgb = hex_to_rgb(col_main)
    d_in.line.fill.background()

out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")

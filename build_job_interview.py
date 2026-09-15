import sys
import os
from PIL import Image
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

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Minute_Powerpoint_Presentation_For_Job_Interview'

def add_header(slide, title_text='10 Minute PowerPoint Presentation For Job Interview', align=PP_ALIGN.CENTER, x=0.8, y=0.55, w=11.733):
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

steps_data = [
    ('Introduction:', 'Briefly introduce yourself and the purpose of the presentation.'),
    ('Relevant Skills and Experience:', 'Highlight your relevant skills and experience for the job.'),
    ('Achievements:', 'Mention your significant achievements and how they relate to the job.'),
    ('Examples of Work:', 'Showcase examples of your work to demonstrate your abilities.'),
    ('Closing Remarks:', 'Conclude the presentation with a strong summary of your qualifications and enthusiasm for the position.'),
]

# Extract illustrations cleanly from the original slide images
os.makedirs('output/interview_assets', exist_ok=True)

# Function to crop illustration region
def crop_illustration(slide_num, bbox_norm, out_name):
    img_path = os.path.join(folder_path, f'Slide_{slide_num:02d}.png')
    img = Image.open(img_path)
    w, h = img.size
    x1, y1, x2, y2 = int(bbox_norm[0]*w), int(bbox_norm[1]*h), int(bbox_norm[2]*w), int(bbox_norm[3]*h)
    crop = img.crop((x1, y1, x2, y2))
    out_path = os.path.abspath(f'output/interview_assets/{out_name}.png')
    crop.save(out_path)
    return out_path

# ==================== SLIDE 1: TITLE SLIDE ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
tb_t1 = s1.shapes.add_textbox(Inches(0.80), Inches(2.20), Inches(5.20), Inches(3.00))
tf_t1 = tb_t1.text_frame
tf_t1.word_wrap = True
p1 = tf_t1.paragraphs[0]
p1.text = '10 Minute PowerPoint\nPresentation For\nJob Interview'
p1.font.name = 'Montserrat'
p1.font.size = Pt(34)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#2D3748')

# Crop and place illustration
crop1 = crop_illustration(1, (0.42, 0.18, 0.96, 0.85), 'illus_s1')
s1.shapes.add_picture(crop1, Inches(6.00), Inches(1.20), width=Inches(6.60))

# ==================== SLIDES 2 - 7 ====================
configs = [
    (2, (0.10, 0.22, 0.48, 0.85), 6.50, 0.80, 5.00, '#2B6CB0', '>'),
    (3, (0.50, 0.18, 0.95, 0.90), 0.80, 6.80, 5.80, '#2B6CB0', '➔'),
    (4, (0.05, 0.22, 0.58, 0.85), 7.50, 0.80, 6.20, '#E53E3E', '▶'),
    (5, (0.52, 0.20, 0.96, 0.90), 0.80, 6.80, 5.80, '#DD6B20', '❯'),
    (6, (0.06, 0.22, 0.50, 0.88), 6.80, 0.80, 5.50, '#1A365D', '▶'),
    (7, (0.48, 0.25, 0.95, 0.85), 0.80, 6.50, 5.80, '#2D3748', '❯'),
]

for s_num, bbox, text_x, illus_x, illus_w, color_hex, bullet_char in configs:
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, '10 Minute PowerPoint Presentation For Job Interview')
    
    # Place illustration
    crop_path = crop_illustration(s_num, bbox, f'illus_s{s_num}')
    s.shapes.add_picture(crop_path, Inches(illus_x), Inches(1.50), width=Inches(illus_w))
    
    # Place 5 Steps
    y_base = 1.70
    step_gap = 1.02
    for idx, (head, body) in enumerate(steps_data):
        y_pos = y_base + idx * step_gap
        
        # Soft Drop Shadow Card
        shd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(text_x + 0.02), Inches(y_pos + 0.02), Inches(5.80), Inches(0.85))
        shd.fill.solid()
        shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
        shd.line.fill.background()
        
        # White Card
        crd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(text_x), Inches(y_pos), Inches(5.80), Inches(0.85))
        crd.fill.solid()
        crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        crd.line.color.rgb = hex_to_rgb('#EDF2F7')
        crd.line.width = Pt(1.0)
        
        # Bullet Badge
        bdg = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(text_x + 0.15), Inches(y_pos + 0.22), Inches(0.40), Inches(0.40))
        bdg.fill.solid()
        bdg.fill.fore_color.rgb = hex_to_rgb(color_hex)
        bdg.line.fill.background()
        p_b = bdg.text_frame.paragraphs[0]
        p_b.text = bullet_char
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_b.alignment = PP_ALIGN.CENTER
        
        # Text Frame
        tb = s.shapes.add_textbox(Inches(text_x + 0.68), Inches(y_pos + 0.08), Inches(4.95), Inches(0.70))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p_h = tf.paragraphs[0]
        p_h.text = head
        p_h.font.name = 'Montserrat'
        p_h.font.size = Pt(11)
        p_h.font.bold = True
        p_h.font.color.rgb = hex_to_rgb(color_hex)
        p_h.space_after = Pt(2)
        
        p_desc = tf.add_paragraph()
        p_desc.text = body
        p_desc.font.name = 'Segoe UI'
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = hex_to_rgb('#4A5568')

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Minute_Powerpoint_Presentation_For_Job_Interview\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 4 saved to:', out_file)

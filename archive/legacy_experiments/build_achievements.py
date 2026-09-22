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

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Personal_Achievements_Examples'
os.makedirs('output/achieve_assets', exist_ok=True)

def crop_center_hub(slide_num, out_name):
    img_path = os.path.join(folder_path, f'Slide_{slide_num:02d}.png')
    img = Image.open(img_path)
    w, h = img.size
    # Crop complete circular central graphic
    crop = img.crop((int(0.325*w), int(0.24*h), int(0.675*w), int(0.86*h)))
    out_path = os.path.abspath(f'output/achieve_assets/{out_name}.png')
    crop.save(out_path)
    return out_path

def add_header(slide, title_text='10 Personal Achievements Examples', align=PP_ALIGN.CENTER, x=0.8, y=0.45, w=11.733):
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

ach_left = [
    ('Team Project', 'Guided a group to meet goals and deliver results.', '#E53E3E', '👥'),
    ('Financial Milestone', 'Saved a target amount or built a strong financial habit.', '#DD6B20', '📈'),
    ('Social Cause', 'Contributed time and effort to community service.', '#00A3C4', '🤝'),
    ('Personal Portfolio', 'Created meaningful work that showcases talent.', '#48BB78', '📁'),
    ('Personal Challenge', 'Demonstrated resilience, growth, and determination.', '#3182CE', '🧗'),
]

ach_right = [
    ('Earned Certification', 'Upgraded skills through dedicated training and coursework.', '#48BB78', '📜'),
    ('Fitness Goal', 'Reached a milestone such as weight loss, strength target, or marathon.', '#00A3C4', '🏋️'),
    ('Improved Communication', 'Developed confidence in public speaking or writing.', '#DD6B20', '🗣️'),
    ('Learned New Skill', 'Gained proficiency through continuous practice.', '#E53E3E', '💡'),
    ('Received Award', 'Honored for outstanding performance or contribution.', '#805AD5', '🎖️'),
]

left_y = [1.50, 2.55, 3.60, 4.65, 5.70]
right_y = [1.50, 2.55, 3.60, 4.65, 5.70]

# ==================== SLIDE 1 ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1)

# Place Center Hub Graphic
hub1 = crop_center_hub(1, 'hub_s1')
s1.shapes.add_picture(hub1, Inches(4.35), Inches(1.80), width=Inches(4.65))

for i, (name, desc, col, icon) in enumerate(ach_left):
    c_rgb = hex_to_rgb(col)
    y = left_y[i]
    
    ln = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.00), Inches(y + 0.38), Inches(0.85), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(3.50), Inches(0.78))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(3.50), Inches(0.78))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    ic_d = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.65), Inches(y + 0.09), Inches(0.60), Inches(0.60))
    ic_d.fill.solid()
    ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    ic_d.line.color.rgb = c_rgb
    ic_d.line.width = Pt(1.5)
    p_i = ic_d.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(11)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s1.shapes.add_textbox(Inches(0.95), Inches(y + 0.08), Inches(2.65), Inches(0.62))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.RIGHT

for i, (name, desc, col, icon) in enumerate(ach_right):
    c_rgb = hex_to_rgb(col)
    y = right_y[i]
    
    ln = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.50), Inches(y + 0.38), Inches(0.85), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    ln.line.fill.background()
    
    shd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.02), Inches(y + 0.02), Inches(3.50), Inches(0.78))
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
    shd.line.fill.background()
    
    crd = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.00), Inches(y), Inches(3.50), Inches(0.78))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.color.rgb = c_rgb
    crd.line.width = Pt(1.5)
    
    ic_d = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.08), Inches(y + 0.09), Inches(0.60), Inches(0.60))
    ic_d.fill.solid()
    ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    ic_d.line.color.rgb = c_rgb
    ic_d.line.width = Pt(1.5)
    p_i = ic_d.text_frame.paragraphs[0]
    p_i.text = icon
    p_i.font.size = Pt(11)
    p_i.alignment = PP_ALIGN.CENTER
    
    tb = s1.shapes.add_textbox(Inches(9.75), Inches(y + 0.08), Inches(2.65), Inches(0.62))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = name
    p1.font.name = 'Montserrat'
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#2D3748')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = 'Segoe UI'
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#718096')
    p2.alignment = PP_ALIGN.LEFT

# ==================== SLIDES 2 - 5 ====================
for s_idx in range(2, 6):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s)
    
    hub = crop_center_hub(s_idx, f'hub_s{s_idx}')
    s.shapes.add_picture(hub, Inches(4.35), Inches(1.80), width=Inches(4.65))
    
    for i, (name, desc, col, icon) in enumerate(ach_left):
        c_rgb = hex_to_rgb(col)
        y = left_y[i]
        
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.00), Inches(y + 0.38), Inches(0.85), Inches(0.02))
        ln.fill.solid()
        ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
        ln.line.fill.background()
        
        shd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(y + 0.02), Inches(3.50), Inches(0.78))
        shd.fill.solid()
        shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
        shd.line.fill.background()
        
        crd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(y), Inches(3.50), Inches(0.78))
        crd.fill.solid()
        crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        crd.line.color.rgb = c_rgb
        crd.line.width = Pt(1.5)
        
        ic_d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.65), Inches(y + 0.09), Inches(0.60), Inches(0.60))
        ic_d.fill.solid()
        ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        ic_d.line.color.rgb = c_rgb
        ic_d.line.width = Pt(1.5)
        p_i = ic_d.text_frame.paragraphs[0]
        p_i.text = icon
        p_i.font.size = Pt(11)
        p_i.alignment = PP_ALIGN.CENTER
        
        tb = s.shapes.add_textbox(Inches(0.95), Inches(y + 0.08), Inches(2.65), Inches(0.62))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p1 = tf.paragraphs[0]
        p1.text = name
        p1.font.name = 'Montserrat'
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = hex_to_rgb('#2D3748')
        p1.alignment = PP_ALIGN.RIGHT
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = hex_to_rgb('#718096')
        p2.alignment = PP_ALIGN.RIGHT
        
    for i, (name, desc, col, icon) in enumerate(ach_right):
        c_rgb = hex_to_rgb(col)
        y = right_y[i]
        
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.50), Inches(y + 0.38), Inches(0.85), Inches(0.02))
        ln.fill.solid()
        ln.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
        ln.line.fill.background()
        
        shd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.02), Inches(y + 0.02), Inches(3.50), Inches(0.78))
        shd.fill.solid()
        shd.fill.fore_color.rgb = hex_to_rgb('#E2E8F0')
        shd.line.fill.background()
        
        crd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.00), Inches(y), Inches(3.50), Inches(0.78))
        crd.fill.solid()
        crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        crd.line.color.rgb = c_rgb
        crd.line.width = Pt(1.5)
        
        ic_d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.08), Inches(y + 0.09), Inches(0.60), Inches(0.60))
        ic_d.fill.solid()
        ic_d.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        ic_d.line.color.rgb = c_rgb
        ic_d.line.width = Pt(1.5)
        p_i = ic_d.text_frame.paragraphs[0]
        p_i.text = icon
        p_i.font.size = Pt(11)
        p_i.alignment = PP_ALIGN.CENTER
        
        tb = s.shapes.add_textbox(Inches(9.75), Inches(y + 0.08), Inches(2.65), Inches(0.62))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p1 = tf.paragraphs[0]
        p1.text = name
        p1.font.name = 'Montserrat'
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = hex_to_rgb('#2D3748')
        p1.alignment = PP_ALIGN.LEFT
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = hex_to_rgb('#718096')
        p2.alignment = PP_ALIGN.LEFT

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Personal_Achievements_Examples\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 5 saved to:', out_file)

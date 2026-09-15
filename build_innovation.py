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

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Types_Of_Innovation_Ppt'
os.makedirs('output/innov_assets', exist_ok=True)

# Crop central brain-bulb-hand graphic
img_path = os.path.join(folder_path, 'Slide_01.png')
img = Image.open(img_path)
w, h = img.size
crop = img.crop((int(0.24*w), int(0.16*h), int(0.76*w), int(0.96*h)))
crop_path = os.path.abspath('output/innov_assets/bulb_brain.png')
crop.save(crop_path)

s = prs.slides.add_slide(prs.slide_layouts[6])

# Title
tb = s.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.85))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'Infographics Of Ideas'
p.font.name = 'Montserrat'
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = hex_to_rgb('#2D3748')
p.alignment = PP_ALIGN.CENTER

# Place center graphic
s.shapes.add_picture(crop_path, Inches(3.20), Inches(1.30), width=Inches(6.93))

# Left Top Callout: Idea
disc1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.25), Inches(2.00), Inches(1.00), Inches(1.00))
disc1.fill.solid()
disc1.fill.fore_color.rgb = hex_to_rgb('#F6C744')
disc1.line.fill.background()
p_i1 = disc1.text_frame.paragraphs[0]
p_i1.text = '💡'
p_i1.font.size = Pt(18)
p_i1.alignment = PP_ALIGN.CENTER

tb1 = s.shapes.add_textbox(Inches(0.60), Inches(2.05), Inches(2.50), Inches(1.00))
tf1 = tb1.text_frame
tf1.word_wrap = True
tf1.margin_left = tf1.margin_right = tf1.margin_top = tf1.margin_bottom = 0
p1_1 = tf1.paragraphs[0]
p1_1.text = 'Idea'
p1_1.font.name = 'Montserrat'
p1_1.font.size = Pt(13)
p1_1.font.bold = True
p1_1.font.color.rgb = hex_to_rgb('#2D3748')
p1_1.alignment = PP_ALIGN.RIGHT
p1_2 = tf1.add_paragraph()
p1_2.text = 'Fusce lobortis porttitor\npurus, vel vestibulum libero'
p1_2.font.name = 'Segoe UI'
p1_2.font.size = Pt(9.5)
p1_2.font.color.rgb = hex_to_rgb('#718096')
p1_2.alignment = PP_ALIGN.RIGHT

# Left Bottom Callout: Research
disc2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.25), Inches(4.00), Inches(1.00), Inches(1.00))
disc2.fill.solid()
disc2.fill.fore_color.rgb = hex_to_rgb('#B76E79')
disc2.line.fill.background()
p_i2 = disc2.text_frame.paragraphs[0]
p_i2.text = '🔍'
p_i2.font.size = Pt(18)
p_i2.alignment = PP_ALIGN.CENTER

tb2 = s.shapes.add_textbox(Inches(0.60), Inches(4.05), Inches(2.50), Inches(1.00))
tf2 = tb2.text_frame
tf2.word_wrap = True
tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = 0
p2_1 = tf2.paragraphs[0]
p2_1.text = 'Research'
p2_1.font.name = 'Montserrat'
p2_1.font.size = Pt(13)
p2_1.font.bold = True
p2_1.font.color.rgb = hex_to_rgb('#2D3748')
p2_1.alignment = PP_ALIGN.RIGHT
p2_2 = tf2.add_paragraph()
p2_2.text = 'Fusce lobortis porttitor\npurus, vel vestibulum libero'
p2_2.font.name = 'Segoe UI'
p2_2.font.size = Pt(9.5)
p2_2.font.color.rgb = hex_to_rgb('#718096')
p2_2.alignment = PP_ALIGN.RIGHT

# Right Top Callout: Teamwork
disc3 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.08), Inches(2.00), Inches(1.00), Inches(1.00))
disc3.fill.solid()
disc3.fill.fore_color.rgb = hex_to_rgb('#E07A5F')
disc3.line.fill.background()
p_i3 = disc3.text_frame.paragraphs[0]
p_i3.text = '🤝'
p_i3.font.size = Pt(18)
p_i3.alignment = PP_ALIGN.CENTER

tb3 = s.shapes.add_textbox(Inches(10.25), Inches(2.05), Inches(2.50), Inches(1.00))
tf3 = tb3.text_frame
tf3.word_wrap = True
tf3.margin_left = tf3.margin_right = tf3.margin_top = tf3.margin_bottom = 0
p3_1 = tf3.paragraphs[0]
p3_1.text = 'Teamwork'
p3_1.font.name = 'Montserrat'
p3_1.font.size = Pt(13)
p3_1.font.bold = True
p3_1.font.color.rgb = hex_to_rgb('#2D3748')
p3_1.alignment = PP_ALIGN.LEFT
p3_2 = tf3.add_paragraph()
p3_2.text = 'Fusce lobortis porttitor\npurus, vel vestibulum libero'
p3_2.font.name = 'Segoe UI'
p3_2.font.size = Pt(9.5)
p3_2.font.color.rgb = hex_to_rgb('#718096')
p3_2.alignment = PP_ALIGN.LEFT

# Right Bottom Callout: Success
disc4 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.08), Inches(4.00), Inches(1.00), Inches(1.00))
disc4.fill.solid()
disc4.fill.fore_color.rgb = hex_to_rgb('#4EA8DE')
disc4.line.fill.background()
p_i4 = disc4.text_frame.paragraphs[0]
p_i4.text = '🏆'
p_i4.font.size = Pt(18)
p_i4.alignment = PP_ALIGN.CENTER

tb4 = s.shapes.add_textbox(Inches(10.25), Inches(4.05), Inches(2.50), Inches(1.00))
tf4 = tb4.text_frame
tf4.word_wrap = True
tf4.margin_left = tf4.margin_right = tf4.margin_top = tf4.margin_bottom = 0
p4_1 = tf4.paragraphs[0]
p4_1.text = 'Success'
p4_1.font.name = 'Montserrat'
p4_1.font.size = Pt(13)
p4_1.font.bold = True
p4_1.font.color.rgb = hex_to_rgb('#2D3748')
p4_1.alignment = PP_ALIGN.LEFT
p4_2 = tf4.add_paragraph()
p4_2.text = 'Fusce lobortis porttitor\npurus, vel vestibulum libero'
p4_2.font.name = 'Segoe UI'
p4_2.font.size = Pt(9.5)
p4_2.font.color.rgb = hex_to_rgb('#718096')
p4_2.alignment = PP_ALIGN.LEFT

out_file = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10_Types_Of_Innovation_Ppt\presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Master presentation for Folder 7 saved to:', out_file)

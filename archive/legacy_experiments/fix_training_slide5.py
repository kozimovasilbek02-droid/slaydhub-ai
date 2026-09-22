import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, hex_to_rgb, add_slide_header, add_card_text, add_soft_shadow

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

folder_path = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Training_Plan'
os.makedirs('output/training_assets', exist_ok=True)

img_path = os.path.join(folder_path, 'Slide_05.png')
img = Image.open(img_path)
w, h = img.size

# 1. Crop complete left cluster (Calendar Wheel + Arc Spine + Numbered Badges 1-5 + Line Connectors)
crop_left = img.crop((int(0.00*w), int(0.20*h), int(0.31*w), int(0.96*h)))
left_path = os.path.abspath('output/training_assets/left_cluster_s5.png')
crop_left.save(left_path)

# 2. Crop 3D buttons for each of the 5 phases
btn_paths = []
btn_crops_y = [(0.21, 0.35), (0.37, 0.50), (0.52, 0.65), (0.68, 0.81), (0.83, 0.96)]
for idx, (y1_r, y2_r) in enumerate(btn_crops_y):
    crop_btn = img.crop((int(0.295*w), int(y1_r*h), int(0.385*w), int(y2_r*h)))
    p_out = os.path.abspath(f'output/training_assets/btn_clean_{idx+1}.png')
    crop_btn.save(p_out)
    btn_paths.append(p_out)

s = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s, 'Training Phases', x=0.80, y=0.55, w=11.733)

# Place Left Cluster Graphic (contains calendar, arc spine, badges 1-5, and connectors seamlessly)
s.shapes.add_picture(left_path, Inches(0.10), Inches(1.50), width=Inches(4.15))

# 5 Phases Data
phases = [
    ('Base Phase', 'Establish aerobic endurance, build strength, develop technique, focus on consistency,', '#0070F3'),
    ('Peak Phase', 'Maximize performance, fine-tune skills, reduce volume, sharpen mental focus.', '#5B45E0'),
    ('Race Phase', 'Execute strategy, perform at peak, manage effort, maintain competitive mindset.', '#F27A24'),
    ('Build Phase', 'Increase intensity, add intervals, improve speed, build muscular and cardiovascular power.', '#E53E3E'),
    ('Taper Phase', 'Reduce training load, enhance recovery, maintain intensity, prepare for competition.', '#00B894'),
]

card_xs = 4.10
card_ys = [1.88, 3.08, 4.28, 5.48, 6.68]
card_w = 3.50
card_h = 0.88

# Draw 5 floating cards + 3D buttons + text
for i, (title, desc, col_hex) in enumerate(phases):
    cy_i = card_ys[i]
    top_y = cy_i - card_h / 2.0
    
    # Floating card background with soft shadow
    add_soft_shadow(s, MSO_SHAPE.ROUNDED_RECTANGLE, card_xs, top_y, card_w, card_h, 0.03, 0.03)
    crd = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(card_xs), Inches(top_y), Inches(card_w), Inches(card_h))
    crd.fill.solid()
    crd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    crd.line.fill.background()
    
    # 3D Button Graphic
    s.shapes.add_picture(btn_paths[i], Inches(card_xs - 0.65), Inches(top_y - 0.08), width=Inches(1.20))
    
    # Card Title
    tb_title = s.shapes.add_textbox(Inches(card_xs + 0.65), Inches(top_y + 0.24), Inches(2.65), Inches(0.40))
    tf_t = tb_title.text_frame
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = Fonts.TITLE
    p_t.font.size = Pt(12.5)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb('#1A202C')
    p_t.alignment = PP_ALIGN.LEFT
    
    # Right Description Block
    tb_desc = s.shapes.add_textbox(Inches(7.90), Inches(top_y + 0.12), Inches(4.90), Inches(card_h))
    tf_d = tb_desc.text_frame
    tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]
    p_d.text = desc
    p_d.font.name = Fonts.BODY
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = hex_to_rgb(Colors.TEXT_TITLE)
    p_d.alignment = PP_ALIGN.LEFT

out_pptx = os.path.abspath('output/training_s5_flawless.pptx')
prs.save(out_pptx)
print('Saved test PPTX to:', out_pptx)

# Export and preview
exported = VisionQA.export_slides_to_png(out_pptx, output_dir='output/training_flawless_previews')
print('Exported preview to:', exported[0])

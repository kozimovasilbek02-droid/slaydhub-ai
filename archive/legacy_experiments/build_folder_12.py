import sys
import os
import shutil
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, hex_to_rgb, VisualDiff

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\1920S_Theme'
os.makedirs('output/f12_assets', exist_ok=True)

BG_RGB = np.array([13, 13, 13, 255], dtype=np.uint8)

# Slide 1: Title Frame Asset
im1 = Image.open(os.path.join(folder_path, 'Slide_01.png')).convert('RGBA')
arr1 = np.array(im1)
w1, h1 = im1.size
# Mask ONLY the inner cartouche text area (y: 50% to 68%, x: 24% to 76%)
arr1[int(0.50*h1):int(0.68*h1), int(0.24*w1):int(0.76*w1)] = BG_RGB
clean_s1_bg = Image.fromarray(arr1)
p1_bg_path = os.path.abspath('output/f12_assets/s1_artdeco_title_bg.png')
clean_s1_bg.save(p1_bg_path)

# Slide 2-5: Content Frame Asset with 100% INTACT top and bottom gold divider lines
im2 = Image.open(os.path.join(folder_path, 'Slide_02.png')).convert('RGBA')
arr2 = np.array(im2)
w2, h2 = im2.size
# The top horizontal divider line is at y: 26.5% - 27.5%
# The bottom horizontal divider line is at y: 72.5% - 73.5%
# Therefore, mask ONLY the text box between y: 31% and y: 69%
arr2[int(0.31*h2):int(0.69*h2), int(0.12*w2):int(0.88*w2)] = BG_RGB
clean_content_bg = Image.fromarray(arr2)
p_content_bg_path = os.path.abspath('output/f12_assets/artdeco_content_bg_fixed.png')
clean_content_bg.save(p_content_bg_path)

# Theme typography styling
GOLD_TITLE = '#E0C79B'  # Soft warm Art Deco gold
GOLD_BODY = '#CDB896'   # Elegant vintage body text
SERIF_FONT = 'Georgia'

# ==================== SLIDE 1: TITLE SLIDE ====================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
s1.shapes.add_picture(p1_bg_path, 0, 0, width=Inches(13.333), height=Inches(7.5))

tb1 = s1.shapes.add_textbox(Inches(2.00), Inches(3.90), Inches(9.333), Inches(1.50))
tf1 = tb1.text_frame
tf1.word_wrap = True
p1 = tf1.paragraphs[0]
p1.text = "1920s Theme"
p1.font.name = SERIF_FONT
p1.font.size = Pt(48)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb(GOLD_TITLE)
p1.alignment = PP_ALIGN.CENTER

# ==================== SLIDES 2 TO 5: CONTENT SLIDES ====================
slides_content = [
    (2, "Geometric Shapes", "Art Deco design is famous for its use of bold geometric shapes, including zigzags, chevrons, triangles, and sunbursts. These shapes are often used in repeating patterns or as decorative motifs on furniture, architecture, and textiles."),
    (3, "Luxurious Materials", "Art Deco designers favored luxurious materials, such as exotic woods, metals, and precious stones, to create furniture and decorative objects. The use of materials such as chrome, stainless steel, and mirrored surfaces were also popular."),
    (4, "Streamlined Aesthetic", "Art Deco design embraced the streamlined aesthetic, which emphasized speed and modernity. This style is characterized by clean lines, smooth surfaces, and a futuristic look. Art Deco designers often drew inspiration from the world of transportation, such as airplanes, trains, and automobiles."),
    (5, "Decorative Motifs", "Art Deco designers used decorative motifs to add flair to their designs. Some of the most popular motifs include sunbursts, stylized flowers, and zigzags. These motifs were often used in repeating patterns or as accent pieces on furniture and decorative objects.")
]

for s_num, title, body in slides_content:
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(p_content_bg_path, 0, 0, width=Inches(13.333), height=Inches(7.5))
    
    tb = s.shapes.add_textbox(Inches(1.80), Inches(2.50), Inches(9.733), Inches(3.20))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = SERIF_FONT
    p_t.font.size = Pt(34)
    p_t.font.bold = True
    p_t.font.color.rgb = hex_to_rgb(GOLD_TITLE)
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(18)
    
    p_b = tf.add_paragraph()
    p_b.text = body
    p_b.font.name = SERIF_FONT
    p_b.font.size = Pt(13.5)
    p_b.font.color.rgb = hex_to_rgb(GOLD_BODY)
    p_b.alignment = PP_ALIGN.CENTER
    p_b.line_spacing = 1.35

out_local = os.path.abspath('output/f12_presentation_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Saved local presentation to:', out_local)

dst_master = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\1920S_Theme\presentation_95_precision.pptx')
try:
    shutil.copy(out_local, dst_master)
    print('Master presentation updated at:', dst_master)
except Exception as e:
    print('Destination file open in PowerPoint:', e)

# Export previews
exported = VisionQA.export_slides_to_png(out_local, output_dir='output/folder12_fixed_previews')
print(f'Exported {len(exported)} preview slides for Folder 12.')

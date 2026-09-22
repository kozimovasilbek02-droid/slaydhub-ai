import sys
import os
import shutil
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, hex_to_rgb, add_slide_header, add_soft_shadow

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\1_Year_Financial_Roadmap'
os.makedirs('output/f13_assets', exist_ok=True)

def get_img(s_num):
    p = os.path.join(folder_path, f'Slide_{s_num:02d}.png')
    return Image.open(p).convert('RGBA')

# 5 Phase Definitions
PHASES = [
    {"name": "Plan", "color": "#C2185B", "desc": "Managing and growing customer relationships", "detail": "Outlines goals and objectives for managing and growing customer relationships."},
    {"name": "Communicate", "color": "#FF5722", "desc": "Ideas between sales teams, account managers", "detail": "Exchange of information and ideas between sales teams, account managers, and customers"},
    {"name": "Execute", "color": "#FFA000", "desc": "Strategies and action steps", "detail": "Implementation of strategies and action steps outlined in the plan"},
    {"name": "Measure", "color": "#00897B", "desc": "Evaluating the progress and success", "detail": "Process of assessing and evaluating the progress and success of the strategies and action"},
    {"name": "Discover", "color": "#1B2A4A", "desc": "Researching and understanding the customer's business", "detail": "Process of researching and understanding the customer's business, industry, challenges, and goals"}
]

# ==================== SLIDE 1: TITLE SLIDE ====================
im1 = get_img(1)
w1, h1 = im1.size
c1 = im1.crop((0, 0, int(0.58*w1), h1))
p1_ill = os.path.abspath('output/f13_assets/s1_ill.png')
c1.save(p1_ill)

s1 = prs.slides.add_slide(prs.slide_layouts[6])
s1.shapes.add_picture(p1_ill, Inches(0.40), Inches(0.20), width=Inches(7.20))

tb1 = s1.shapes.add_textbox(Inches(7.80), Inches(2.20), Inches(5.00), Inches(2.80))
tf1 = tb1.text_frame
tf1.word_wrap = True
p1_1 = tf1.paragraphs[0]
p1_1.text = "Account\nPlan"
p1_1.font.name = 'Georgia'
p1_1.font.size = Pt(50)
p1_1.font.bold = True
p1_1.font.color.rgb = hex_to_rgb('#2D3748')

bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.85), Inches(4.75), Inches(3.60), Inches(0.06))
bar1.fill.solid()
bar1.fill.fore_color.rgb = hex_to_rgb('#00897B')
bar1.line.fill.background()

# ==================== SLIDES 2 TO 15 ====================
for s_num in range(2, 16):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    im = get_img(s_num)
    w, h = im.size
    arr = np.array(im)
    
    # Check if this slide needs a top header
    has_top_header = s_num not in [4, 8]
    if has_top_header:
        arr[:int(0.18*h), :] = [255, 255, 255, 255]
    
    # 1. Slide 2: 5 Vertical Location Pins
    if s_num == 2:
        arr[int(0.50*h):int(0.80*h), :] = [255, 255, 255, 255] # text inside colored blocks
        
        p_bg = os.path.abspath(f'output/f13_assets/s{s_num}_clean_bg.png')
        Image.fromarray(arr).save(p_bg)
        s.shapes.add_picture(p_bg, 0, 0, width=Inches(13.333), height=Inches(7.5))
        
        add_slide_header(s, 'Account Plan', x=0.80, y=0.55, w=11.733, title_size=32)
        
        col_w = 2.15
        gap = 0.30
        for i, p in enumerate(PHASES):
            cx = 0.70 + i * (col_w + gap)
            tb_title = s.shapes.add_textbox(Inches(cx), Inches(3.75), Inches(col_w), Inches(0.60))
            tf_t = tb_title.text_frame
            tf_t.word_wrap = True
            pt = tf_t.paragraphs[0]
            pt.text = p['name']
            pt.font.name = Fonts.TITLE
            pt.font.size = Pt(14)
            pt.font.bold = True
            pt.font.color.rgb = hex_to_rgb('#FFFFFF')
            pt.alignment = PP_ALIGN.CENTER
            
            tb_body = s.shapes.add_textbox(Inches(cx + 0.10), Inches(4.55), Inches(col_w - 0.20), Inches(1.30))
            tf_b = tb_body.text_frame
            tf_b.word_wrap = True
            pb = tf_b.paragraphs[0]
            pb.text = p['desc']
            pb.font.name = Fonts.BODY
            pb.font.size = Pt(10.5)
            pb.font.color.rgb = hex_to_rgb('#FFFFFF')
            pb.alignment = PP_ALIGN.CENTER
            pb.line_spacing = 1.15
            
    # 2. Slides 3, 7, 9, 11, 12, 13, 15: Horizontal Diagrams with bottom text
    elif s_num in [3, 7, 9, 11, 12, 13, 15]:
        arr[int(0.70*h):, :] = [255, 255, 255, 255] # bottom text labels area
        
        p_bg = os.path.abspath(f'output/f13_assets/s{s_num}_clean_bg.png')
        Image.fromarray(arr).save(p_bg)
        s.shapes.add_picture(p_bg, 0, 0, width=Inches(13.333), height=Inches(7.5))
        
        add_slide_header(s, 'Account Plan', x=0.80, y=0.55, w=11.733, title_size=32)
        
        col_w = 2.25
        gap = 0.25
        for i, p in enumerate(PHASES):
            cx = 0.85 + i * (col_w + gap)
            tb = s.shapes.add_textbox(Inches(cx), Inches(5.40), Inches(col_w), Inches(1.80))
            tf = tb.text_frame
            tf.word_wrap = True
            
            pt = tf.paragraphs[0]
            pt.text = p['name']
            pt.font.name = Fonts.TITLE
            pt.font.size = Pt(13)
            pt.font.bold = True
            pt.font.color.rgb = hex_to_rgb('#2D3748')
            pt.alignment = PP_ALIGN.CENTER
            pt.space_after = Pt(4)
            
            pb = tf.add_paragraph()
            pb.text = p['desc']
            pb.font.name = Fonts.BODY
            pb.font.size = Pt(10)
            pb.font.color.rgb = hex_to_rgb('#64748B')
            pb.alignment = PP_ALIGN.CENTER
            pb.line_spacing = 1.2
            
    # 3. Slides 4, 5, 6, 8, 10, 14: Left Graphic + Right Vertical List
    elif s_num in [4, 5, 6, 8, 10, 14]:
        if s_num == 4:
            arr[:, int(0.66*w):] = [255, 255, 255, 255]
        elif s_num in [5, 8]:
            arr[:, int(0.65*w):] = [255, 255, 255, 255]
        elif s_num == 6:
            arr[:, int(0.66*w):] = [255, 255, 255, 255]
        elif s_num == 10:
            arr[:, int(0.58*w):] = [255, 255, 255, 255]
        elif s_num == 14:
            arr[:, int(0.73*w):] = [255, 255, 255, 255]
            
        p_bg = os.path.abspath(f'output/f13_assets/s{s_num}_clean_bg.png')
        Image.fromarray(arr).save(p_bg)
        s.shapes.add_picture(p_bg, 0, 0, width=Inches(13.333), height=Inches(7.5))
        
        if has_top_header:
            add_slide_header(s, 'Account Plan', x=0.80, y=0.55, w=11.733, title_size=32)
        
        if s_num in [4, 6]:
            y_start = 1.30
            y_gap = 1.15
            for i, p in enumerate(PHASES):
                y_pos = y_start + i * y_gap
                tb = s.shapes.add_textbox(Inches(8.90), Inches(y_pos), Inches(4.10), Inches(1.10))
                tf = tb.text_frame
                tf.word_wrap = True
                
                pt = tf.paragraphs[0]
                pt.text = p['name']
                pt.font.name = Fonts.TITLE
                pt.font.size = Pt(13.5)
                pt.font.bold = True
                pt.font.color.rgb = hex_to_rgb('#2D3748')
                pt.space_after = Pt(2)
                
                pb = tf.add_paragraph()
                pb.text = p['detail']
                pb.font.name = Fonts.BODY
                pb.font.size = Pt(9.5)
                pb.font.color.rgb = hex_to_rgb('#64748B')
                pb.line_spacing = 1.15
                
        elif s_num in [5, 8]:
            y_start = 1.30
            y_gap = 1.16
            for i, p in enumerate(PHASES):
                y_pos = y_start + i * y_gap
                tb = s.shapes.add_textbox(Inches(9.00), Inches(y_pos), Inches(4.00), Inches(1.10))
                tf = tb.text_frame
                tf.word_wrap = True
                
                pt = tf.paragraphs[0]
                pt.text = p['name']
                pt.font.name = Fonts.TITLE
                pt.font.size = Pt(13.5)
                pt.font.bold = True
                pt.font.color.rgb = hex_to_rgb('#2D3748')
                pt.space_after = Pt(2)
                
                pb = tf.add_paragraph()
                pb.text = p['detail']
                pb.font.name = Fonts.BODY
                pb.font.size = Pt(9.5)
                pb.font.color.rgb = hex_to_rgb('#64748B')
                pb.line_spacing = 1.15
                
        elif s_num == 10:
            y_start = 1.40
            y_gap = 1.10
            for i, p in enumerate(PHASES):
                y_pos = y_start + i * y_gap
                tb = s.shapes.add_textbox(Inches(7.85), Inches(y_pos), Inches(5.10), Inches(1.10))
                tf = tb.text_frame
                tf.word_wrap = True
                
                pt = tf.paragraphs[0]
                pt.text = p['name']
                pt.font.name = Fonts.TITLE
                pt.font.size = Pt(13.5)
                pt.font.bold = True
                pt.font.color.rgb = hex_to_rgb('#2D3748')
                pt.space_after = Pt(2)
                
                pb = tf.add_paragraph()
                pb.text = p['detail']
                pb.font.name = Fonts.BODY
                pb.font.size = Pt(9.5)
                pb.font.color.rgb = hex_to_rgb('#64748B')
                pb.line_spacing = 1.15
                
        elif s_num == 14:
            y_start = 1.30
            y_gap = 1.18
            for i, p in enumerate(PHASES):
                y_pos = y_start + i * y_gap
                tb_desc = s.shapes.add_textbox(Inches(8.80), Inches(y_pos + 0.05), Inches(4.30), Inches(1.00))
                tf_d = tb_desc.text_frame
                tf_d.word_wrap = True
                pb = tf_d.paragraphs[0]
                pb.text = p['detail']
                pb.font.name = Fonts.BODY
                pb.font.size = Pt(9.5)
                pb.font.color.rgb = hex_to_rgb('#64748B')
                pb.line_spacing = 1.15

out_local = os.path.abspath('output/f13_presentation_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Saved Folder 13 local presentation to:', out_local)

dst_master = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\1_Year_Financial_Roadmap\presentation_95_precision.pptx')
try:
    shutil.copy(out_local, dst_master)
    print('Master presentation for Folder 13 updated at:', dst_master)
except Exception as e:
    print('Destination file open in PowerPoint:', e)

# Export previews
exported = VisionQA.export_slides_to_png(out_local, output_dir='output/folder13_perfect_previews')
print(f'Exported {len(exported)} perfect preview slides for Folder 13.')

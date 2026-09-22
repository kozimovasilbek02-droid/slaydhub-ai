import os
import math
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, VisualDiff, hex_to_rgb, add_slide_header, add_soft_shadow

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_header(s2, '100 Days Action Plan', subtitle_text='In 4 Phases', x=0.80, y=1.00, w=11.733, title_size=32)

def create_smooth_wave_arc(slide, cx, cy, r_in, r_out, is_top=True, is_first=False, is_last=False, color='#88C13F', num_pts=36):
    """
    Creates an exact crescent arc where the overlapping ends terminate inside the adjacent card
    with zero exposed flat end-caps or steps on the horizontal midline.
    """
    # Angle calculation: terminate at junction angle theta_junc so it hides completely under neighbor card
    theta_junc = math.radians(24.0)  # ~24 degrees cutoff inside card mask
    
    if is_top:
        start_angle = math.pi if is_first else (math.pi - theta_junc)
        end_angle = theta_junc if not is_last else 0
    else:
        start_angle = (math.pi + theta_junc) if not is_first else math.pi
        end_angle = (2 * math.pi - theta_junc) if not is_last else (2 * math.pi)
        
    pts = []
    # Outer arc points
    for i in range(num_pts):
        theta = start_angle + (end_angle - start_angle) * (i / (num_pts - 1))
        px = cx + r_out * math.cos(theta)
        py = cy - r_out * math.sin(theta)
        pts.append((Inches(px), Inches(py)))
        
    # Inner arc points (reverse)
    for i in range(num_pts - 1, -1, -1):
        theta = start_angle + (end_angle - start_angle) * (i / (num_pts - 1))
        px = cx + r_in * math.cos(theta)
        py = cy - r_in * math.sin(theta)
        pts.append((Inches(px), Inches(py)))
        
    # Drop shadow
    shd_pts = [(p[0] + Inches(0.03), p[1] + Inches(0.03)) for p in pts]
    fb_shd = slide.shapes.build_freeform(shd_pts[0][0], shd_pts[0][1])
    fb_shd.add_line_segments(shd_pts[1:], close=True)
    shd = fb_shd.convert_to_shape()
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    shd.line.fill.background()
    
    # Colored shape
    fb = slide.shapes.build_freeform(pts[0][0], pts[0][1])
    fb.add_line_segments(pts[1:], close=True)
    shape = fb.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.fill.background()
    return shape

phases_data = [
    {
        'title': 'First 25 Days',
        'bullets': ['Observe for understanding', 'Start opportunity assessment'],
        'color': '#88C13F',
        'badge': '01',
        'arc_top': True,
        'cx': 2.38,
        'is_first': True,
        'is_last': False
    },
    {
        'title': '25-50 Days',
        'bullets': ['Build month sales calendar', 'Sales reporting'],
        'color': '#29B6D3',
        'badge': '02',
        'arc_top': False,
        'cx': 5.22,
        'is_first': False,
        'is_last': False
    },
    {
        'title': '50-75 Days',
        'bullets': ['Forecasting', 'Report assessment findings', 'Introduce changes in roles, responsibilities', 'Constructive feedback'],
        'color': '#F5A623',
        'badge': '03',
        'arc_top': True,
        'cx': 8.06,
        'is_first': False,
        'is_last': False
    },
    {
        'title': '75-100 Days',
        'bullets': ['Introduce changes in roles, responsibilities'],
        'color': '#E54343',
        'badge': '04',
        'arc_top': False,
        'cx': 10.90,
        'is_first': False,
        'is_last': True
    }
]

cy = 4.75
r_out = 1.55
r_in = 1.34
r_card = 1.35

for p in phases_data:
    cx = p['cx']
    c_hex = p['color']
    c_rgb = hex_to_rgb(c_hex)
    arc_top = p['arc_top']
    badge_num = p['badge']
    
    # 1. Smooth Wave Arc with Zero Step Endcaps
    create_smooth_wave_arc(s2, cx, cy, r_in, r_out, is_top=arc_top, 
                           is_first=p['is_first'], is_last=p['is_last'], color=c_hex)
    
    # 2. Number Badge (Behind white card)
    badge_size = 0.90
    badge_cx = cx
    badge_cy = (cy - r_card) if not arc_top else (cy + r_card)
    
    add_soft_shadow(s2, MSO_SHAPE.OVAL, badge_cx - badge_size/2.0, badge_cy - badge_size/2.0, badge_size, badge_size, 0.03, 0.03)
    bd = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_cx - badge_size/2.0), Inches(badge_cy - badge_size/2.0), Inches(badge_size), Inches(badge_size))
    bd.fill.solid()
    bd.fill.fore_color.rgb = c_rgb
    bd.line.fill.background()
    
    # 3. White Center Circular Card (Smooth shadow, overlapping previous card)
    add_soft_shadow(s2, MSO_SHAPE.OVAL, cx - r_card, cy - r_card, r_card*2, r_card*2, 0.035, 0.035)
    c_in = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - r_card), Inches(cy - r_card), Inches(r_card*2), Inches(r_card*2))
    c_in.fill.solid()
    c_in.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    c_in.line.fill.background()
    
    # 4. Badge Number (Foreground)
    tb_badge_y = (badge_cy - 0.32) if not arc_top else (badge_cy + 0.04)
    tb_badge = s2.shapes.add_textbox(Inches(badge_cx - 0.40), Inches(tb_badge_y), Inches(0.80), Inches(0.35))
    p_num = tb_badge.text_frame.paragraphs[0]
    p_num.text = badge_num
    p_num.font.name = Fonts.TITLE
    p_num.font.size = Pt(13)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER
    
    # 5. Text Frame inside White Card
    tb_w = 2.15
    tb_h = 1.95
    y_offset = -0.04 if not arc_top else 0.04
    tb = s2.shapes.add_textbox(Inches(cx - tb_w/2.0), Inches(cy - tb_h/2.0 + y_offset), Inches(tb_w), Inches(tb_h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    
    p_title = tf.paragraphs[0]
    p_title.text = p['title']
    p_title.font.name = Fonts.TITLE
    p_title.font.size = Pt(13.5)
    p_title.font.bold = True
    p_title.font.color.rgb = hex_to_rgb('#1A202C')
    p_title.alignment = PP_ALIGN.CENTER
    p_title.space_after = Pt(6)
    
    for b_text in p['bullets']:
        p_bullet = tf.add_paragraph()
        p_bullet.text = f"•  {b_text}"
        p_bullet.font.name = Fonts.BODY
        p_bullet.font.size = Pt(8.8)
        p_bullet.font.color.rgb = hex_to_rgb('#4A5568')
        p_bullet.alignment = PP_ALIGN.LEFT
        p_bullet.space_after = Pt(2.5)

out_file = os.path.abspath('output/action_plan_s2_flawless_chain.pptx')
prs.save(out_file)

# Export and preview
exported = VisionQA.export_slides_to_png(out_file, output_dir='output/action_plan_previews')
print('Exported to:', exported[0])

# Crop junction to verify zero end-caps
im_gen = Image.open(exported[0])
w_g, h_g = im_gen.size
crop_junc_gen = im_gen.crop((int(0.20*w_g), int(0.35*h_g), int(0.45*w_g), int(0.85*h_g)))
junc_gen_path = os.path.abspath('output/inspection/flawless_chain_junction.png')
crop_junc_gen.save(junc_gen_path)
print('Saved junction to:', junc_gen_path)

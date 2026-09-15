import math
from typing import List, Dict, Any
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine.theme import hex_to_rgb, Fonts, Colors
from ppt_engine.geometry import add_soft_shadow
from ppt_engine.typography import add_slide_header

def create_smooth_wave_arc(slide, cx: float, cy: float, r_in: float, r_out: float, 
                           is_top=True, is_first=False, is_last=False, color='#88C13F', num_pts=36):
    """
    Creates an exact crescent arc where the overlapping ends terminate inside the adjacent card
    with zero exposed flat end-caps or steps on the horizontal midline.
    """
    theta_junc = math.radians(24.0)  # ~24 degrees cutoff inside card mask
    
    if is_top:
        start_angle = math.pi if is_first else (math.pi - theta_junc)
        end_angle = theta_junc if not is_last else 0
    else:
        start_angle = (math.pi + theta_junc) if not is_first else math.pi
        end_angle = (2 * math.pi - theta_junc) if not is_last else (2 * math.pi)
        
    pts = []
    for i in range(num_pts):
        theta = start_angle + (end_angle - start_angle) * (i / (num_pts - 1))
        px = cx + r_out * math.cos(theta)
        py = cy - r_out * math.sin(theta)
        pts.append((Inches(px), Inches(py)))
        
    for i in range(num_pts - 1, -1, -1):
        theta = start_angle + (end_angle - start_angle) * (i / (num_pts - 1))
        px = cx + r_in * math.cos(theta)
        py = cy - r_in * math.sin(theta)
        pts.append((Inches(px), Inches(py)))
        
    shd_pts = [(p[0] + Inches(0.03), p[1] + Inches(0.03)) for p in pts]
    fb_shd = slide.shapes.build_freeform(shd_pts[0][0], shd_pts[0][1])
    fb_shd.add_line_segments(shd_pts[1:], close=True)
    shd = fb_shd.convert_to_shape()
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
    shd.line.fill.background()
    
    fb = slide.shapes.build_freeform(pts[0][0], pts[0][1])
    fb.add_line_segments(pts[1:], close=True)
    shape = fb.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.fill.background()
    return shape

def render_wave_timeline(slide, title: str, subtitle: str, phases: List[Dict[str, Any]], 
                         cy=4.75, r_out=1.55, r_in=1.34, r_card=1.35):
    """
    Renders an exact seamless overlapping serpentine wave timeline slide.
    Zero exposed flat endcaps on horizontal midline.
    """
    add_slide_header(slide, title, subtitle_text=subtitle, x=0.80, y=1.00, w=11.733, title_size=32)
    
    n = len(phases)
    for i, p in enumerate(phases):
        cx = p['cx']
        c_hex = p['color']
        c_rgb = hex_to_rgb(c_hex)
        arc_top = p.get('arc_top', True)
        badge_num = p.get('badge', f"{i+1:02d}")
        is_first = (i == 0)
        is_last = (i == n - 1)
        
        # 1. Cluster Arc (Masked overlap)
        create_smooth_wave_arc(slide, cx, cy, r_in, r_out, is_top=arc_top, 
                               is_first=is_first, is_last=is_last, color=c_hex)
        
        # 2. Cluster Badge Tab Disc (Behind white card)
        badge_size = 0.90
        badge_cx = cx
        badge_cy = (cy - r_card) if not arc_top else (cy + r_card)
        
        add_soft_shadow(slide, MSO_SHAPE.OVAL, badge_cx - badge_size/2.0, badge_cy - badge_size/2.0, badge_size, badge_size, 0.03, 0.03)
        bd = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_cx - badge_size/2.0), Inches(badge_cy - badge_size/2.0), Inches(badge_size), Inches(badge_size))
        bd.fill.solid()
        bd.fill.fore_color.rgb = c_rgb
        bd.line.fill.background()
        
        # 3. Cluster White Floating Circle Card (Overlaps previous cluster)
        add_soft_shadow(slide, MSO_SHAPE.OVAL, cx - r_card, cy - r_card, r_card*2, r_card*2, 0.035, 0.035)
        c_in = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - r_card), Inches(cy - r_card), Inches(r_card*2), Inches(r_card*2))
        c_in.fill.solid()
        c_in.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
        c_in.line.fill.background()
        
        # 4. Badge Number (Foreground)
        tb_badge_y = (badge_cy - 0.32) if not arc_top else (badge_cy + 0.04)
        tb_badge = slide.shapes.add_textbox(Inches(badge_cx - 0.40), Inches(tb_badge_y), Inches(0.80), Inches(0.35))
        p_num = tb_badge.text_frame.paragraphs[0]
        p_num.text = badge_num
        p_num.font.name = Fonts.TITLE
        p_num.font.size = Pt(13)
        p_num.font.bold = True
        p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_num.alignment = PP_ALIGN.CENTER
        
        # 5. Card Text (Foreground)
        tb_w = 2.15
        tb_h = 1.95
        y_offset = -0.04 if not arc_top else 0.04
        tb = slide.shapes.add_textbox(Inches(cx - tb_w/2.0), Inches(cy - tb_h/2.0 + y_offset), Inches(tb_w), Inches(tb_h))
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
        
        for b_text in p.get('bullets', []):
            p_bullet = tf.add_paragraph()
            p_bullet.text = f"•  {b_text}"
            p_bullet.font.name = Fonts.BODY
            p_bullet.font.size = Pt(8.8)
            p_bullet.font.color.rgb = hex_to_rgb('#4A5568')
            p_bullet.alignment = PP_ALIGN.LEFT
            p_bullet.space_after = Pt(2.5)

import sys
import os
import math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from ppt_engine.theme import hex_to_rgb, Colors, Fonts, Sizes
from ppt_engine.geometry import add_soft_shadow
from ppt_engine.qa import VisionQA

# Initialize Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.500)

def add_header(slide, title_text='Training Phases', align=PP_ALIGN.CENTER, x=0.8, y=0.45, w=11.733):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = 'Montserrat'
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#111827')
    p.alignment = align

# Color definitions
C_BLUE = '#0D73F2'
C_BLUE_DARK = '#074FA8'
C_PURPLE = '#584DD8'
C_PURPLE_DARK = '#3C32A0'
C_ORANGE = '#F37E39'
C_ORANGE_DARK = '#B55417'
C_RED = '#E0403D'
C_RED_DARK = '#A82220'
C_GREEN = '#0FBE8E'
C_GREEN_DARK = '#098261'
C_CYAN = '#1FB6B7'
C_CYAN_DARK = '#127A7B'

# ==============================================================================
# SLIDE 1: 3-PHASE MATRIX PROCESS (PRE-TRAINING, TRAINING, POST TRAINING)
# ==============================================================================
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s1, 'Training Phases')

s1_sections = [
    {
        'header': 'Pre-Training Phase',
        'y_hdr': 1.35,
        'y_tiles': 1.75,
        'tiles': [
            ('Scanning market and\nanalyse potential', C_BLUE, 1),
            ('Selection of\ncourses', C_PURPLE, 1),
            ('Curriculum\ndevelopment', C_ORANGE, 1),
            ('Selecting and\ntraining of faculties', C_RED, 1),
            ('Community\nmobilization drive', C_GREEN, 1),
            ('Identification and\nselection of students', C_CYAN, 1),
        ]
    },
    {
        'header': 'Training Phase',
        'y_hdr': 3.20,
        'y_tiles': 3.60,
        'tiles': [
            ('Theory & practical\ntraining sessions', C_BLUE, 1),
            ('Like skill and personality\ndevelopment', C_PURPLE, 1),
            ('lectures', C_ORANGE, 1),
            ('Inputs from\nemployers', C_RED, 1),
            ('Performance\nappraisal of students', C_GREEN, 1),
            ('On the job training\nand internship', C_CYAN, 1),
        ]
    },
    {
        'header': 'Post Training Phase',
        'y_hdr': 5.05,
        'y_tiles': 5.45,
        'tiles': [
            ('Interviews and\nplacement support', C_BLUE, 1),
            ('Post-placement\nfollow up', C_PURPLE, 1),
            ('Employer feedback\nfor improvement', C_ORANGE, 1),
            ('Formation of alumni\nassociations', C_RED, 1),
            ('Experience\nsharing', C_GREEN, 2),
        ]
    }
]

grid_x = 0.60
grid_w = 12.133
unit_w = grid_w / 6.0

for sec in s1_sections:
    # 1. Header Bar Shadow & Shape
    add_soft_shadow(s1, MSO_SHAPE.RECTANGLE, grid_x, sec['y_hdr'], grid_w, 0.40, 0.02, 0.02, '#E2E8F0')
    hdr_box = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(grid_x), Inches(sec['y_hdr']), Inches(grid_w), Inches(0.40))
    hdr_box.fill.solid()
    hdr_box.fill.fore_color.rgb = hex_to_rgb('#F1F5F9')
    hdr_box.line.fill.background()
    
    tf_h = hdr_box.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = sec['header']
    p_h.font.name = 'Montserrat'
    p_h.font.size = Pt(12)
    p_h.font.bold = True
    p_h.font.color.rgb = hex_to_rgb('#334155')
    p_h.alignment = PP_ALIGN.CENTER
    
    # 2. Tiles Row
    curr_x = grid_x
    for text, col, span in sec['tiles']:
        tile_w = unit_w * span
        
        # Soft shadow behind tile row
        add_soft_shadow(s1, MSO_SHAPE.RECTANGLE, curr_x, sec['y_tiles'], tile_w, 1.15, 0.02, 0.02, '#E2E8F0')
        
        tile = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(curr_x), Inches(sec['y_tiles']), Inches(tile_w), Inches(1.15))
        tile.fill.solid()
        tile.fill.fore_color.rgb = hex_to_rgb(col)
        tile.line.fill.background()
        
        tf_t = tile.text_frame
        tf_t.word_wrap = True
        tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_t.margin_left = Inches(0.08)
        tf_t.margin_right = Inches(0.08)
        tf_t.margin_top = Inches(0.05)
        tf_t.margin_bottom = Inches(0.05)
        
        p_t = tf_t.paragraphs[0]
        p_t.text = text
        p_t.font.name = 'Segoe UI'
        p_t.font.size = Pt(10.5)
        p_t.font.bold = True
        p_t.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_t.alignment = PP_ALIGN.CENTER
        
        curr_x += tile_w

# ==============================================================================
# SLIDE 2: 6-NODE CIRCULAR CYCLE PROCESS WITH LEFT TITLE
# ==============================================================================
s2 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Bold Title
tb_t2 = s2.shapes.add_textbox(Inches(1.00), Inches(2.90), Inches(3.80), Inches(2.00))
tf_t2 = tb_t2.text_frame
tf_t2.word_wrap = True
p2_1 = tf_t2.paragraphs[0]
p2_1.text = 'Training'
p2_1.font.name = 'Montserrat'
p2_1.font.size = Pt(40)
p2_1.font.bold = True
p2_1.font.color.rgb = hex_to_rgb('#111827')

p2_2 = tf_t2.add_paragraph()
p2_2.text = 'Phases'
p2_2.font.name = 'Montserrat'
p2_2.font.size = Pt(40)
p2_2.font.bold = True
p2_2.font.color.rgb = hex_to_rgb('#111827')

# Right Circular Diagram
cx, cy = 8.80, 3.75
R = 2.15

s2_nodes = [
    ('Build\nPhase', C_PURPLE, 90),     # Top (12 o'clock)
    ('Peak\nPhase', C_ORANGE, 30),     # Top-Right (2 o'clock)
    ('Taper\nPhase', C_RED, 330),       # Bottom-Right (4 o'clock)
    ('Race\nPhase', C_GREEN, 270),      # Bottom (6 o'clock)
    ('Recovery\nPhase', C_CYAN, 210),  # Bottom-Left (8 o'clock)
    ('Base\nPhase', C_BLUE, 150),      # Top-Left (10 o'clock)
]

# Connecting circular track / arc ring behind nodes
track_r = R
for i in range(len(s2_nodes)):
    n1 = s2_nodes[i]
    n2 = s2_nodes[(i + 1) % len(s2_nodes)]
    col = n1[1]
    
    # Generate fine arc points
    a_start = math.radians(n1[2])
    a_end = math.radians(n2[2])
    if a_start < a_end:
        a_end -= 2 * math.pi
        
    num_pts = 16
    pts = []
    for step in range(num_pts + 1):
        ang = a_start + (a_end - a_start) * (step / num_pts)
        px = cx + track_r * math.cos(ang)
        py = cy - track_r * math.sin(ang)
        pts.append((Inches(px), Inches(py)))
        
    fb_arc = s2.shapes.build_freeform(pts[0][0], pts[0][1])
    fb_arc.add_line_segments(pts[1:], close=False)
    arc_shape = fb_arc.convert_to_shape()
    arc_shape.fill.background()
    arc_shape.line.color.rgb = hex_to_rgb(col)
    arc_shape.line.width = Pt(4.5)

# Central White Disc / Hub
hub_d = 2.80
add_soft_shadow(s2, MSO_SHAPE.OVAL, cx - hub_d/2, cy - hub_d/2, hub_d, hub_d, 0.04, 0.04, '#CBD5E1')

hub_outer = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - hub_d/2), Inches(cy - hub_d/2), Inches(hub_d), Inches(hub_d))
hub_outer.fill.solid()
hub_outer.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
hub_outer.line.color.rgb = hex_to_rgb('#E2E8F0')
hub_outer.line.width = Pt(1.5)

hub_in_d = 2.45
hub_in = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - hub_in_d/2), Inches(cy - hub_in_d/2), Inches(hub_in_d), Inches(hub_in_d))
hub_in.fill.solid()
hub_in.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
hub_in.line.fill.background()

tf_hub = hub_in.text_frame
tf_hub.vertical_anchor = MSO_ANCHOR.MIDDLE
p_hub = tf_hub.paragraphs[0]
p_hub.text = 'Training\nPhases'
p_hub.font.name = 'Montserrat'
p_hub.font.size = Pt(18)
p_hub.font.bold = True
p_hub.font.color.rgb = hex_to_rgb('#2D3748')
p_hub.alignment = PP_ALIGN.CENTER

# Draw 6 Circular Nodes
node_d = 1.45
for label, col, deg in s2_nodes:
    rad = math.radians(deg)
    nx = cx + R * math.cos(rad)
    ny = cy - R * math.sin(rad)
    
    # Soft shadow
    add_soft_shadow(s2, MSO_SHAPE.OVAL, nx - node_d/2, ny - node_d/2, node_d, node_d, 0.03, 0.03, '#94A3B8')
    
    # Circle Node
    node = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(nx - node_d/2), Inches(ny - node_d/2), Inches(node_d), Inches(node_d))
    node.fill.solid()
    node.fill.fore_color.rgb = hex_to_rgb(col)
    node.line.fill.background()
    
    tf_n = node.text_frame
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_n.margin_left = tf_n.margin_right = tf_n.margin_top = tf_n.margin_bottom = 0
    p_n = tf_n.paragraphs[0]
    p_n.text = label
    p_n.font.name = 'Montserrat'
    p_n.font.size = Pt(12)
    p_n.font.bold = True
    p_n.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_n.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 3: 5 3D PODIUMS WITH FLOATING BADGES AND ARROWS
# ==============================================================================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s3, 'Training Phases')

s3_data = [
    ('Analyze', 'Examine needs, gather data,\nidentify goals, and define\nobjectives clearly.', C_BLUE, C_BLUE_DARK, 'analyze'),
    ('Design', 'Create detailed plans,\nstructure content, develop\ntools, and instructional\nstrategy.', C_PURPLE, C_PURPLE_DARK, 'design'),
    ('Development', 'Build content, produce\nmaterials, integrate media,\ntest usability, ensure\nreadiness.', C_ORANGE, C_ORANGE_DARK, 'development'),
    ('Implementation', 'Deliver content, engage\nlearners, monitor process,\nsupport users, apply logistics.', C_RED, C_RED_DARK, 'implementation'),
    ('Evaluation', 'Assess outcomes, gather\nfeedback, analyze\nperformance, refine strategy,\nimprove effectiveness.', C_GREEN, C_GREEN_DARK, 'evaluation'),
]

xc_s3 = [1.70, 4.15, 6.60, 9.05, 11.50]
pw = 2.05
ph_base = 0.42
ph_top = 0.40
podium_y = 3.20

for i, (title, desc, col, col_dark, icon_key) in enumerate(s3_data):
    x_mid = xc_s3[i]
    x_left = x_mid - pw/2
    
    # 1. 3D Podium Base Cylinder
    # Shadow
    add_soft_shadow(s3, MSO_SHAPE.OVAL, x_left, podium_y + 0.15, pw, ph_top, 0.04, 0.04, '#CBD5E1')
    
    # Darker Lower Cylinder Body
    body = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_left), Inches(podium_y + ph_top/2), Inches(pw), Inches(ph_base))
    body.fill.solid()
    body.fill.fore_color.rgb = hex_to_rgb(col_dark)
    body.line.fill.background()
    
    # Top Vibrant Ellipse Face
    top_el = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_left), Inches(podium_y), Inches(pw), Inches(ph_top))
    top_el.fill.solid()
    top_el.fill.fore_color.rgb = hex_to_rgb(col)
    top_el.line.fill.background()
    
    # Center Dark Depression Dot
    dot = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - 0.22), Inches(podium_y + ph_top/2 - 0.05), Inches(0.44), Inches(0.12))
    dot.fill.solid()
    dot.fill.fore_color.rgb = hex_to_rgb(col_dark)
    dot.line.fill.background()
    
    # 2. Floating Pin Badge
    badge_yc = 1.95
    badge_d = 1.15
    
    # Teardrop Pointer
    ptr_pts = [
        (Inches(x_mid - 0.22), Inches(badge_yc + 0.35)),
        (Inches(x_mid + 0.22), Inches(badge_yc + 0.35)),
        (Inches(x_mid), Inches(badge_yc + 0.72)),
    ]
    fb_ptr = s3.shapes.build_freeform(ptr_pts[0][0], ptr_pts[0][1])
    fb_ptr.add_line_segments(ptr_pts[1:], close=True)
    ptr_shape = fb_ptr.convert_to_shape()
    ptr_shape.fill.solid()
    ptr_shape.fill.fore_color.rgb = hex_to_rgb(col)
    ptr_shape.line.fill.background()
    
    # Outer White Glow / Rim
    add_soft_shadow(s3, MSO_SHAPE.OVAL, x_mid - (badge_d+0.12)/2, badge_yc - (badge_d+0.12)/2, badge_d+0.12, badge_d+0.12, 0.03, 0.03, '#94A3B8')
    outer_b = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - (badge_d+0.12)/2), Inches(badge_yc - (badge_d+0.12)/2), Inches(badge_d+0.12), Inches(badge_d+0.12))
    outer_b.fill.solid()
    outer_b.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    outer_b.line.color.rgb = hex_to_rgb('#E2E8F0')
    outer_b.line.width = Pt(1.5)
    
    # Inner Colored Disc
    inner_b = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - badge_d/2), Inches(badge_yc - badge_d/2), Inches(badge_d), Inches(badge_d))
    inner_b.fill.solid()
    inner_b.fill.fore_color.rgb = hex_to_rgb(col)
    inner_b.line.fill.background()
    
    # Icon inside Badge
    icon_p = f'output/icons_12m/icon_white_{icon_key}.png'
    if os.path.exists(icon_p):
        s3.shapes.add_picture(icon_p, Inches(x_mid - 0.32), Inches(badge_yc - 0.32), width=Inches(0.64), height=Inches(0.64))
        
    # 3. Connecting Arrow to next step
    if i < len(s3_data) - 1:
        arr_x = (x_mid + xc_s3[i+1]) / 2.0 - 0.18
        arr = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(arr_x), Inches(podium_y + 0.18), Inches(0.36), Inches(0.18))
        arr.fill.solid()
        arr.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
        arr.line.fill.background()
        
    # 4. Text Below Podium
    tb_t = s3.shapes.add_textbox(Inches(x_mid - 1.10), Inches(4.15), Inches(2.20), Inches(0.40))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p_title = tf_t.paragraphs[0]
    p_title.text = title
    p_title.font.name = 'Montserrat'
    p_title.font.size = Pt(14)
    p_title.font.bold = True
    p_title.font.color.rgb = hex_to_rgb('#111827')
    p_title.alignment = PP_ALIGN.CENTER
    
    tb_d = s3.shapes.add_textbox(Inches(x_mid - 1.10), Inches(4.55), Inches(2.20), Inches(1.60))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    p_desc = tf_d.paragraphs[0]
    p_desc.text = desc
    p_desc.font.name = 'Segoe UI'
    p_desc.font.size = Pt(9.5)
    p_desc.font.color.rgb = hex_to_rgb('#4A5568')
    p_desc.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 4: 5 OVERLAPPING CHEVRON DONUT RINGS
# ==============================================================================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s4, 'Training Phases')

s4_data = [
    ('Analyze', 'Examine needs, gather\ndata, identify goals, and\ndefine objectives clearly.', C_BLUE),
    ('Design', 'Create detailed plans,\nstructure content,\ndevelop tools, and\ninstructional strategy.', C_PURPLE),
    ('Development', 'Build content, produce\nmaterials, integrate\nmedia, test usability,\nensure readiness.', C_ORANGE),
    ('Implementation', 'Deliver content, engage\nlearners, monitor\nprocess, support users,\napply logistics.', C_RED),
    ('Evaluation', 'Assess outcomes, gather\nfeedback, analyze\nperformance, refine\nstrategy, improve\neffectiveness.', C_GREEN),
]

xc_s4 = [1.80, 4.05, 6.30, 8.55, 10.80]
yc_s4 = 2.65
ring_d = 2.60

for i, (title, desc, col) in enumerate(s4_data):
    x_mid = xc_s4[i]
    
    # 1. Outer Ring (Donut shape)
    ring = s4.shapes.add_shape(MSO_SHAPE.DONUT, Inches(x_mid - ring_d/2), Inches(yc_s4 - ring_d/2), Inches(ring_d), Inches(ring_d))
    ring.fill.solid()
    ring.fill.fore_color.rgb = hex_to_rgb(col)
    ring.line.fill.background()
    try:
        ring.adjustments[0] = 0.28
    except Exception:
        pass
        
    # Inner White Hole disc to keep hole pure white
    hole_d = ring_d * (1.0 - 0.28 * 2)
    hole = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - hole_d/2), Inches(yc_s4 - hole_d/2), Inches(hole_d), Inches(hole_d))
    hole.fill.solid()
    hole.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    hole.line.fill.background()
        
    # 2. Horizontal Center Chevron Banner
    ban_w = 2.05
    ban_h = 0.54
    ban_x = x_mid - ban_w/2
    ban_y = yc_s4 - ban_h/2
    
    chv_pts = [
        (Inches(ban_x), Inches(ban_y)),
        (Inches(ban_x + ban_w - 0.22), Inches(ban_y)),
        (Inches(ban_x + ban_w), Inches(yc_s4)),
        (Inches(ban_x + ban_w - 0.22), Inches(ban_y + ban_h)),
        (Inches(ban_x), Inches(ban_y + ban_h)),
        (Inches(ban_x + 0.22), Inches(yc_s4)),
    ] if i > 0 else [
        (Inches(ban_x), Inches(ban_y)),
        (Inches(ban_x + ban_w - 0.22), Inches(ban_y)),
        (Inches(ban_x + ban_w), Inches(yc_s4)),
        (Inches(ban_x + ban_w - 0.22), Inches(ban_y + ban_h)),
        (Inches(ban_x), Inches(ban_y + ban_h)),
    ]
    
    fb_ban = s4.shapes.build_freeform(chv_pts[0][0], chv_pts[0][1])
    fb_ban.add_line_segments(chv_pts[1:], close=True)
    ban = fb_ban.convert_to_shape()
    ban.fill.solid()
    ban.fill.fore_color.rgb = hex_to_rgb(col)
    ban.line.fill.background()
    
    tf_b = ban.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_b.paragraphs[0]
    p_b.text = title
    p_b.font.name = 'Montserrat'
    p_b.font.size = Pt(11.5)
    p_b.font.bold = True
    p_b.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_b.alignment = PP_ALIGN.CENTER
    
    # 3. White Chevron Arrow notch on right edge
    notch_pts = [
        (Inches(x_mid + ring_d/2 - 0.12), Inches(yc_s4 - 0.22)),
        (Inches(x_mid + ring_d/2 + 0.14), Inches(yc_s4)),
        (Inches(x_mid + ring_d/2 - 0.12), Inches(yc_s4 + 0.22)),
        (Inches(x_mid + ring_d/2 - 0.04), Inches(yc_s4)),
    ]
    fb_notch = s4.shapes.build_freeform(notch_pts[0][0], notch_pts[0][1])
    fb_notch.add_line_segments(notch_pts[1:], close=True)
    notch = fb_notch.convert_to_shape()
    notch.fill.solid()
    notch.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    notch.line.fill.background()
    
    # 4. Text Below Ring
    tb_d4 = s4.shapes.add_textbox(Inches(x_mid - 1.05), Inches(4.45), Inches(2.10), Inches(1.80))
    tf_d4 = tb_d4.text_frame
    tf_d4.word_wrap = True
    p_d4 = tf_d4.paragraphs[0]
    p_d4.text = desc
    p_d4.font.name = 'Segoe UI'
    p_d4.font.size = Pt(9.5)
    p_d4.font.color.rgb = hex_to_rgb('#4A5568')
    p_d4.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 5: LEFT HUB CALENDAR BADGE WITH MULTI-LEVEL ROWS
# ==============================================================================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s5, 'Training Phases')

s5_rows = [
    ('Base Phase', 'Establish aerobic endurance, build strength,\ndevelop technique, focus on consistency,', C_BLUE, C_BLUE_DARK, '1', 1.65, 2.75, 1.80),
    ('Peak Phase', 'Maximize performance, fine-tune skills, reduce volume,\nsharpen mental focus.', C_PURPLE, C_PURPLE_DARK, '2', 2.55, 3.35, 2.70),
    ('Race Phase', 'Execute strategy, perform at peak, manage effort, maintain\ncompetitive mindset.', C_ORANGE, C_ORANGE_DARK, '3', 3.45, 3.65, 3.60),
    ('Build Phase', 'Increase intensity, add intervals, improve speed, build\nmuscular and cardiovascular power.', C_RED, C_RED_DARK, '4', 4.35, 3.35, 4.50),
    ('Taper Phase', 'Reduce training load, enhance recovery, maintain intensity,\nprepare for competition.', C_GREEN, C_GREEN_DARK, '5', 5.25, 2.75, 5.40),
]

# Left Hub Center
hub_xc, hub_yc = 1.70, 3.90
hub_d5 = 2.45

# Outer soft shadow
add_soft_shadow(s5, MSO_SHAPE.OVAL, hub_xc - hub_d5/2, hub_yc - hub_d5/2, hub_d5, hub_d5, 0.04, 0.04, '#CBD5E1')

# Left White Circle Hub
hub5 = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(hub_xc - hub_d5/2), Inches(hub_yc - hub_d5/2), Inches(hub_d5), Inches(hub_d5))
hub5.fill.solid()
hub5.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
hub5.line.color.rgb = hex_to_rgb('#E2E8F0')
hub5.line.width = Pt(1.5)

# Place Calendar Icon in Hub
cal_icon = 'output/icons_12m/icon_calendar_clock.png'
if os.path.exists(cal_icon):
    s5.shapes.add_picture(cal_icon, Inches(hub_xc - 0.60), Inches(hub_yc - 0.60), width=Inches(1.20), height=Inches(1.20))

# Multi-colored perimeter arc segments along right side of hub
hub_r = hub_d5 / 2.0
arc_angles = [(60, 24, C_BLUE), (24, -12, C_PURPLE), (-12, -48, C_ORANGE), (-48, -84, C_RED), (-84, -120, C_GREEN)]
for a1, a2, col in arc_angles:
    rad1, rad2 = math.radians(a1), math.radians(a2)
    pts_arc = []
    for step in range(8 + 1):
        ang = rad1 + (rad2 - rad1) * (step / 8.0)
        px = hub_xc + (hub_r + 0.06) * math.cos(ang)
        py = hub_yc - (hub_r + 0.06) * math.sin(ang)
        pts_arc.append((Inches(px), Inches(py)))
    fb_a = s5.shapes.build_freeform(pts_arc[0][0], pts_arc[0][1])
    fb_a.add_line_segments(pts_arc[1:], close=False)
    arc_s = fb_a.convert_to_shape()
    arc_s.fill.background()
    arc_s.line.color.rgb = hex_to_rgb(col)
    arc_s.line.width = Pt(4.5)

# Curved branching lines linking hub -> badges
for title, desc, col, col_dark, num, row_y, bx, by in s5_rows:
    # Connector line from hub perimeter to badge
    line_pts = [
        (Inches(hub_xc + hub_r * math.cos(math.radians(90 - (by - 1.5) * 35))), Inches(hub_yc - hub_r * math.sin(math.radians(90 - (by - 1.5) * 35)))),
        (Inches(bx - 0.20), Inches(by)),
        (Inches(4.40), Inches(row_y + 0.35))
    ]
    fb_l = s5.shapes.build_freeform(line_pts[0][0], line_pts[0][1])
    fb_l.add_line_segments(line_pts[1:], close=False)
    conn = fb_l.convert_to_shape()
    conn.fill.background()
    conn.line.color.rgb = hex_to_rgb('#CBD5E1')
    conn.line.width = Pt(1.5)
    
    # Numbered circular badge
    bd_size = 0.38
    add_soft_shadow(s5, MSO_SHAPE.OVAL, bx - bd_size/2, by - bd_size/2, bd_size, bd_size, 0.02, 0.02, '#CBD5E1')
    badge = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(bx - bd_size/2), Inches(by - bd_size/2), Inches(bd_size), Inches(bd_size))
    badge.fill.solid()
    badge.fill.fore_color.rgb = hex_to_rgb(col)
    badge.line.fill.background()
    
    tf_bd = badge.text_frame
    tf_bd.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_bd = tf_bd.paragraphs[0]
    p_bd.text = num
    p_bd.font.name = 'Montserrat'
    p_bd.font.size = Pt(10)
    p_bd.font.bold = True
    p_bd.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_bd.alignment = PP_ALIGN.CENTER
    
    # Right Row Pill Container
    pill_w = 3.25
    pill_h = 0.70
    pill_x = 4.40
    
    add_soft_shadow(s5, MSO_SHAPE.ROUNDED_RECTANGLE, pill_x, row_y, pill_w, pill_h, 0.02, 0.02, '#E2E8F0')
    pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(pill_x), Inches(row_y), Inches(pill_w), Inches(pill_h))
    pill.fill.solid()
    pill.fill.fore_color.rgb = hex_to_rgb('#F4F5F8')
    pill.line.color.rgb = hex_to_rgb('#E2E8F0')
    pill.line.width = Pt(1.0)
    
    # 3D Button inside Pill (Left)
    btn_d = 0.52
    btn_x = pill_x + 0.12
    btn_y = row_y + (pill_h - btn_d)/2
    
    # Darker base cylinder for 3D depth
    add_soft_shadow(s5, MSO_SHAPE.OVAL, btn_x, btn_y, btn_d, btn_d, 0.02, 0.02, '#94A3B8')
    btn_base = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(btn_x), Inches(btn_y + 0.05), Inches(btn_d), Inches(btn_d))
    btn_base.fill.solid()
    btn_base.fill.fore_color.rgb = hex_to_rgb(col_dark)
    btn_base.line.fill.background()
    
    btn = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(btn_x), Inches(btn_y), Inches(btn_d), Inches(btn_d))
    btn.fill.solid()
    btn.fill.fore_color.rgb = hex_to_rgb(col)
    btn.line.fill.background()
    
    # Title inside Pill (Right of Button)
    tb_pt = s5.shapes.add_textbox(Inches(pill_x + 0.75), Inches(row_y + 0.12), Inches(pill_w - 0.85), Inches(0.45))
    tf_pt = tb_pt.text_frame
    tf_pt.word_wrap = True
    p_pt = tf_pt.paragraphs[0]
    p_pt.text = title
    p_pt.font.name = 'Montserrat'
    p_pt.font.size = Pt(13)
    p_pt.font.bold = True
    p_pt.font.color.rgb = hex_to_rgb('#111827')
    
    # Description to the right of Pill
    tb_pd = s5.shapes.add_textbox(Inches(7.85), Inches(row_y + 0.08), Inches(4.90), Inches(0.55))
    tf_pd = tb_pd.text_frame
    tf_pd.word_wrap = True
    p_pd = tf_pd.paragraphs[0]
    p_pd.text = desc
    p_pd.font.name = 'Segoe UI'
    p_pd.font.size = Pt(9.5)
    p_pd.font.color.rgb = hex_to_rgb('#4A5568')

# ==============================================================================
# SLIDE 6: 6 INTERCONNECTED 3D CHEVRONS WITH ALTERNATING CALLOUTS
# ==============================================================================
s6 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s6, 'Training Phases')

s6_data = [
    ('Base Phase', 'Establish aerobic endurance, build\nstrength, develop technique, focus\non consistency,', C_BLUE, C_BLUE_DARK, '01', True),
    ('Build Phase', 'Increase intensity, add intervals,\nimprove speed, build muscular and\ncardiovascular power.', C_PURPLE, C_PURPLE_DARK, '02', False),
    ('Peak Phase', 'Maximize performance, fine-tune\nskills, reduce volume, sharpen\nmental focus.', C_ORANGE, C_ORANGE_DARK, '03', True),
    ('Taper Phase', 'Reduce training load, enhance\nrecovery, maintain intensity, prepare\nfor competition.', C_RED, C_RED_DARK, '04', False),
    ('Race Phase', 'Execute strategy, perform at peak,\nmanage effort, maintain\ncompetitive mindset.', C_GREEN, C_GREEN_DARK, '05', True),
    ('Recovery Phase', 'Rest body, repair muscles, restore\nenergy, reflect, plan next cycle.', C_CYAN, C_CYAN_DARK, '06', False),
]

# Continuous interlocking chevrons
xc_s6 = [1.60, 3.55, 5.50, 7.45, 9.40, 11.35]
cw_s6 = 2.25
ch_s6 = 0.85
cy_s6 = 3.45

for i, (title, desc, col, col_dark, num, is_top) in enumerate(s6_data):
    x_mid = xc_s6[i]
    x_left = x_mid - cw_s6/2
    
    # 1. 3D Chevron Block
    # Extruded bottom 3D shadow face
    chv_base = s6.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x_left), Inches(cy_s6 + 0.12), Inches(cw_s6), Inches(ch_s6))
    chv_base.fill.solid()
    chv_base.fill.fore_color.rgb = hex_to_rgb(col_dark)
    chv_base.line.fill.background()
    
    # Top Vibrant Chevron Face
    chv_top = s6.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x_left), Inches(cy_s6), Inches(cw_s6), Inches(ch_s6))
    chv_top.fill.solid()
    chv_top.fill.fore_color.rgb = hex_to_rgb(col)
    chv_top.line.fill.background()
    
    # 2. Pin Badge over Chevron (All 6 at top of chevron)
    pin_d = 0.95
    pin_yc = 2.65
    
    # Teardrop pointer pointing to chevron
    ptr_pts6 = [
        (Inches(x_mid - 0.18), Inches(pin_yc + 0.30)),
        (Inches(x_mid + 0.18), Inches(pin_yc + 0.30)),
        (Inches(x_mid), Inches(pin_yc + 0.62)),
    ]
    fb_ptr6 = s6.shapes.build_freeform(ptr_pts6[0][0], ptr_pts6[0][1])
    fb_ptr6.add_line_segments(ptr_pts6[1:], close=True)
    ptr6 = fb_ptr6.convert_to_shape()
    ptr6.fill.solid()
    ptr6.fill.fore_color.rgb = hex_to_rgb(col)
    ptr6.line.fill.background()
    
    # Outer White Glow / Rim
    add_soft_shadow(s6, MSO_SHAPE.OVAL, x_mid - (pin_d+0.10)/2, pin_yc - (pin_d+0.10)/2, pin_d+0.10, pin_d+0.10, 0.03, 0.03, '#94A3B8')
    outer_pin = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - (pin_d+0.10)/2), Inches(pin_yc - (pin_d+0.10)/2), Inches(pin_d+0.10), Inches(pin_d+0.10))
    outer_pin.fill.solid()
    outer_pin.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    outer_pin.line.color.rgb = hex_to_rgb('#E2E8F0')
    outer_pin.line.width = Pt(1.5)
    
    # Inner Solid Disc with Number
    inner_pin = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - pin_d/2), Inches(pin_yc - pin_d/2), Inches(pin_d), Inches(pin_d))
    inner_pin.fill.solid()
    inner_pin.fill.fore_color.rgb = hex_to_rgb(col)
    inner_pin.line.fill.background()
    
    tf_pin = inner_pin.text_frame
    tf_pin.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_pin = tf_pin.paragraphs[0]
    p_pin.text = num
    p_pin.font.name = 'Montserrat'
    p_pin.font.size = Pt(13)
    p_pin.font.bold = True
    p_pin.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_pin.alignment = PP_ALIGN.CENTER
    
    # 3. Alternating Stem Line & Text Block
    if is_top:
        # Line up to top callout
        stem = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_mid - 0.01), Inches(1.80), Inches(0.02), Inches(pin_yc - pin_d/2 - 1.80))
        stem.fill.solid()
        stem.fill.fore_color.rgb = hex_to_rgb('#A0AEC0')
        stem.line.fill.background()
        
        dot = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - 0.04), Inches(1.78), Inches(0.08), Inches(0.08))
        dot.fill.solid()
        dot.fill.fore_color.rgb = hex_to_rgb('#718096')
        dot.line.fill.background()
        
        # Text Block Above
        tb = s6.shapes.add_textbox(Inches(x_mid - 0.95), Inches(0.85), Inches(1.90), Inches(0.95))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = 'Montserrat'
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = hex_to_rgb('#111827')
        p1.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = hex_to_rgb('#4A5568')
        p2.alignment = PP_ALIGN.CENTER
    else:
        # Line down to bottom callout
        stem = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_mid - 0.01), Inches(cy_s6 + ch_s6 + 0.12), Inches(0.02), Inches(5.15 - (cy_s6 + ch_s6 + 0.12)))
        stem.fill.solid()
        stem.fill.fore_color.rgb = hex_to_rgb('#A0AEC0')
        stem.line.fill.background()
        
        dot = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - 0.04), Inches(5.15), Inches(0.08), Inches(0.08))
        dot.fill.solid()
        dot.fill.fore_color.rgb = hex_to_rgb('#718096')
        dot.line.fill.background()
        
        # Text Block Below
        tb = s6.shapes.add_textbox(Inches(x_mid - 0.95), Inches(5.25), Inches(1.90), Inches(0.95))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = 'Montserrat'
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = hex_to_rgb('#111827')
        p1.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = hex_to_rgb('#4A5568')
        p2.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 7: 5 OVERLAPPING COLORFUL CIRCLES WITH WHITE ICONS
# ==============================================================================
s7 = prs.slides.add_slide(prs.slide_layouts[6])
add_header(s7, 'Training Phases')

s7_data = [
    ('Analyze', 'Examine needs, gather\ndata, identify goals, and\ndefine objectives clearly.', C_BLUE, 'analyze'),
    ('Design', 'Create detailed plans,\nstructure content,\ndevelop tools, and\ninstructional strategy.', C_PURPLE, 'design'),
    ('Development', 'Build content, produce\nmaterials, integrate\nmedia, test usability,\nensure readiness.', C_ORANGE, 'development'),
    ('Implementation', 'Deliver content, engage\nlearners, monitor\nprocess, support users,\napply logistics.', C_RED, 'implementation'),
    ('Evaluation', 'Assess outcomes, gather\nfeedback, analyze\nperformance, refine\nstrategy, improve\neffectiveness.', C_GREEN, 'evaluation'),
]

xc_s7 = [1.85, 4.10, 6.35, 8.60, 10.85]
yc_s7 = 2.65
circle_d = 2.70

for i, (title, desc, col, icon_key) in enumerate(s7_data):
    x_mid = xc_s7[i]
    
    # 1. Soft Shadow behind circle
    add_soft_shadow(s7, MSO_SHAPE.OVAL, x_mid - circle_d/2, yc_s7 - circle_d/2, circle_d, circle_d, 0.04, 0.04, '#CBD5E1')
    
    # 2. Main Colored Circle
    circ = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_mid - circle_d/2), Inches(yc_s7 - circle_d/2), Inches(circle_d), Inches(circle_d))
    circ.fill.solid()
    circ.fill.fore_color.rgb = hex_to_rgb(col)
    circ.line.fill.background()
    
    # 3. White Icon inside Circle
    icon_p = f'output/icons_12m/icon_white_{icon_key}.png'
    if os.path.exists(icon_p):
        s7.shapes.add_picture(icon_p, Inches(x_mid - 0.38), Inches(yc_s7 - 0.70), width=Inches(0.76), height=Inches(0.76))
        
    # 4. Phase Title inside Circle
    tb_ct = s7.shapes.add_textbox(Inches(x_mid - 1.10), Inches(yc_s7 + 0.28), Inches(2.20), Inches(0.40))
    tf_ct = tb_ct.text_frame
    tf_ct.word_wrap = True
    p_ct = tf_ct.paragraphs[0]
    p_ct.text = title
    p_ct.font.name = 'Montserrat'
    p_ct.font.size = Pt(13)
    p_ct.font.bold = True
    p_ct.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_ct.alignment = PP_ALIGN.CENTER
    
    # 5. Text Below Circle
    tb_cd = s7.shapes.add_textbox(Inches(x_mid - 1.05), Inches(4.35), Inches(2.10), Inches(1.80))
    tf_cd = tb_cd.text_frame
    tf_cd.word_wrap = True
    p_cd = tf_cd.paragraphs[0]
    p_cd.text = desc
    p_cd.font.name = 'Segoe UI'
    p_cd.font.size = Pt(9.5)
    p_cd.font.color.rgb = hex_to_rgb('#4A5568')
    p_cd.alignment = PP_ALIGN.CENTER

# Save final presentation
target_dir = sys.argv[1] if len(sys.argv) > 1 else r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\12_Month_Training_Plan'
output_pptx = os.path.join(target_dir, 'presentation_95_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print(f'Successfully generated: {output_pptx}')

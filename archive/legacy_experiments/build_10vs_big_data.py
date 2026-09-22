import os
import sys
import math
import numpy as np
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from ppt_engine.theme import hex_to_rgb, Colors, Fonts, Sizes
from ppt_engine.geometry import add_soft_shadow

SOURCE_DIR = sys.argv[1] if len(sys.argv) > 1 else (
    r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10Vs_Of_Big_Data'
    if os.path.exists(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10Vs_Of_Big_Data')
    else r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\10Vs_Of_Big_Data'
)

output_assets = os.path.join(current_dir, 'output', 'assets_10vs')
os.makedirs(output_assets, exist_ok=True)
icons_dir = os.path.join(current_dir, 'output', 'icons_10vs')
os.makedirs(icons_dir, exist_ok=True)

# ==================== PALETTE DEFINITIONS ====================
C_VALUE = '#1E293B'      # Dark Navy / Charcoal
C_VIRALITY = '#EF4444'   # Vibrant Red
C_VENUE = '#F59E0B'      # Amber Orange
C_VARIABILITY = '#06B6D4'# Bright Cyan
C_VOLUME = '#84CC16'     # Fresh Green
C_VISCOSITY = '#A855F7'  # Vibrant Purple
C_VERACITY = '#1D4ED8'   # Royal Blue
C_VISUALIZATION = '#D97706' # Ochre / Brown
C_VELOCITY = '#0F766E'   # Deep Teal
C_VARIETY = '#991B1B'    # Dark Burgundy

ITEMS_DATA = [
    ('Value', 'Can you find the information you are looking for?', 'value', C_VALUE),
    ('Virality', "Portable 'Aha' message for presentation or Instagram.", 'virality', C_VIRALITY),
    ('Venue', 'Distributed Heterogeneous Data from multiple Platforms', 'venue', C_VENUE),
    ('Variability', 'Dynamic, Evolving Behavior in Data Source', 'variability', C_VARIABILITY),
    ('Volume', 'Can you find it when you most need it?', 'volume', C_VOLUME),
    ('Viscosity', 'Does it stick with you? Does it call for action?', 'viscosity', C_VISCOSITY),
    ('Veracity', 'Are you dealing with information or disinformation?', 'veracity', C_VERACITY),
    ('Visualization', 'Can you make sense at a glance? Does it trigger a decision?', 'visualization', C_VISUALIZATION),
    ('Velocity', "Info accelerates, crises and chances unfold. Today's outlook?", 'velocity', C_VELOCITY),
    ('Variety', 'Image impact, multilingual context, information equilibrium check.', 'variety', C_VARIETY),
]

ITEMS_DATA_ALT = [
    ('Volume', "Big data's vast, ever-growing volume.", 'volume', C_VOLUME),
    ('Velocity', 'The Speed at which Data is generated', 'velocity', C_VELOCITY),
    ('Variety', 'Diverse forms, extensive data types', 'variety', C_VARIETY),
    ('Veracity', "Veracity's Validity Check", 'veracity', C_VERACITY),
    ('Value', 'Insights, innovation, efficiency, growth', 'value', C_VISCOSITY),
    ('Validity Data', 'Quality, Governance, Master Data Management', 'value', C_VALUE),
    ('Variability', 'Dynamic, Evolving Behavior In Data Science', 'virality', C_VIRALITY),
    ('Venue', 'Distributed Heterogeneous Data from multiple Platforms', 'venue', C_VENUE),
    ('Vocabulary', 'Data models & semantics that describe data structure', 'variability', C_VARIABILITY),
    ('Vagueness', 'Confusion over meanings of Big data and Tools used', 'volume', C_VOLUME),
]

# ==================== EXTRACT CLEAN ICONS WITH CONNECTED COMPONENTS ====================
from scipy import ndimage

im1_src = Image.open(os.path.join(SOURCE_DIR, 'Slide_01.png')).convert('RGBA')

y_ranges = [
    (36, 74),   # Value
    (83, 121),  # Virality
    (130, 168), # Venue
    (177, 215), # Variability
    (224, 262), # Volume
    (271, 309), # Viscosity
    (318, 356), # Veracity
    (365, 403), # Visualization
    (412, 450), # Velocity
    (459, 497), # Variety
]
items_keys = ['value', 'virality', 'venue', 'variability', 'volume', 'viscosity', 'veracity', 'visualization', 'velocity', 'variety']

icon_paths_white = {}
icon_paths_dark = {}

for name, (y1, y2) in zip(items_keys, y_ranges):
    patch = im1_src.crop((445, y1, 485, y2))
    arr = np.array(patch, dtype=np.float32)
    
    r, g, b = arr[:,:,0], arr[:,:,1], arr[:,:,2]
    min_rgb = np.minimum(np.minimum(r, g), b)
    is_white = min_rgb > 140
    
    lbl, num = ndimage.label(is_white)
    clean_mask = np.zeros_like(is_white, dtype=bool)
    
    for c in range(1, num + 1):
        comp = (lbl == c)
        ys, xs = np.where(comp)
        w = xs.max() - xs.min() + 1
        h = ys.max() - ys.min() + 1
        if w > 24 and h <= 4:
            continue
        if ys.min() == 0 or ys.max() == patch.height - 1:
            continue
        if xs.min() == 0 or xs.max() == patch.width - 1:
            continue
        clean_mask |= comp
        
    alpha = np.where(clean_mask, np.clip((min_rgb - 100) / (220 - 100) * 255.0, 0, 255), 0).astype(np.uint8)
    
    # 1. White icon
    out_w = np.zeros((arr.shape[0], arr.shape[1], 4), dtype=np.uint8)
    out_w[:,:,:3] = 255
    out_w[:,:,3] = alpha
    p_w = os.path.join(icons_dir, f'icon_w_{name}.png')
    Image.fromarray(out_w).resize((140, 140), Image.Resampling.LANCZOS).save(p_w)
    icon_paths_white[name] = p_w
    
    # 2. Dark icon
    out_d = np.zeros((arr.shape[0], arr.shape[1], 4), dtype=np.uint8)
    out_d[:,:,:3] = 35
    out_d[:,:,3] = alpha
    p_d = os.path.join(icons_dir, f'icon_d_{name}.png')
    Image.fromarray(out_d).resize((140, 140), Image.Resampling.LANCZOS).save(p_d)
    icon_paths_dark[name] = p_d

# ==================== SLIDE 4 DONUT GRAPHIC ====================
scale_d = 4
w_d, h_d = 500 * scale_d, 350 * scale_d
donut_im = Image.new('RGBA', (w_d, h_d), (0, 0, 0, 0))
draw_d = ImageDraw.Draw(donut_im)
cx_d, cy_d = w_d // 2, h_d // 2
rx_out, ry_out = int(220 * scale_d), int(150 * scale_d)
rx_in, ry_in = int(140 * scale_d), int(95 * scale_d)

donut_colors_rgb = [
    (30, 41, 59), (239, 68, 68), (245, 158, 11), (6, 182, 212), (132, 204, 22),
    (168, 85, 247), (29, 78, 216), (217, 119, 6), (15, 118, 110), (153, 27, 27)
]

for i in range(10):
    start_a = i * 36
    end_a = start_a + 30
    col = donut_colors_rgb[i]
    for deg in np.linspace(start_a, end_a, 40):
        rad = math.radians(deg)
        x1 = cx_d + rx_in * math.cos(rad)
        y1 = cy_d + ry_in * math.sin(rad)
        x2 = cx_d + rx_out * math.cos(rad)
        y2 = cy_d + ry_out * math.sin(rad)
p_donut = os.path.join(output_assets, 's4_donut_clean.png')
donut_im.resize((w_d // scale_d, h_d // scale_d), Image.Resampling.LANCZOS).save(p_donut)

# ==================== SLIDE 7 DONUT GRAPHIC ====================
size_s7 = 600
donut_s7_im = Image.new('RGBA', (size_s7, size_s7), (255, 255, 255, 0))
draw_s7 = ImageDraw.Draw(donut_s7_im)
cx_7, cy_7 = size_s7 // 2, size_s7 // 2
r_out_7, r_in_7 = 270, 220

for i, col in enumerate(donut_colors_rgb):
    start_deg = -90 + i * 36
    end_deg = start_deg + 34
    draw_s7.pieslice([cx_7 - r_out_7, cy_7 - r_out_7, cx_7 + r_out_7, cy_7 + r_out_7], start_deg, end_deg, fill=col + (255,))
draw_s7.ellipse([cx_7 - r_in_7, cy_7 - r_in_7, cx_7 + r_in_7, cy_7 + r_in_7], fill=(255, 255, 255, 0))

p_donut_s7 = os.path.join(output_assets, 's7_donut_clean.png')
donut_s7_im.save(p_donut_s7)

# ==================== PRESENTATION BUILD ====================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

FONT_TITLE = 'Georgia'
FONT_BODY = 'Segoe UI'

def add_slide_title(slide, text="10Vs Of Big Data", align=PP_ALIGN.CENTER, x=1.0, y=0.45, w=11.333):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.70))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_TITLE
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb('#2D3748')
    p.alignment = align
    return tb

# ==============================================================================
# SLIDE 1: 10 STACKED ROUNDED RIBBON BARS
# ==============================================================================
s1 = prs.slides.add_slide(prs.slide_layouts[6])

# Left Title
tb_s1_t = s1.shapes.add_textbox(Inches(0.80), Inches(2.60), Inches(4.50), Inches(2.20))
tf_s1_t = tb_s1_t.text_frame
p1 = tf_s1_t.paragraphs[0]
p1.text = "10Vs Of"
p1.font.name = FONT_TITLE
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#2D3748')
p1.alignment = PP_ALIGN.CENTER

p2 = tf_s1_t.add_paragraph()
p2.text = "Big Data"
p2.font.name = FONT_TITLE
p2.font.size = Pt(44)
p2.font.bold = True
p2.font.color.rgb = hex_to_rgb('#2D3748')
p2.alignment = PP_ALIGN.CENTER

# Right 10 Stacked Bars
y_start_s1 = 0.50
bar_h_s1 = 0.56
gap_s1 = 0.12
x_bar_s1 = 6.00
w_bar_s1 = 6.70

for idx, (title, desc, key, col) in enumerate(ITEMS_DATA):
    y_pos = y_start_s1 + idx * (bar_h_s1 + gap_s1)
    col_rgb = hex_to_rgb(col)
    
    # 1. Main Background Rounded Pill
    add_soft_shadow(s1, MSO_SHAPE.ROUNDED_RECTANGLE, x_bar_s1, y_pos, w_bar_s1, bar_h_s1, 0.02, 0.02, '#E2E8F0')
    bar = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_bar_s1), Inches(y_pos), Inches(w_bar_s1), Inches(bar_h_s1))
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bar.line.color.rgb = col_rgb
    bar.line.width = Pt(1.5)
    
    # 2. Left Colored Ribbon Tab
    tab = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_bar_s1 - 0.45), Inches(y_pos), Inches(0.95), Inches(bar_h_s1))
    tab.fill.solid()
    tab.fill.fore_color.rgb = col_rgb
    tab.line.fill.background()
    
    # 3. White Icon
    if key in icon_paths_white:
        s1.shapes.add_picture(icon_paths_white[key], Inches(x_bar_s1 - 0.35), Inches(y_pos + 0.08), width=Inches(0.40), height=Inches(0.40))
        
    # 4. Text inside Bar
    tb_txt = s1.shapes.add_textbox(Inches(x_bar_s1 + 0.65), Inches(y_pos + 0.03), Inches(w_bar_s1 - 0.75), Inches(bar_h_s1 - 0.06))
    tf_txt = tb_txt.text_frame
    tf_txt.word_wrap = True
    tf_txt.margin_left = tf_txt.margin_top = tf_txt.margin_right = tf_txt.margin_bottom = 0
    
    pt = tf_txt.paragraphs[0]
    pt.text = title
    pt.font.name = FONT_BODY
    pt.font.size = Pt(11)
    pt.font.bold = True
    pt.font.color.rgb = hex_to_rgb('#1E293B')
    
    pd = tf_txt.add_paragraph()
    pd.text = desc
    pd.font.name = FONT_BODY
    pd.font.size = Pt(8.5)
    pd.font.color.rgb = hex_to_rgb('#64748B')

# ==============================================================================
# SLIDE 2: CENTER DIAMOND SPINE & ALTERNATING LEFT/RIGHT PILL WINGS
# ==============================================================================
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_title(s2, "10Vs Of Big Data")

y_start_s2 = 1.35
row_h_s2 = 0.54
gap_s2 = 0.06
cx_spine = 6.666
wing_w = 5.20

for idx, (title, desc, key, col) in enumerate(ITEMS_DATA):
    y_pos = y_start_s2 + idx * (row_h_s2 + gap_s2)
    col_rgb = hex_to_rgb(col)
    is_left = (idx % 2 == 0)
    
    if is_left:
        x_wing = cx_spine - wing_w - 0.15
        add_soft_shadow(s2, MSO_SHAPE.ROUNDED_RECTANGLE, x_wing, y_pos, wing_w, row_h_s2, 0.02, 0.02, '#E2E8F0')
        wing = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_wing), Inches(y_pos), Inches(wing_w), Inches(row_h_s2))
        wing.fill.solid()
        wing.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
        wing.line.color.rgb = col_rgb
        wing.line.width = Pt(1.5)
        
        tb_w = s2.shapes.add_textbox(Inches(x_wing + 0.20), Inches(y_pos + 0.03), Inches(wing_w - 0.60), Inches(row_h_s2 - 0.06))
        tf_w = tb_w.text_frame
        tf_w.word_wrap = True
        tf_w.margin_left = tf_w.margin_top = tf_w.margin_right = tf_w.margin_bottom = 0
        
        pt = tf_w.paragraphs[0]
        pt.text = title
        pt.font.name = FONT_BODY
        pt.font.size = Pt(11)
        pt.font.bold = True
        pt.font.color.rgb = hex_to_rgb('#1E293B')
        pt.alignment = PP_ALIGN.RIGHT
        
        pd = tf_w.add_paragraph()
        pd.text = desc
        pd.font.name = FONT_BODY
        pd.font.size = Pt(8.5)
        pd.font.color.rgb = hex_to_rgb('#64748B')
        pd.alignment = PP_ALIGN.RIGHT
    else:
        x_wing = cx_spine + 0.15
        add_soft_shadow(s2, MSO_SHAPE.ROUNDED_RECTANGLE, x_wing, y_pos, wing_w, row_h_s2, 0.02, 0.02, '#E2E8F0')
        wing = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_wing), Inches(y_pos), Inches(wing_w), Inches(row_h_s2))
        wing.fill.solid()
        wing.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
        wing.line.color.rgb = col_rgb
        wing.line.width = Pt(1.5)
        
        tb_w = s2.shapes.add_textbox(Inches(x_wing + 0.40), Inches(y_pos + 0.03), Inches(wing_w - 0.60), Inches(row_h_s2 - 0.06))
        tf_w = tb_w.text_frame
        tf_w.word_wrap = True
        tf_w.margin_left = tf_w.margin_top = tf_w.margin_right = tf_w.margin_bottom = 0
        
        pt = tf_w.paragraphs[0]
        pt.text = title
        pt.font.name = FONT_BODY
        pt.font.size = Pt(11)
        pt.font.bold = True
        pt.font.color.rgb = hex_to_rgb('#1E293B')
        pt.alignment = PP_ALIGN.LEFT
        
        pd = tf_w.add_paragraph()
        pd.text = desc
        pd.font.name = FONT_BODY
        pd.font.size = Pt(8.5)
        pd.font.color.rgb = hex_to_rgb('#64748B')
        pd.alignment = PP_ALIGN.LEFT
        
    # Center Diamond
    d_size = 0.52
    d_shape = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx_spine - d_size/2), Inches(y_pos + (row_h_s2 - d_size)/2), Inches(d_size), Inches(d_size))
    d_shape.rotation = 45
    d_shape.fill.solid()
    d_shape.fill.fore_color.rgb = col_rgb
    d_shape.line.color.rgb = hex_to_rgb('#FFFFFF')
    d_shape.line.width = Pt(1.5)
    
    if key in icon_paths_white:
        s2.shapes.add_picture(icon_paths_white[key], Inches(cx_spine - 0.16), Inches(y_pos + (row_h_s2 - 0.32)/2), width=Inches(0.32), height=Inches(0.32))

# ==============================================================================
# SLIDE 3: 10-SPOKE RADIAL HEXAGON NETWORK
# ==============================================================================
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_title(s3, "10Vs Of Big Data")

cx_s3, cy_s3 = 6.666, 4.30
hex_w, hex_h = 2.40, 2.60

# 10 Outer Nodes Coordinates
left_nodes = [
    (cx_s3 - 2.80, cy_s3 - 2.00, ITEMS_DATA_ALT[0]),
    (cx_s3 - 3.60, cy_s3 - 1.00, ITEMS_DATA_ALT[1]),
    (cx_s3 - 4.00, cy_s3 + 0.20, ITEMS_DATA_ALT[2]),
    (cx_s3 - 3.60, cy_s3 + 1.40, ITEMS_DATA_ALT[3]),
    (cx_s3 - 2.80, cy_s3 + 2.30, ITEMS_DATA_ALT[4]),
]

right_nodes = [
    (cx_s3 + 2.80, cy_s3 - 2.00, ITEMS_DATA_ALT[5]),
    (cx_s3 + 3.60, cy_s3 - 1.00, ITEMS_DATA_ALT[6]),
    (cx_s3 + 4.00, cy_s3 + 0.20, ITEMS_DATA_ALT[7]),
    (cx_s3 + 3.60, cy_s3 + 1.40, ITEMS_DATA_ALT[8]),
    (cx_s3 + 2.80, cy_s3 + 2.30, ITEMS_DATA_ALT[9]),
]

node_size = 0.82

# Draw Connectors first (behind nodes)
for nx, ny, (title, desc, key, col) in left_nodes:
    col_rgb = hex_to_rgb(col)
    conn = s3.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s3), Inches(cy_s3), Inches(nx + node_size/2), Inches(ny + node_size/2))
    conn.line.color.rgb = col_rgb
    conn.line.width = Pt(3.5)

for nx, ny, (title, desc, key, col) in right_nodes:
    col_rgb = hex_to_rgb(col)
    conn = s3.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s3), Inches(cy_s3), Inches(nx + node_size/2), Inches(ny + node_size/2))
    conn.line.color.rgb = col_rgb
    conn.line.width = Pt(3.5)

# Center Big Hexagon over connectors
add_soft_shadow(s3, MSO_SHAPE.HEXAGON, cx_s3 - hex_w/2, cy_s3 - hex_h/2, hex_w, hex_h, 0.04, 0.04, '#CBD5E1')
c_hex = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(cx_s3 - hex_w/2), Inches(cy_s3 - hex_h/2), Inches(hex_w), Inches(hex_h))
c_hex.fill.solid()
c_hex.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
c_hex.line.color.rgb = hex_to_rgb('#FFFFFF')
c_hex.line.width = Pt(2.5)

tf_ch = c_hex.text_frame
p_ch1 = tf_ch.paragraphs[0]
p_ch1.text = "10Vs Of"
p_ch1.font.name = FONT_TITLE
p_ch1.font.size = Pt(20)
p_ch1.font.bold = True
p_ch1.font.color.rgb = hex_to_rgb('#1E293B')
p_ch1.alignment = PP_ALIGN.CENTER
p_ch2 = tf_ch.add_paragraph()
p_ch2.text = "Big Data"
p_ch2.font.name = FONT_TITLE
p_ch2.font.size = Pt(20)
p_ch2.font.bold = True
p_ch2.font.color.rgb = hex_to_rgb('#1E293B')
p_ch2.alignment = PP_ALIGN.CENTER

# Outer Hexagons & Text
for nx, ny, (title, desc, key, col) in left_nodes:
    col_rgb = hex_to_rgb(col)
    add_soft_shadow(s3, MSO_SHAPE.HEXAGON, nx, ny, node_size, node_size, 0.02, 0.02, '#CBD5E1')
    hex_node = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(nx), Inches(ny), Inches(node_size), Inches(node_size))
    hex_node.fill.solid()
    hex_node.fill.fore_color.rgb = col_rgb
    hex_node.line.fill.background()
    
    if key in icon_paths_white:
        s3.shapes.add_picture(icon_paths_white[key], Inches(nx + 0.20), Inches(ny + 0.20), width=Inches(0.42), height=Inches(0.42))
        
    tb_l = s3.shapes.add_textbox(Inches(nx - 2.80), Inches(ny - 0.10), Inches(2.65), Inches(0.95))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p1 = tf_l.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf_l.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.RIGHT

for nx, ny, (title, desc, key, col) in right_nodes:
    col_rgb = hex_to_rgb(col)
    add_soft_shadow(s3, MSO_SHAPE.HEXAGON, nx, ny, node_size, node_size, 0.02, 0.02, '#CBD5E1')
    hex_node = s3.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(nx), Inches(ny), Inches(node_size), Inches(node_size))
    hex_node.fill.solid()
    hex_node.fill.fore_color.rgb = col_rgb
    hex_node.line.fill.background()
    
    if key in icon_paths_white:
        s3.shapes.add_picture(icon_paths_white[key], Inches(nx + 0.20), Inches(ny + 0.20), width=Inches(0.42), height=Inches(0.42))
        
    tb_r = s3.shapes.add_textbox(Inches(nx + node_size + 0.15), Inches(ny - 0.10), Inches(2.65), Inches(0.95))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p1 = tf_r.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf_r.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.LEFT

# ==============================================================================
# SLIDE 4: 10-SEGMENT DONUT COLOR WHEEL + 5 LEFT / 5 RIGHT PILL CARDS
# ==============================================================================
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_title(s4, "10Vs Of Big Data")

cx_s4, cy_s4 = 6.666, 4.30
donut_w = 3.60
s4.shapes.add_picture(p_donut, Inches(cx_s4 - donut_w/2), Inches(cy_s4 - 1.25), width=Inches(donut_w))

c_hole = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx_s4 - 1.05), Inches(cy_s4 - 0.70), Inches(2.10), Inches(1.40))
c_hole.fill.solid()
c_hole.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
c_hole.line.fill.background()

tf_h = c_hole.text_frame
p_h1 = tf_h.paragraphs[0]
p_h1.text = "10Vs Of"
p_h1.font.name = FONT_TITLE
p_h1.font.size = Pt(18)
p_h1.font.bold = True
p_h1.font.color.rgb = hex_to_rgb('#1E293B')
p_h1.alignment = PP_ALIGN.CENTER
p_h2 = tf_h.add_paragraph()
p_h2.text = "Big Data"
p_h2.font.name = FONT_TITLE
p_h2.font.size = Pt(18)
p_h2.font.bold = True
p_h2.font.color.rgb = hex_to_rgb('#1E293B')
p_h2.alignment = PP_ALIGN.CENTER

y_start_s4 = 1.60
gap_s4 = 1.05
card_w_s4 = 3.90
card_h_s4 = 0.75

# Left Cards
for i in range(5):
    title, desc, key, col = ITEMS_DATA_ALT[i]
    y_pos = y_start_s4 + i * gap_s4
    col_rgb = hex_to_rgb(col)
    x_card = 0.80
    
    # Elbow connector line to donut
    conn = s4.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x_card + card_w_s4), Inches(y_pos + card_h_s4/2), Inches(cx_s4 - 1.30), Inches(y_pos + card_h_s4/2))
    conn.line.color.rgb = hex_to_rgb('#CBD5E1')
    conn.line.width = Pt(1.5)
    
    add_soft_shadow(s4, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, card_w_s4, card_h_s4, 0.02, 0.02, '#E2E8F0')
    card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(card_w_s4), Inches(card_h_s4))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
    card.line.fill.background()
    
    arch_r = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card + card_w_s4 - 0.90), Inches(y_pos), Inches(0.90), Inches(card_h_s4))
    arch_r.fill.solid()
    arch_r.fill.fore_color.rgb = col_rgb
    arch_r.line.fill.background()
    
    bd = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_card + card_w_s4 - 0.72), Inches(y_pos + 0.10), Inches(0.55), Inches(0.55))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.color.rgb = col_rgb
    bd.line.width = Pt(1.5)
    
    if key in icon_paths_dark:
        s4.shapes.add_picture(icon_paths_dark[key], Inches(x_card + card_w_s4 - 0.64), Inches(y_pos + 0.18), width=Inches(0.38), height=Inches(0.38))
        
    tb_t = s4.shapes.add_textbox(Inches(x_card + 0.20), Inches(y_pos + 0.08), Inches(card_w_s4 - 1.15), Inches(card_h_s4 - 0.16))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf_t.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.0)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.RIGHT

# Right Cards
for i in range(5):
    title, desc, key, col = ITEMS_DATA_ALT[5 + i]
    y_pos = y_start_s4 + i * gap_s4
    col_rgb = hex_to_rgb(col)
    x_card = 8.633
    
    conn = s4.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s4 + 1.30), Inches(y_pos + card_h_s4/2), Inches(x_card), Inches(y_pos + card_h_s4/2))
    conn.line.color.rgb = hex_to_rgb('#CBD5E1')
    conn.line.width = Pt(1.5)
    
    add_soft_shadow(s4, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, card_w_s4, card_h_s4, 0.02, 0.02, '#E2E8F0')
    card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(card_w_s4), Inches(card_h_s4))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
    card.line.fill.background()
    
    arch_l = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(0.90), Inches(card_h_s4))
    arch_l.fill.solid()
    arch_l.fill.fore_color.rgb = col_rgb
    arch_l.line.fill.background()
    
    bd = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_card + 0.17), Inches(y_pos + 0.10), Inches(0.55), Inches(0.55))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.color.rgb = col_rgb
    bd.line.width = Pt(1.5)
    
    if key in icon_paths_dark:
        s4.shapes.add_picture(icon_paths_dark[key], Inches(x_card + 0.25), Inches(y_pos + 0.18), width=Inches(0.38), height=Inches(0.38))
        
    tb_t = s4.shapes.add_textbox(Inches(x_card + 1.05), Inches(y_pos + 0.08), Inches(card_w_s4 - 1.15), Inches(card_h_s4 - 0.16))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf_t.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.0)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.LEFT

# ==============================================================================
# SLIDE 5: CENTER CIRCLE HUB + 10 RADIAL CONNECTORS TO 10 CAPSULE CARDS
# ==============================================================================
s5 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_title(s5, "10Vs Of Big Data")

cx_s5, cy_s5 = 6.666, 4.35
c_diam = 2.20

left_items_s5 = [ITEMS_DATA[0], ITEMS_DATA[2], ITEMS_DATA[4], ITEMS_DATA[6], ITEMS_DATA[8]]
right_items_s5 = [ITEMS_DATA[1], ITEMS_DATA[3], ITEMS_DATA[5], ITEMS_DATA[7], ITEMS_DATA[9]]

y_start_s5 = 1.60
gap_s5 = 1.05
capsule_w = 4.20
capsule_h = 0.76

# Draw Connector Spokes first (behind hub and cards)
for i, (title, desc, key, col) in enumerate(left_items_s5):
    col_rgb = hex_to_rgb(col)
    y_pos = y_start_s5 + i * gap_s5
    x_in = 0.80 + capsule_w - 0.35
    y_mid = y_pos + capsule_h/2
    conn = s5.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s5), Inches(cy_s5), Inches(x_in), Inches(y_mid))
    conn.line.color.rgb = col_rgb
    conn.line.width = Pt(3.5)

for i, (title, desc, key, col) in enumerate(right_items_s5):
    col_rgb = hex_to_rgb(col)
    y_pos = y_start_s5 + i * gap_s5
    x_in = 8.333 + 0.35
    y_mid = y_pos + capsule_h/2
    conn = s5.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s5), Inches(cy_s5), Inches(x_in), Inches(y_mid))
    conn.line.color.rgb = col_rgb
    conn.line.width = Pt(3.5)

# Center Hub Circle over spokes
add_soft_shadow(s5, MSO_SHAPE.OVAL, cx_s5 - c_diam/2, cy_s5 - c_diam/2, c_diam, c_diam, 0.03, 0.03, '#CBD5E1')
hub = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx_s5 - c_diam/2), Inches(cy_s5 - c_diam/2), Inches(c_diam), Inches(c_diam))
hub.fill.solid()
hub.fill.fore_color.rgb = hex_to_rgb('#CBD5E1')
hub.line.color.rgb = hex_to_rgb('#FFFFFF')
hub.line.width = Pt(2.5)

tf_hb = hub.text_frame
p_hb1 = tf_hb.paragraphs[0]
p_hb1.text = "10Vs Of"
p_hb1.font.name = FONT_TITLE
p_hb1.font.size = Pt(20)
p_hb1.font.bold = True
p_hb1.font.color.rgb = hex_to_rgb('#1E293B')
p_hb1.alignment = PP_ALIGN.CENTER
p_hb2 = tf_hb.add_paragraph()
p_hb2.text = "Big Data"
p_hb2.font.name = FONT_TITLE
p_hb2.font.size = Pt(20)
p_hb2.font.bold = True
p_hb2.font.color.rgb = hex_to_rgb('#1E293B')
p_hb2.alignment = PP_ALIGN.CENTER

# 5 Left Capsule Cards
for i, (title, desc, key, col) in enumerate(left_items_s5):
    y_pos = y_start_s5 + i * gap_s5
    col_rgb = hex_to_rgb(col)
    x_card = 0.80
    x_in = x_card + capsule_w
    
    add_soft_shadow(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, capsule_w, capsule_h, 0.03, 0.03, '#CBD5E1')
    cap = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(capsule_w), Inches(capsule_h))
    cap.fill.solid()
    cap.fill.fore_color.rgb = col_rgb
    cap.line.fill.background()
    
    bd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_in - 0.70), Inches(y_pos + 0.10), Inches(0.56), Inches(0.56))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.fill.background()
    
    if key in icon_paths_dark:
        s5.shapes.add_picture(icon_paths_dark[key], Inches(x_in - 0.62), Inches(y_pos + 0.18), width=Inches(0.40), height=Inches(0.40))
        
    tb_tx = s5.shapes.add_textbox(Inches(x_card + 0.20), Inches(y_pos + 0.06), Inches(capsule_w - 0.95), Inches(capsule_h - 0.12))
    tf_tx = tb_tx.text_frame
    tf_tx.word_wrap = True
    p1 = tf_tx.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf_tx.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#F1F5F9')
    p2.alignment = PP_ALIGN.RIGHT

# 5 Right Capsule Cards
for i, (title, desc, key, col) in enumerate(right_items_s5):
    y_pos = y_start_s5 + i * gap_s5
    col_rgb = hex_to_rgb(col)
    x_card = 8.333
    
    add_soft_shadow(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, capsule_w, capsule_h, 0.03, 0.03, '#CBD5E1')
    cap = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(capsule_w), Inches(capsule_h))
    cap.fill.solid()
    cap.fill.fore_color.rgb = col_rgb
    cap.line.fill.background()
    
    bd = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_card + 0.14), Inches(y_pos + 0.10), Inches(0.56), Inches(0.56))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.fill.background()
    
    if key in icon_paths_dark:
        s5.shapes.add_picture(icon_paths_dark[key], Inches(x_card + 0.22), Inches(y_pos + 0.18), width=Inches(0.40), height=Inches(0.40))
        
    tb_tx = s5.shapes.add_textbox(Inches(x_card + 0.80), Inches(y_pos + 0.06), Inches(capsule_w - 0.95), Inches(capsule_h - 0.12))
    tf_tx = tb_tx.text_frame
    tf_tx.word_wrap = True
    p1 = tf_tx.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#FFFFFF')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf_tx.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = hex_to_rgb('#F1F5F9')
    p2.alignment = PP_ALIGN.LEFT

# ==============================================================================
# SLIDE 6: VERTICAL CONNECTED TIMELINE LIST (WITH 01..10 NUMBER BADGES)
# ==============================================================================
s6 = prs.slides.add_slide(prs.slide_layouts[6])

# Left big title
tb_s6_title = s6.shapes.add_textbox(Inches(0.80), Inches(3.20), Inches(3.20), Inches(2.00))
tf_s6 = tb_s6_title.text_frame
tf_s6.word_wrap = True
p_s6_1 = tf_s6.paragraphs[0]
p_s6_1.text = "10Vs Of"
p_s6_1.font.name = FONT_TITLE
p_s6_1.font.size = Pt(36)
p_s6_1.font.bold = True
p_s6_1.font.color.rgb = hex_to_rgb('#1E293B')
p_s6_2 = tf_s6.add_paragraph()
p_s6_2.text = "Big Data"
p_s6_2.font.name = FONT_TITLE
p_s6_2.font.size = Pt(36)
p_s6_2.font.bold = True
p_s6_2.font.color.rgb = hex_to_rgb('#1E293B')

y_start_s6 = 0.55
gap_s6 = 0.65
card_w_s6 = 7.10
card_h_s6 = 0.52
badge_x = 4.35

# Vertical continuous connector line behind nodes
line_s6 = s6.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(badge_x + 0.22), Inches(y_start_s6 + 0.26), Inches(badge_x + 0.22), Inches(y_start_s6 + 9 * gap_s6 + 0.26))
line_s6.line.color.rgb = hex_to_rgb('#CBD5E1')
line_s6.line.width = Pt(2.5)

for i in range(10):
    title, desc, key, col = ITEMS_DATA[i]
    y_pos = y_start_s6 + i * gap_s6
    col_rgb = hex_to_rgb(col)
    
    # 1. Main text card
    card_x = 4.90
    add_soft_shadow(s6, MSO_SHAPE.ROUNDED_RECTANGLE, card_x, y_pos, card_w_s6, card_h_s6, 0.02, 0.02, '#E2E8F0')
    card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(card_x), Inches(y_pos), Inches(card_w_s6), Inches(card_h_s6))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#F8FAFC')
    card.line.color.rgb = hex_to_rgb('#E2E8F0')
    card.line.width = Pt(1.0)
    
    # 2. Left circular node over vertical line
    node = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_x), Inches(y_pos + 0.04), Inches(0.44), Inches(0.44))
    node.fill.solid()
    node.fill.fore_color.rgb = col_rgb
    node.line.color.rgb = hex_to_rgb('#FFFFFF')
    node.line.width = Pt(1.5)
    
    if key in icon_paths_white:
        s6.shapes.add_picture(icon_paths_white[key], Inches(badge_x + 0.06), Inches(y_pos + 0.10), width=Inches(0.32), height=Inches(0.32))
        
    # 3. Card Title & Description
    tb = s6.shapes.add_textbox(Inches(card_x + 0.20), Inches(y_pos + 0.03), Inches(card_w_s6 - 1.10), Inches(card_h_s6 - 0.06))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.0)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    
    # 4. Right Number Badge (01..10)
    num_badge = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(card_x + card_w_s6 - 0.65), Inches(y_pos), Inches(0.65), Inches(card_h_s6))
    num_badge.fill.solid()
    num_badge.fill.fore_color.rgb = col_rgb
    num_badge.line.fill.background()
    
    tf_num = num_badge.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.text = f"{i+1:02d}"
    p_num.font.name = FONT_BODY
    p_num.font.size = Pt(11)
    p_num.font.bold = True
    p_num.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_num.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 7: CENTER THIN COLOR DONUT WITH RADIATING ARROWS & 10 ACCENT CARDS
# ==============================================================================
s7 = prs.slides.add_slide(prs.slide_layouts[6])
add_slide_title(s7, "10Vs Of Big Data")

p_donut_s7 = os.path.join(output_assets, 's7_donut_clean.png')
cx_s7, cy_s7 = 6.666, 4.30
donut_w_s7 = 2.50

y_start_s7 = 1.60
gap_s7 = 1.05
card_w_s7 = 4.10
card_h_s7 = 0.78

# 1. Draw Connectors FIRST (behind donut and cards)
for i in range(5):
    y_pos = y_start_s7 + i * gap_s7
    x_card = 0.80
    x_inner = x_card + card_w_s7
    y_mid = y_pos + card_h_s7/2
    conn = s7.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s7), Inches(cy_s7), Inches(x_inner), Inches(y_mid))
    conn.line.color.rgb = hex_to_rgb('#CBD5E1')
    conn.line.width = Pt(1.5)

for i in range(5):
    y_pos = y_start_s7 + i * gap_s7
    x_card = 8.433
    y_mid = y_pos + card_h_s7/2
    conn = s7.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx_s7), Inches(cy_s7), Inches(x_card), Inches(y_mid))
    conn.line.color.rgb = hex_to_rgb('#CBD5E1')
    conn.line.width = Pt(1.5)

# 2. Thin Donut image OVER connectors
s7.shapes.add_picture(p_donut_s7, Inches(cx_s7 - donut_w_s7/2), Inches(cy_s7 - donut_w_s7/2), width=Inches(donut_w_s7), height=Inches(donut_w_s7))

# 3. Center hole white circle
add_soft_shadow(s7, MSO_SHAPE.OVAL, cx_s7 - 0.90, cy_s7 - 0.90, 1.80, 1.80, 0.02, 0.02, '#CBD5E1')
c_hole_s7 = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx_s7 - 0.90), Inches(cy_s7 - 0.90), Inches(1.80), Inches(1.80))
c_hole_s7.fill.solid()
c_hole_s7.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
c_hole_s7.line.fill.background()

tf_h7 = c_hole_s7.text_frame
p_h7_1 = tf_h7.paragraphs[0]
p_h7_1.text = "10Vs Of"
p_h7_1.font.name = FONT_TITLE
p_h7_1.font.size = Pt(17)
p_h7_1.font.bold = True
p_h7_1.font.color.rgb = hex_to_rgb('#1E293B')
p_h7_1.alignment = PP_ALIGN.CENTER
p_h7_2 = tf_h7.add_paragraph()
p_h7_2.text = "Big Data"
p_h7_2.font.name = FONT_TITLE
p_h7_2.font.size = Pt(17)
p_h7_2.font.bold = True
p_h7_2.font.color.rgb = hex_to_rgb('#1E293B')
p_h7_2.alignment = PP_ALIGN.CENTER

# 4. Left Cards (Slide 7)
for i in range(5):
    title, desc, key, col = ITEMS_DATA_ALT[i]
    y_pos = y_start_s7 + i * gap_s7
    col_rgb = hex_to_rgb(col)
    x_card = 0.80
    x_inner = x_card + card_w_s7
    
    add_soft_shadow(s7, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, card_w_s7, card_h_s7, 0.02, 0.02, '#E2E8F0')
    card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(card_w_s7), Inches(card_h_s7))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.color.rgb = hex_to_rgb('#E2E8F0')
    card.line.width = Pt(1.0)
    
    tab_r = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_inner - 0.90), Inches(y_pos), Inches(0.90), Inches(card_h_s7))
    tab_r.fill.solid()
    tab_r.fill.fore_color.rgb = col_rgb
    tab_r.line.fill.background()
    
    bd = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_inner - 0.72), Inches(y_pos + 0.11), Inches(0.56), Inches(0.56))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.color.rgb = col_rgb
    bd.line.width = Pt(1.5)
    
    if key in icon_paths_dark:
        s7.shapes.add_picture(icon_paths_dark[key], Inches(x_inner - 0.64), Inches(y_pos + 0.19), width=Inches(0.40), height=Inches(0.40))
        
    tb_t = s7.shapes.add_textbox(Inches(x_card + 0.20), Inches(y_pos + 0.08), Inches(card_w_s7 - 1.20), Inches(card_h_s7 - 0.16))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.RIGHT
    p2 = tf_t.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.0)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.RIGHT

# 5. Right Cards (Slide 7)
for i in range(5):
    title, desc, key, col = ITEMS_DATA_ALT[5 + i]
    y_pos = y_start_s7 + i * gap_s7
    col_rgb = hex_to_rgb(col)
    x_card = 8.433
    
    add_soft_shadow(s7, MSO_SHAPE.ROUNDED_RECTANGLE, x_card, y_pos, card_w_s7, card_h_s7, 0.02, 0.02, '#E2E8F0')
    card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(card_w_s7), Inches(card_h_s7))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    card.line.color.rgb = hex_to_rgb('#E2E8F0')
    card.line.width = Pt(1.0)
    
    tab_l = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_card), Inches(y_pos), Inches(0.90), Inches(card_h_s7))
    tab_l.fill.solid()
    tab_l.fill.fore_color.rgb = col_rgb
    tab_l.line.fill.background()
    
    bd = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_card + 0.16), Inches(y_pos + 0.11), Inches(0.56), Inches(0.56))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
    bd.line.color.rgb = col_rgb
    bd.line.width = Pt(1.5)
    
    if key in icon_paths_dark:
        s7.shapes.add_picture(icon_paths_dark[key], Inches(x_card + 0.24), Inches(y_pos + 0.19), width=Inches(0.40), height=Inches(0.40))
        
    tb_t = s7.shapes.add_textbox(Inches(x_card + 1.05), Inches(y_pos + 0.08), Inches(card_w_s7 - 1.20), Inches(card_h_s7 - 0.16))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = title
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10.5)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb('#1E293B')
    p1.alignment = PP_ALIGN.LEFT
    p2 = tf_t.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_BODY
    p2.font.size = Pt(8.0)
    p2.font.color.rgb = hex_to_rgb('#64748B')
    p2.alignment = PP_ALIGN.LEFT

# ==================== SAVE OUTPUT PPTX ====================
final_output_path = os.path.abspath(os.path.join(SOURCE_DIR, 'presentation_95_precision.pptx'))
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print(f"Presentation saved successfully to {final_output_path}")

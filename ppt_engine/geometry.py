import math
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from ppt_engine.theme import hex_to_rgb, Colors

def add_soft_shadow(slide, shape_type, x, y, w, h, offset_x=0.03, offset_y=0.03, shadow_color=Colors.SHADOW_AMBIENT):
    """Adds a soft ambient drop shadow behind a shape."""
    shd = slide.shapes.add_shape(
        shape_type,
        Inches(x + offset_x),
        Inches(y + offset_y),
        Inches(w),
        Inches(h)
    )
    shd.fill.solid()
    shd.fill.fore_color.rgb = hex_to_rgb(shadow_color)
    shd.line.fill.background()
    return shd

def create_seamless_callout(slide, x, y, w, h, notch_w=0.35, notch_h=0.25, corner_r=0.15, 
                            bg_color='#FFFFFF', border_color='#CBD5E1', border_width=1.5,
                            has_shadow=True):
    """
    Builds a single, unified seamless speech bubble callout (no seams or separate triangles).
    The bottom pointer is mathematically integrated into the perimeter path.
    """
    x_in = Inches(x)
    y_in = Inches(y)
    w_in = Inches(w)
    h_in = Inches(h)
    
    cx_in = Inches(x + w / 2.0)
    nw2 = Inches(notch_w / 2.0)
    nh = Inches(notch_h)
    
    points = [
        (x_in, y_in),
        (x_in + w_in, y_in),
        (x_in + w_in, y_in + h_in),
        (cx_in + nw2, y_in + h_in),
        (cx_in, y_in + h_in + nh),
        (cx_in - nw2, y_in + h_in),
        (x_in, y_in + h_in),
    ]
    
    if has_shadow:
        shd_pts = [(px + Inches(0.03), py + Inches(0.03)) for px, py in points]
        fb_shd = slide.shapes.build_freeform(shd_pts[0][0], shd_pts[0][1])
        fb_shd.add_line_segments(shd_pts[1:], close=True)
        shd = fb_shd.convert_to_shape()
        shd.fill.solid()
        shd.fill.fore_color.rgb = hex_to_rgb(Colors.SHADOW_AMBIENT)
        shd.line.fill.background()
        
    fb = slide.shapes.build_freeform(points[0][0], points[0][1])
    fb.add_line_segments(points[1:], close=True)
    shape = fb.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(bg_color)
    shape.line.color.rgb = hex_to_rgb(border_color)
    shape.line.width = Pt(border_width)
    return shape

def create_chevron(slide, x, y, w, h, fill_color, text='', font_color='#FFFFFF', font_size=11, has_shadow=False):
    """Creates a clean chevron with optional shadow."""
    if has_shadow:
        add_soft_shadow(slide, MSO_SHAPE.CHEVRON, x, y, w, h, 0.03, 0.03)
        
    chv = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
    chv.fill.solid()
    chv.fill.fore_color.rgb = hex_to_rgb(fill_color)
    chv.line.fill.background()
    
    if text:
        tf = chv.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = 'Montserrat'
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(font_color)
    return chv

def create_downward_banner(slide, x, y, w, h, point_h=0.35, fill_color='#2B6CB0', has_shadow=True):
    """Creates a downward pointing banner ribbon shape."""
    x_in = Inches(x)
    y_in = Inches(y)
    w_in = Inches(w)
    h_in = Inches(h)
    p_in = Inches(point_h)
    
    points = [
        (x_in, y_in),
        (x_in + w_in, y_in),
        (x_in + w_in, y_in + h_in - p_in),
        (x_in + w_in / 2.0, y_in + h_in),
        (x_in, y_in + h_in - p_in),
    ]
    
    if has_shadow:
        shd_pts = [(px + Inches(0.03), py + Inches(0.03)) for px, py in points]
        fb_shd = slide.shapes.build_freeform(shd_pts[0][0], shd_pts[0][1])
        fb_shd.add_line_segments(shd_pts[1:], close=True)
        shd = fb_shd.convert_to_shape()
        shd.fill.solid()
        shd.fill.fore_color.rgb = hex_to_rgb(Colors.SHADOW_AMBIENT)
        shd.line.fill.background()
        
    fb = slide.shapes.build_freeform(points[0][0], points[0][1])
    fb.add_line_segments(points[1:], close=True)
    shape = fb.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
    shape.line.fill.background()
    return shape

def create_smooth_path(slide, points, line_color='#CBD5E1', line_width=1.5, is_dashed=False):
    """Builds a connected multi-point path for winding roads and timelines."""
    pts_in = [(Inches(px), Inches(py)) for px, py in points]
    fb = slide.shapes.build_freeform(pts_in[0][0], pts_in[0][1])
    fb.add_line_segments(pts_in[1:], close=False)
    line = fb.convert_to_shape()
    line.fill.background()
    line.line.color.rgb = hex_to_rgb(line_color)
    line.line.width = Pt(line_width)
    return line

def create_circular_arc(slide, cx: float, cy: float, radius: float, start_deg: float, end_deg: float,
                        line_color='#CBD5E1', line_width=1.5, num_pts=36):
    """
    Creates a mathematically smooth circular arc spine from start_deg to end_deg.
    Zero degrees is at 3 o'clock (positive X axis).
    """
    pts = []
    step = (end_deg - start_deg) / (num_pts - 1)
    for i in range(num_pts):
        deg = start_deg + i * step
        rad = math.radians(deg)
        pts.append((cx + radius * math.cos(rad), cy + radius * math.sin(rad)))
    return create_smooth_path(slide, pts, line_color=line_color, line_width=line_width)

def create_horizontal_connector(slide, x_start: float, x_end: float, y: float,
                                line_color='#CBD5E1', line_width=1.5):
    """Creates a clean horizontal line connecting a node to a card."""
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_start), Inches(y - 0.01), Inches(x_end - x_start), Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = hex_to_rgb(line_color)
    ln.line.fill.background()
    return ln

def add_header_accent(slide, x: float, y: float, color=Colors.ORANGE, dash_w=0.45, dash_h=0.08, dot_size=0.08, gap=0.08):
    """
    Renders the signature brand accent above section titles: a horizontal pill and a circular dot (— •).
    """
    dash = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(dash_w), Inches(dash_h))
    dash.fill.solid()
    dash.fill.fore_color.rgb = hex_to_rgb(color)
    dash.line.fill.background()
    
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + dash_w + gap), Inches(y), Inches(dot_size), Inches(dot_size))
    dot.fill.solid()
    dot.fill.fore_color.rgb = hex_to_rgb(color)
    dot.line.fill.background()
    return dash, dot

def add_background_wireframe_squares(slide, positions=None, border_color='#EDF2F7', border_width=1.2):
    """
    Renders soft floating rounded square wireframes in the canvas background to match modern slide aesthetics.
    """
    if positions is None:
        positions = [
            (0.80, 1.20, 1.30, 1.30),
            (1.50, 4.50, 1.00, 1.00),
            (11.20, 0.80, 1.40, 1.40),
            (10.50, 5.20, 1.10, 1.10)
        ]
    squares = []
    for x, y, w, h in positions:
        sq = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sq.fill.background()
        sq.line.color.rgb = hex_to_rgb(border_color)
        sq.line.width = Pt(border_width)
        squares.append(sq)
    return squares

def add_dashed_vertical_divider(slide, x: float, y: float, h: float, dash_count=7, dash_w=0.015, color='#CBD5E1'):
    """
    Renders a clean dashed vertical divider line between metric columns.
    """
    dash_h = h / (dash_count * 2 - 1)
    dashes = []
    for i in range(dash_count):
        dy = y + i * 2 * dash_h
        d = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(dy), Inches(dash_w), Inches(dash_h))
        d.fill.solid()
        d.fill.fore_color.rgb = hex_to_rgb(color)
        d.line.fill.background()
        dashes.append(d)
    return dashes

def add_progress_indicators(slide, x: float, y: float, total=4, active_idx=0, active_color=Colors.NAVY, inactive_color='#CBD5E1', bar_w=0.55, bar_h=0.06, gap=0.10):
    """
    Renders slider navigation progress indicator bars at the bottom of the slide.
    """
    bars = []
    for i in range(total):
        bx = x + i * (bar_w + gap)
        b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx), Inches(y), Inches(bar_w), Inches(bar_h))
        b.fill.solid()
        b.fill.fore_color.rgb = hex_to_rgb(active_color if i == active_idx else inactive_color)
        b.line.fill.background()
        bars.append(b)
    return bars


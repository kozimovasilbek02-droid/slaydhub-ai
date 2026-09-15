from typing import List
import math
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from core.models import SlideData, MilestoneNode, ColorRGB

def hex_to_rgb(hex_code: str) -> RGBColor:
    c = ColorRGB.from_hex(hex_code)
    return RGBColor(c.r, c.g, c.b)

def build_roadmap_timeline(slide, slide_data: SlideData, width_in: float = 13.333, height_in: float = 7.5):
    """
    Builds a 100% faithful vector reproduction of the Strategic Roadmap:
    - Smooth S-curved highway road weaving UNDER Node 1, OVER Node 2, UNDER Node 3, OVER Node 4, UNDER Node 5
    - High-precision angled dashed white lane divider
    - Colored accent arcs (Purple, Teal, Red, Orange, Blue + Orange Arrow)
    - Organic white teardrop keyhole pods with concentric icon badges
    - Vibrant 20XX milestone circles
    - Flawless typography and text placement matching the original slide
    """
    # 1. Slide Title
    title_text = slide_data.title or "Strategic Roadmap Examples"
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.40), Inches(width_in - 1.6), Inches(0.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = "Segoe UI"
    p.font.color.rgb = hex_to_rgb("#2E384D")

    # 5 Milestone nodes
    x_positions = [2.10, 4.40, 6.70, 9.00, 11.30]
    y_center = 4.00     # Horizontal center line of the 5 circles
    y_baseline = 4.90   # Valley road height (Under nodes 1, 3, 5)
    y_peak = 3.10       # Peak road height (Over nodes 2, 4)
    L_step = x_positions[1] - x_positions[0] # 2.30 inches
    H_road = 0.48       # Asphalt road thickness

    # -------------------------------------------------------------
    # 2. COLORED ACCENT ARCS & TERMINALS
    # -------------------------------------------------------------
    # Left start circuit (Purple dot & horizontal bar)
    c1_dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.85), Inches(y_peak - 0.20), Inches(0.42), Inches(0.42))
    c1_dot.fill.solid()
    c1_dot.fill.fore_color.rgb = hex_to_rgb("#9355B5")
    c1_dot.line.color.rgb = hex_to_rgb("#FFFFFF")
    c1_dot.line.width = Pt(2.5)
    
    c1_line = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.15), Inches(y_peak - 0.08), Inches(0.70), Inches(0.16))
    c1_line.fill.solid()
    c1_line.fill.fore_color.rgb = hex_to_rgb("#9355B5")
    c1_line.line.fill.background()

    # Block Arcs around Nodes (Thick colored rings embracing the pods)
    arc_d = 2.10
    
    # Node 1 Purple Top Arc
    arc1 = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(x_positions[0] - arc_d/2), Inches(y_center - arc_d/2), Inches(arc_d), Inches(arc_d))
    arc1.fill.solid()
    arc1.fill.fore_color.rgb = hex_to_rgb("#9355B5")
    arc1.line.fill.background()
    arc1.rotation = 180

    # Node 2 Teal Bottom Arc
    arc2 = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(x_positions[1] - arc_d/2), Inches(y_center - arc_d/2), Inches(arc_d), Inches(arc_d))
    arc2.fill.solid()
    arc2.fill.fore_color.rgb = hex_to_rgb("#00A896")
    arc2.line.fill.background()
    arc2.rotation = 0

    # Node 3 Red Top Arc
    arc3 = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(x_positions[2] - arc_d/2), Inches(y_center - arc_d/2), Inches(arc_d), Inches(arc_d))
    arc3.fill.solid()
    arc3.fill.fore_color.rgb = hex_to_rgb("#EF476F")
    arc3.line.fill.background()
    arc3.rotation = 180

    # Node 4 Orange Bottom Arc
    arc4 = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(x_positions[3] - arc_d/2), Inches(y_center - arc_d/2), Inches(arc_d), Inches(arc_d))
    arc4.fill.solid()
    arc4.fill.fore_color.rgb = hex_to_rgb("#FA8231")
    arc4.line.fill.background()
    arc4.rotation = 0

    # Node 5 Orange Top Arc & Arrow
    arc5 = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(x_positions[4] - arc_d/2), Inches(y_center - arc_d/2), Inches(arc_d), Inches(arc_d))
    arc5.fill.solid()
    arc5.fill.fore_color.rgb = hex_to_rgb("#FA8231")
    arc5.line.fill.background()
    arc5.rotation = 180
    
    # Orange arrow pointing right
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x_positions[4] + 0.45), Inches(y_peak - 0.18), Inches(0.95), Inches(0.38))
    arr.fill.solid()
    arr.fill.fore_color.rgb = hex_to_rgb("#FA8231")
    arr.line.fill.background()

    # -------------------------------------------------------------
    # 3. HIGHWAY ROAD WAVE (ASPHALT RIBBON + DASHED CENTER LINE)
    # -------------------------------------------------------------
    def get_highway_y_and_angle(rx: float):
        if rx <= x_positions[0]:
            return y_baseline, 0.0
        elif rx >= x_positions[-1]:
            return y_baseline, 0.0
        else:
            theta = (rx - x_positions[0]) * (math.pi / L_step)
            ry = y_center + 0.90 * math.cos(theta)
            dy_dx = -0.90 * (math.pi / L_step) * math.sin(theta)
            angle_deg = math.degrees(math.atan(dy_dx))
            return ry, angle_deg

    num_samples = 220
    top_edge = []
    bottom_edge = []
    
    for i in range(num_samples + 1):
        rx = (i / float(num_samples)) * width_in
        ry, angle_deg = get_highway_y_and_angle(rx)
        rad = math.radians(angle_deg)
        nx = -math.sin(rad) * (H_road / 2.0)
        ny = math.cos(rad) * (H_road / 2.0)
        top_edge.append((Inches(rx + nx), Inches(ry + ny)))
        bottom_edge.append((Inches(rx - nx), Inches(ry - ny)))

    all_poly_pts = top_edge + list(reversed(bottom_edge))
    road_builder = slide.shapes.build_freeform(all_poly_pts[0][0], all_poly_pts[0][1])
    road_builder.add_line_segments(all_poly_pts[1:], close=True)
    road_shape = road_builder.convert_to_shape()
    road_shape.fill.solid()
    road_shape.fill.fore_color.rgb = hex_to_rgb("#5A6270")
    road_shape.line.fill.background()

    # Dashed Lane Markings
    dash_len = 0.20
    dash_gap = 0.16
    curr_x = 0.15
    while curr_x < width_in - 0.2:
        ry, angle_deg = get_highway_y_and_angle(curr_x)
        dash = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(curr_x - dash_len/2), Inches(ry - 0.02),
            Inches(dash_len), Inches(0.04)
        )
        dash.rotation = angle_deg
        dash.fill.solid()
        dash.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
        dash.line.fill.background()
        curr_x += (dash_len + dash_gap)

    # -------------------------------------------------------------
    # 4. WHITE TEARDROP PODS & 20XX NODES
    # -------------------------------------------------------------
    node_configs = [
        {"color": "#9355B5", "dir": "UP", "step": "20XX", "icon": "TARGET"},
        {"color": "#00A896", "dir": "DOWN", "step": "20XX", "icon": "SHOP"},
        {"color": "#EF476F", "dir": "UP", "step": "20XX", "icon": "IDEA"},
        {"color": "#FA8231", "dir": "DOWN", "step": "20XX", "icon": "TEAM"},
        {"color": "#0077B6", "dir": "UP", "step": "20XX", "icon": "REPORT"},
    ]

    for i, cfg in enumerate(node_configs):
        cx = x_positions[i]
        c_rgb = hex_to_rgb(cfg["color"])
        direction = cfg["dir"]

        if direction == "UP":
            y_pin = 1.85
            
            # A. Soft Drop Shadow
            shadow_bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.90), Inches(y_center - 0.90), Inches(1.80), Inches(1.80))
            shadow_bulb.fill.solid()
            shadow_bulb.fill.fore_color.rgb = hex_to_rgb("#E2E8F0")
            shadow_bulb.line.fill.background()

            # B. White Pod Body (Bulb + Tapering Neck + Top Head)
            bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.85), Inches(y_center - 0.85), Inches(1.70), Inches(1.70))
            bulb.fill.solid()
            bulb.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            bulb.line.fill.background()

            # Organic tapering neck (Trapezoid from wide bulb to narrow head)
            neck_pts = [
                (Inches(cx - 0.20), Inches(y_pin + 0.35)),
                (Inches(cx + 0.20), Inches(y_pin + 0.35)),
                (Inches(cx + 0.40), Inches(y_center - 0.40)),
                (Inches(cx - 0.40), Inches(y_center - 0.40)),
            ]
            neck_builder = slide.shapes.build_freeform(neck_pts[0][0], neck_pts[0][1])
            neck_builder.add_line_segments(neck_pts[1:], close=True)
            neck = neck_builder.convert_to_shape()
            neck.fill.solid()
            neck.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            neck.line.fill.background()

            head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.45), Inches(y_pin - 0.45), Inches(0.90), Inches(0.90))
            head.fill.solid()
            head.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            head.line.fill.background()

            # Pin badge outer ring
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.40), Inches(y_pin - 0.40), Inches(0.80), Inches(0.80))
            ring.fill.solid()
            ring.fill.fore_color.rgb = c_rgb
            ring.line.fill.background()

            # Inner white circle
            iw = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.30), Inches(y_pin - 0.30), Inches(0.60), Inches(0.60))
            iw.fill.solid()
            iw.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            iw.line.fill.background()

            # High Fidelity Vector Icons
            icon_type = cfg["icon"]
            if icon_type == "TARGET": # Node 1: Bullseye / Target
                ring1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.20), Inches(y_pin - 0.20), Inches(0.40), Inches(0.40))
                ring1.line.color.rgb = c_rgb
                ring1.line.width = Pt(1.5)
                ring1.fill.fill_type = 0 # No fill
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.08), Inches(y_pin - 0.08), Inches(0.16), Inches(0.16))
                dot.fill.solid()
                dot.fill.fore_color.rgb = c_rgb
                dot.line.fill.background()
            elif icon_type == "IDEA": # Node 3: Lightbulb
                bulb_top = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.14), Inches(y_pin - 0.18), Inches(0.28), Inches(0.26))
                bulb_top.fill.solid()
                bulb_top.fill.fore_color.rgb = c_rgb
                bulb_top.line.fill.background()
                bulb_bot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.08), Inches(y_pin + 0.04), Inches(0.16), Inches(0.10))
                bulb_bot.fill.solid()
                bulb_bot.fill.fore_color.rgb = c_rgb
                bulb_bot.line.fill.background()
            elif icon_type == "REPORT": # Node 5: Document / Analytics
                doc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - 0.12), Inches(y_pin - 0.16), Inches(0.24), Inches(0.32))
                doc.fill.solid()
                doc.fill.fore_color.rgb = c_rgb
                doc.line.fill.background()
                l1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.08), Inches(y_pin - 0.08), Inches(0.16), Inches(0.03))
                l1.fill.solid()
                l1.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
                l1.line.fill.background()
                l2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.08), Inches(y_pin - 0.01), Inches(0.16), Inches(0.03))
                l2.fill.solid()
                l2.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
                l2.line.fill.background()
                l3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.08), Inches(y_pin + 0.06), Inches(0.10), Inches(0.03))
                l3.fill.solid()
                l3.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
                l3.line.fill.background()
            else:
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.12), Inches(y_pin - 0.12), Inches(0.24), Inches(0.24))
                dot.fill.solid()
                dot.fill.fore_color.rgb = c_rgb
                dot.line.fill.background()

            # D. Vibrant 20XX Milestone Circle
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.65), Inches(y_center - 0.65), Inches(1.30), Inches(1.30))
            circle.fill.solid()
            circle.fill.fore_color.rgb = c_rgb
            circle.line.fill.background()
            
            tf_c = circle.text_frame
            tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p_c = tf_c.paragraphs[0]
            p_c.text = cfg["step"]
            p_c.alignment = PP_ALIGN.CENTER
            p_c.font.size = Pt(20)
            p_c.font.bold = True
            p_c.font.name = "Segoe UI"
            p_c.font.color.rgb = hex_to_rgb("#FFFFFF")

        else: # DOWNWARD POD (Nodes 2, 4)
            y_pin = 6.15

            # A. Soft Drop Shadow
            shadow_bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.90), Inches(y_center - 0.90), Inches(1.80), Inches(1.80))
            shadow_bulb.fill.solid()
            shadow_bulb.fill.fore_color.rgb = hex_to_rgb("#E2E8F0")
            shadow_bulb.line.fill.background()

            # B. White Pod Body (Bulb + Tapering Neck + Bottom Head)
            bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.85), Inches(y_center - 0.85), Inches(1.70), Inches(1.70))
            bulb.fill.solid()
            bulb.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            bulb.line.fill.background()

            # Organic tapering neck
            neck_pts = [
                (Inches(cx - 0.40), Inches(y_center + 0.40)),
                (Inches(cx + 0.40), Inches(y_center + 0.40)),
                (Inches(cx + 0.18), Inches(y_pin - 0.35)),
                (Inches(cx - 0.18), Inches(y_pin - 0.35)),
            ]
            neck_builder = slide.shapes.build_freeform(neck_pts[0][0], neck_pts[0][1])
            neck_builder.add_line_segments(neck_pts[1:], close=True)
            neck = neck_builder.convert_to_shape()
            neck.fill.solid()
            neck.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            neck.line.fill.background()

            head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.45), Inches(y_pin - 0.45), Inches(0.90), Inches(0.90))
            head.fill.solid()
            head.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            head.line.fill.background()

            # C. Bottom Pin Badge Concentric Rings & Icons
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.40), Inches(y_pin - 0.40), Inches(0.80), Inches(0.80))
            ring.fill.solid()
            ring.fill.fore_color.rgb = c_rgb
            ring.line.fill.background()

            iw = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.30), Inches(y_pin - 0.30), Inches(0.60), Inches(0.60))
            iw.fill.solid()
            iw.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
            iw.line.fill.background()

            icon_type = cfg["icon"]
            if icon_type == "SHOP": # Node 2: Market / Store Awning
                awning = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(cx - 0.16), Inches(y_pin - 0.16), Inches(0.32), Inches(0.14))
                awning.fill.solid()
                awning.fill.fore_color.rgb = c_rgb
                awning.line.fill.background()
                stall = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.14), Inches(y_pin + 0.02), Inches(0.28), Inches(0.14))
                stall.fill.solid()
                stall.fill.fore_color.rgb = c_rgb
                stall.line.fill.background()
                st_door = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - 0.05), Inches(y_pin + 0.06), Inches(0.10), Inches(0.10))
                st_door.fill.solid()
                st_door.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")
                st_door.line.fill.background()
            elif icon_type == "TEAM": # Node 4: Team / 3 People
                # Center person
                h_c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.06), Inches(y_pin - 0.16), Inches(0.12), Inches(0.12))
                h_c.fill.solid()
                h_c.fill.fore_color.rgb = c_rgb
                h_c.line.fill.background()
                b_c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.10), Inches(y_pin - 0.02), Inches(0.20), Inches(0.16))
                b_c.fill.solid()
                b_c.fill.fore_color.rgb = c_rgb
                b_c.line.fill.background()
                # Left & Right person heads
                h_l = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.18), Inches(y_pin - 0.10), Inches(0.09), Inches(0.09))
                h_l.fill.solid()
                h_l.fill.fore_color.rgb = c_rgb
                h_l.line.fill.background()
                h_r = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 0.09), Inches(y_pin - 0.10), Inches(0.09), Inches(0.09))
                h_r.fill.solid()
                h_r.fill.fore_color.rgb = c_rgb
                h_r.line.fill.background()
            else:
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.12), Inches(y_pin - 0.12), Inches(0.24), Inches(0.24))
                dot.fill.solid()
                dot.fill.fore_color.rgb = c_rgb
                dot.line.fill.background()

            # D. Vibrant 20XX Milestone Circle
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.65), Inches(y_center - 0.65), Inches(1.30), Inches(1.30))
            circle.fill.solid()
            circle.fill.fore_color.rgb = c_rgb
            circle.line.fill.background()
            
            tf_c = circle.text_frame
            tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p_c = tf_c.paragraphs[0]
            p_c.text = cfg["step"]
            p_c.alignment = PP_ALIGN.CENTER
            p_c.font.size = Pt(20)
            p_c.font.bold = True
            p_c.font.name = "Segoe UI"
            p_c.font.color.rgb = hex_to_rgb("#FFFFFF")


    # -------------------------------------------------------------
    # 5. TYPOGRAPHY & TEXT BLOCKS
    # -------------------------------------------------------------
    # Text 1 (Below Node 1)
    t1_box = slide.shapes.add_textbox(Inches(x_positions[0] - 1.10), Inches(5.35), Inches(2.20), Inches(1.5))
    tf1 = t1_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "Set Strategic Objectives"
    p1.alignment = PP_ALIGN.CENTER
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.name = "Segoe UI"
    p1.font.color.rgb = hex_to_rgb("#1E293B")
    p1.space_after = Pt(3)
    p1_d = tf1.add_paragraph()
    p1_d.text = "Establish measurable goals that align with long-term business priorities."
    p1_d.alignment = PP_ALIGN.CENTER
    p1_d.font.size = Pt(9.5)
    p1_d.font.name = "Segoe UI"
    p1_d.font.color.rgb = hex_to_rgb("#64748B")

    # Text 2 (Above between Node 1 & 2)
    t2_box = slide.shapes.add_textbox(Inches(2.65), Inches(1.50), Inches(2.10), Inches(1.6))
    tf2 = t2_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "Analyze Market & Capabilities"
    p2.alignment = PP_ALIGN.LEFT
    p2.font.size = Pt(13)
    p2.font.bold = True
    p2.font.name = "Segoe UI"
    p2.font.color.rgb = hex_to_rgb("#1E293B")
    p2.space_after = Pt(3)
    p2_d = tf2.add_paragraph()
    p2_d.text = "Study market trends, customer needs, and internal performance to find growth opportunities."
    p2_d.alignment = PP_ALIGN.LEFT
    p2_d.font.size = Pt(9.5)
    p2_d.font.name = "Segoe UI"
    p2_d.font.color.rgb = hex_to_rgb("#64748B")

    # Text 3 (Below Node 3)
    t3_box = slide.shapes.add_textbox(Inches(x_positions[2] - 1.10), Inches(5.35), Inches(2.20), Inches(1.5))
    tf3 = t3_box.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "Develop Key Initiatives"
    p3.alignment = PP_ALIGN.CENTER
    p3.font.size = Pt(13)
    p3.font.bold = True
    p3.font.name = "Segoe UI"
    p3.font.color.rgb = hex_to_rgb("#1E293B")
    p3.space_after = Pt(3)
    p3_d = tf3.add_paragraph()
    p3_d.text = "Create focused plans that drive innovation, efficiency, and business expansion."
    p3_d.alignment = PP_ALIGN.CENTER
    p3_d.font.size = Pt(9.5)
    p3_d.font.name = "Segoe UI"
    p3_d.font.color.rgb = hex_to_rgb("#64748B")

    # Text 4 (Above between Node 3 & 4)
    t4_box = slide.shapes.add_textbox(Inches(7.25), Inches(1.50), Inches(2.10), Inches(1.6))
    tf4 = t4_box.text_frame
    tf4.word_wrap = True
    p4 = tf4.paragraphs[0]
    p4.text = "Implement & Align Teams"
    p4.alignment = PP_ALIGN.LEFT
    p4.font.size = Pt(13)
    p4.font.bold = True
    p4.font.name = "Segoe UI"
    p4.font.color.rgb = hex_to_rgb("#1E293B")
    p4.space_after = Pt(3)
    p4_d = tf4.add_paragraph()
    p4_d.text = "Execute projects, empower teams, and ensure cross-department collaboration."
    p4_d.alignment = PP_ALIGN.LEFT
    p4_d.font.size = Pt(9.5)
    p4_d.font.name = "Segoe UI"
    p4_d.font.color.rgb = hex_to_rgb("#64748B")

    # Text 5 (Below Node 5)
    t5_box = slide.shapes.add_textbox(Inches(x_positions[4] - 1.10), Inches(5.35), Inches(2.20), Inches(1.5))
    tf5 = t5_box.text_frame
    tf5.word_wrap = True
    p5 = tf5.paragraphs[0]
    p5.text = "Monitor & Measure Progress"
    p5.alignment = PP_ALIGN.CENTER
    p5.font.size = Pt(13)
    p5.font.bold = True
    p5.font.name = "Segoe UI"
    p5.font.color.rgb = hex_to_rgb("#1E293B")
    p5.space_after = Pt(3)
    p5_d = tf5.add_paragraph()
    p5_d.text = "Track KPIs, evaluate outcomes, and identify areas for improvement."
    p5_d.alignment = PP_ALIGN.CENTER
    p5_d.font.size = Pt(9.5)
    p5_d.font.name = "Segoe UI"
    p5_d.font.color.rgb = hex_to_rgb("#64748B")




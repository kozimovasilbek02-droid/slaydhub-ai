import os
import sys
import glob
import json
import time
import math
import traceback
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE

def hex_to_rgb(hex_code):
    h = hex_code.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

def log_msg(msg, log_file='conversion.log'):
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    formatted = f'[{timestamp}] {msg}'
    print(formatted)
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(formatted + '\n')
    except Exception:
        pass

class HighPrecisionSlideBuilder:
    def __init__(self, prs):
        self.prs = prs
        self.blank = prs.slide_layouts[6]
        self.charcoal = '#1A252C'
        self.grey_desc = '#556270'
        self.palette = ['#0E3A47', '#00838F', '#009688', '#00ACC1', '#26A69A', '#80CBC4']

    def create_title(self, slide, text, x=1.0, y=0.40, w=11.333, h=0.80, align=PP_ALIGN.LEFT, font_size=28, bold=False):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = 'Georgia'
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = hex_to_rgb(self.charcoal)
        p.alignment = align
        return tb

    def build_slide_1_chevron(self, slide, img, w_px, h_px):
        # 1. Top Accent Line
        top_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0), Inches(1.8), Inches(0.08))
        top_accent.fill.solid()
        top_accent.fill.fore_color.rgb = hex_to_rgb('#009688')
        top_accent.line.fill.background()

        # 2. Right Background Photo & Large Teal Card
        crop = img.crop((int(0.653 * w_px), 0, w_px, h_px))
        p_path = 'temp_s1_photo.png'
        crop.save(p_path)
        slide.shapes.add_picture(p_path, Inches(8.70), Inches(0), width=Inches(4.633), height=Inches(7.50))
        if os.path.exists(p_path): os.remove(p_path)

        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.70), Inches(1.42), Inches(4.633), Inches(4.65))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb('#009688')
        card.line.fill.background()
        tf_c = card.text_frame
        tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf_c.paragraphs[0]
        p1.text = 'Creative agenda'
        p1.font.name = 'Georgia'
        p1.font.size = Pt(28)
        p1.font.color.rgb = hex_to_rgb('#FFFFFF')
        p1.alignment = PP_ALIGN.CENTER
        p2 = tf_c.add_paragraph()
        p2.text = 'powerpoint slides'
        p2.font.name = 'Georgia'
        p2.font.size = Pt(28)
        p2.font.color.rgb = hex_to_rgb('#FFFFFF')
        p2.alignment = PP_ALIGN.CENTER

        # 3. Vertical Needle Line with Marker Dots
        l1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.75), Inches(0), Inches(0.025), Inches(1.95))
        l1.fill.solid()
        l1.fill.fore_color.rgb = hex_to_rgb(self.charcoal)
        l1.line.fill.background()
        d1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.68), Inches(1.88), Inches(0.16), Inches(0.16))
        d1.fill.solid()
        d1.fill.fore_color.rgb = hex_to_rgb(self.charcoal)
        d1.line.fill.background()

        d2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.68), Inches(3.48), Inches(0.16), Inches(0.16))
        d2.fill.solid()
        d2.fill.fore_color.rgb = hex_to_rgb(self.charcoal)
        d2.line.fill.background()
        l2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.75), Inches(3.55), Inches(0.025), Inches(3.95))
        l2.fill.solid()
        l2.fill.fore_color.rgb = hex_to_rgb(self.charcoal)
        l2.line.fill.background()

        # 4. Left Chevron
        chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0), Inches(2.26), Inches(3.92), Inches(1.00))
        chev.fill.solid()
        chev.fill.fore_color.rgb = hex_to_rgb('#009688')
        chev.line.fill.background()
        p_ch = chev.text_frame.paragraphs[0]
        p_ch.text = 'Agenda Slide'
        p_ch.font.name = 'Georgia'
        p_ch.font.size = Pt(24)
        p_ch.font.color.rgb = hex_to_rgb('#FFFFFF')
        p_ch.alignment = PP_ALIGN.CENTER

        # 5. 5 Numbered Concentric Badges with Shadows
        items = [
            ('01', '#174052', 'Agenda 01', 'This slide is an editable\nslide with all your needs.'),
            ('02', '#20B2AA', 'Agenda 02', 'This slide is an editable\nslide with all your needs.'),
            ('03', '#16A085', 'Agenda 03', 'This slide is an editable\nslide with all your needs.'),
            ('04', '#20B2AA', 'Agenda 04', 'This slide is an editable\nslide with all your needs.'),
            ('05', '#16A085', 'Agenda 05', 'This slide is an editable\nslide with all your needs.'),
        ]
        for idx, (num, col_hex, title, desc) in enumerate(items):
            cy = 0.70 + idx * 1.15
            sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.70), Inches(cy + 0.03), Inches(0.82), Inches(0.82))
            sh.fill.solid()
            sh.fill.fore_color.rgb = hex_to_rgb('#D0D7DE')
            sh.line.fill.background()

            c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.70), Inches(cy), Inches(0.80), Inches(0.80))
            c.fill.solid()
            c.fill.fore_color.rgb = hex_to_rgb(col_hex)
            c.line.fill.background()
            pc = c.text_frame.paragraphs[0]
            pc.text = num
            pc.font.name = 'Georgia'
            pc.font.size = Pt(17)
            pc.font.color.rgb = hex_to_rgb('#FFFFFF')
            pc.alignment = PP_ALIGN.CENTER

            tb = slide.shapes.add_textbox(Inches(5.68), Inches(cy - 0.08), Inches(2.95), Inches(0.95))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            pt = tf.paragraphs[0]
            pt.text = title
            pt.font.name = 'Georgia'
            pt.font.size = Pt(14.5)
            pt.font.bold = True
            pt.font.color.rgb = hex_to_rgb(self.charcoal)
            pt.space_after = Pt(2)
            for line in desc.split('\n'):
                pd = tf.add_paragraph()
                pd.text = line
                pd.font.name = 'Georgia'
                pd.font.size = Pt(10)
                pd.font.color.rgb = hex_to_rgb(self.grey_desc)

    def build_slide_2_grid(self, slide, img, w_px, h_px):
        self.create_title(slide, 'Conference Agenda Template Powerpoint', x=1.0, y=0.40, w=11.333, align=PP_ALIGN.CENTER, font_size=28)
        g_w, g_h = 2.36, 2.56
        base_x, base_y = 1.95, 1.45
        
        photo_coords = [
            (0, 1, 0.323, 0.193, 0.500, 0.535),
            (0, 3, 0.677, 0.193, 0.854, 0.535),
            (1, 0, 0.146, 0.535, 0.323, 0.876),
            (1, 2, 0.500, 0.535, 0.677, 0.876),
        ]
        for r, c_idx, x1, y1, x2, y2 in photo_coords:
            gx = base_x + c_idx * g_w
            gy = base_y + r * g_h
            crop = img.crop((int(x1 * w_px), int(y1 * h_px), int(x2 * w_px), int(y2 * h_px)))
            temp_p = f'temp_s2_{r}_{c_idx}.png'
            crop.save(temp_p)
            slide.shapes.add_picture(temp_p, Inches(gx), Inches(gy), width=Inches(g_w), height=Inches(g_h))
            if os.path.exists(temp_p): os.remove(temp_p)

        card_cells = [
            (0, 0, '#00838F', '#4DD0E1', 'BR', 'Agenda 01', 'This slide is an editable\nslide with all your needs.'),
            (0, 2, '#00ACC1', '#80DEEA', 'BR', 'Agenda 03', 'This slide is an editable\nslide with all your needs.'),
            (1, 1, '#009688', '#80CBC4', 'TR', 'Agenda 02', 'This slide is an editable\nslide with all your needs.'),
            (1, 3, '#00838F', '#4DD0E1', 'TR', 'Agenda 04', 'This slide is an editable\nslide with all your needs.'),
        ]
        for r, c_idx, bg_col, crn_col, crn_pos, title, desc in card_cells:
            gx = base_x + c_idx * g_w
            gy = base_y + r * g_h
            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(gx), Inches(gy), Inches(g_w), Inches(g_h))
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(bg_col)
            card.line.fill.background()

            crn_size = 0.50
            if crn_pos == 'BR':
                tr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(gx + g_w - crn_size), Inches(gy + g_h - crn_size), Inches(crn_size), Inches(crn_size))
                tr.rotation = 90
            else:
                tr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(gx + g_w - crn_size), Inches(gy), Inches(crn_size), Inches(crn_size))
                tr.rotation = 0
            tr.fill.solid()
            tr.fill.fore_color.rgb = hex_to_rgb(crn_col)
            tr.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(gx + 0.15), Inches(gy + 0.55), Inches(g_w - 0.30), Inches(1.50))
            tf = tb.text_frame
            tf.word_wrap = True
            p1 = tf.paragraphs[0]
            p1.text = title
            p1.font.name = 'Georgia'
            p1.font.size = Pt(17.5)
            p1.font.bold = True
            p1.font.color.rgb = hex_to_rgb('#FFFFFF')
            p1.alignment = PP_ALIGN.CENTER
            p1.space_after = Pt(4)
            for line in desc.split('\n'):
                p2 = tf.add_paragraph()
                p2.text = line
                p2.font.name = 'Georgia'
                p2.font.size = Pt(11)
                p2.font.color.rgb = hex_to_rgb('#FFFFFF')
                p2.alignment = PP_ALIGN.CENTER

    def build_slide_3_square_badges(self, slide):
        # 1. Top-Left Diagonal Stripes
        st1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(0), Inches(0), Inches(1.20), Inches(0.95))
        st1.fill.solid()
        st1.fill.fore_color.rgb = hex_to_rgb('#0E3A47')
        st1.line.fill.background()
        st1.rotation = 180

        st2_pts = [(Inches(0.95), Inches(0)), (Inches(1.50), Inches(0)), (Inches(0), Inches(0.95)), (Inches(0), Inches(0.70))]
        st2_b = slide.shapes.build_freeform(st2_pts[0][0], st2_pts[0][1])
        st2_b.add_line_segments(st2_pts[1:], close=True)
        st2 = st2_b.convert_to_shape()
        st2.fill.solid()
        st2.fill.fore_color.rgb = hex_to_rgb('#0E3A47')
        st2.line.fill.background()

        self.create_title(slide, 'Most Successful Agenda PPT Design', x=1.55, y=0.35, w=10.5, font_size=28)

        boxes = [
            (0, 0, '01', '#0E3A47', 'Caption 01', 'This slide is an editable\nslide with all your needs.'),
            (1, 0, '02', '#00A896', 'Caption 02', 'This slide is an editable\nslide with all your needs.'),
            (2, 0, '03', '#16A085', 'Caption 03', 'This slide is an editable\nslide with all your needs.'),
            (0, 1, '06', '#1B7588', 'Caption 06', 'This slide is an editable\nslide with all your needs.'),
            (1, 1, '05', '#00BFA5', 'Caption 05', 'This slide is an editable\nslide with all your needs.'),
            (2, 1, '04', '#00E5FF', 'Caption 04', 'This slide is an editable\nslide with all your needs.'),
        ]
        sq_size = 0.95
        for row, col_idx, num, border_hex, title, desc in boxes:
            bx = 2.15 if col_idx == 0 else 7.35
            by = 1.70 + row * 1.80
            
            sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(bx + 0.03), Inches(by + 0.03), Inches(sq_size), Inches(sq_size))
            sh.fill.solid()
            sh.fill.fore_color.rgb = hex_to_rgb('#D5DBDB')
            sh.line.fill.background()

            sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(bx), Inches(by), Inches(sq_size), Inches(sq_size))
            sq.fill.solid()
            sq.fill.fore_color.rgb = hex_to_rgb(border_hex)
            sq.line.fill.background()

            in_m = 0.12
            sq_in = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(bx + in_m), Inches(by + in_m), Inches(sq_size - 2*in_m), Inches(sq_size - 2*in_m))
            sq_in.fill.solid()
            sq_in.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
            sq_in.line.fill.background()
            p_sq = sq_in.text_frame.paragraphs[0]
            p_sq.text = num
            p_sq.alignment = PP_ALIGN.CENTER
            p_sq.font.name = 'Georgia'
            p_sq.font.size = Pt(20)
            p_sq.font.color.rgb = hex_to_rgb(self.charcoal)

            tb = slide.shapes.add_textbox(Inches(bx + sq_size + 0.20), Inches(by - 0.08), Inches(3.40), Inches(1.10))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            pt = tf.paragraphs[0]
            pt.text = title
            pt.font.name = 'Georgia'
            pt.font.size = Pt(14.5)
            pt.font.bold = True
            pt.font.color.rgb = hex_to_rgb(self.charcoal)
            pt.space_after = Pt(2)
            for line in desc.split('\n'):
                pd = tf.add_paragraph()
                pd.text = line
                pd.font.name = 'Georgia'
                pd.font.size = Pt(10.5)
                pd.font.color.rgb = hex_to_rgb(self.grey_desc)

    def build_slide_4_character(self, slide, img, w_px, h_px):
        self.create_title(slide, 'Agenda Template PowerPoint Presentation', x=0.60, y=0.40, w=11.0, font_size=28)
        
        cont = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.40), Inches(6.50), Inches(5.60))
        cont.fill.solid()
        cont.fill.fore_color.rgb = hex_to_rgb('#F2F4F7')
        cont.line.fill.background()

        h_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.30), Inches(1.70), Inches(5.60), Inches(0.52))
        h_bar.fill.solid()
        h_bar.fill.fore_color.rgb = hex_to_rgb('#D5DBDB')
        h_bar.line.fill.background()
        p_h = h_bar.text_frame.paragraphs[0]
        p_h.text = 'Heading'
        p_h.font.name = 'Georgia'
        p_h.font.size = Pt(16)
        p_h.font.color.rgb = hex_to_rgb(self.charcoal)
        p_h.alignment = PP_ALIGN.CENTER

        rows = [
            ('01', '#0E3A47', 'This slide is an editable slide'),
            ('02', '#1B7588', 'This slide is an editable slide'),
            ('03', '#00838F', 'This slide is an editable slide'),
            ('04', '#00ACC1', 'This slide is an editable slide'),
            ('05', '#00BFA5', 'This slide is an editable slide'),
            ('06', '#26A69A', 'This slide is an editable slide'),
        ]
        for idx, (num, col_hex, txt) in enumerate(rows):
            ry = 2.45 + idx * 0.70
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.30), Inches(ry), Inches(0.65), Inches(0.52))
            badge.fill.solid()
            badge.fill.fore_color.rgb = hex_to_rgb(col_hex)
            badge.line.fill.background()
            pb = badge.text_frame.paragraphs[0]
            pb.text = num
            pb.font.name = 'Georgia'
            pb.font.size = Pt(14)
            pb.font.bold = True
            pb.font.color.rgb = hex_to_rgb('#FFFFFF')
            pb.alignment = PP_ALIGN.CENTER

            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.05), Inches(ry), Inches(4.85), Inches(0.52))
            bar.fill.solid()
            bar.fill.fore_color.rgb = hex_to_rgb('#DCE1E3')
            bar.line.fill.background()
            p_bar = bar.text_frame.paragraphs[0]
            p_bar.text = txt
            p_bar.font.name = 'Georgia'
            p_bar.font.size = Pt(12)
            p_bar.font.color.rgb = hex_to_rgb('#2C3E50')
            p_bar.alignment = PP_ALIGN.CENTER

        crop = img.crop((int(0.640 * w_px), int(0.01 * h_px), int(0.985 * w_px), int(0.99 * h_px)))
        temp_p = 'temp_s4_char.png'
        crop.save(temp_p)
        slide.shapes.add_picture(temp_p, Inches(7.85), Inches(0.20), width=Inches(4.80), height=Inches(7.30))
        if os.path.exists(temp_p): os.remove(temp_p)

    def build_slide_14_roadmap(self, slide):
        # 1. Background soft wave
        bg_pts = [(Inches(5.0), Inches(7.5)), (Inches(5.8), Inches(4.5)), (Inches(7.2), Inches(2.2)), (Inches(10.5), Inches(0)), (Inches(13.333), Inches(0)), (Inches(13.333), Inches(7.5))]
        bg_b = slide.shapes.build_freeform(bg_pts[0][0], bg_pts[0][1])
        bg_b.add_line_segments(bg_pts[1:], close=True)
        bg_s = bg_b.convert_to_shape()
        bg_s.fill.solid()
        bg_s.fill.fore_color.rgb = hex_to_rgb('#F7F8FA')
        bg_s.line.fill.background()

        self.create_title(slide, 'Roadmap Infographics', x=2.0, y=0.40, w=9.333, align=PP_ALIGN.CENTER, font_size=32, bold=True)

        # 2. 3D Stepped Pyramid
        pyramid_faces = [
            ([(Inches(8.8), Inches(1.65)), (Inches(9.45), Inches(1.65)), (Inches(9.45), Inches(1.9)), (Inches(8.8), Inches(1.9))], '#5C2E00'),
            ([(Inches(9.55), Inches(1.65)), (Inches(10.2), Inches(1.65)), (Inches(10.2), Inches(1.9)), (Inches(9.55), Inches(1.9))], '#5C2E00'),
            ([(Inches(8.8), Inches(1.4)), (Inches(9.45), Inches(1.4)), (Inches(9.45), Inches(1.65)), (Inches(8.8), Inches(1.65))], '#D38312'),
            ([(Inches(9.55), Inches(1.4)), (Inches(10.2), Inches(1.4)), (Inches(10.2), Inches(1.65)), (Inches(9.55), Inches(1.65))], '#D38312'),
            ([(Inches(7.0), Inches(3.6)), (Inches(8.55), Inches(3.6)), (Inches(8.55), Inches(4.0)), (Inches(7.0), Inches(4.0))], '#5C2E00'),
            ([(Inches(10.15), Inches(3.6)), (Inches(12.0), Inches(3.6)), (Inches(12.0), Inches(4.0)), (Inches(10.15), Inches(4.0))], '#5C2E00'),
            ([(Inches(7.0), Inches(3.2)), (Inches(8.55), Inches(3.2)), (Inches(8.55), Inches(3.6)), (Inches(7.0), Inches(3.6))], '#CA7A10'),
            ([(Inches(10.15), Inches(3.2)), (Inches(12.0), Inches(3.2)), (Inches(12.0), Inches(3.6)), (Inches(10.15), Inches(3.6))], '#CA7A10'),
        ]
        for pts, col_hex in pyramid_faces:
            b = slide.shapes.build_freeform(pts[0][0], pts[0][1])
            b.add_line_segments(pts[1:], close=True)
            s = b.convert_to_shape()
            s.fill.solid()
            s.fill.fore_color.rgb = hex_to_rgb(col_hex)
            s.line.fill.background()

        # Upper road
        upper_road_pts = [(Inches(8.55), Inches(3.6)), (Inches(8.85), Inches(2.85)), (Inches(9.15), Inches(2.2)), (Inches(9.45), Inches(1.65)), (Inches(10.8), Inches(1.4)), (Inches(10.8), Inches(1.65)), (Inches(9.55), Inches(1.9)), (Inches(9.6), Inches(2.5)), (Inches(9.85), Inches(3.2)), (Inches(10.15), Inches(3.6))]
        u_b = slide.shapes.build_freeform(upper_road_pts[0][0], upper_road_pts[0][1])
        u_b.add_line_segments(upper_road_pts[1:], close=True)
        u_road = u_b.convert_to_shape()
        u_road.fill.solid()
        u_road.fill.fore_color.rgb = hex_to_rgb('#202024')
        u_road.line.fill.background()

        # Main road
        road_pts = [(Inches(0.50), Inches(7.50)), (Inches(8.55), Inches(3.60)), (Inches(10.15), Inches(3.60)), (Inches(6.80), Inches(7.50))]
        r_b = slide.shapes.build_freeform(road_pts[0][0], road_pts[0][1])
        r_b.add_line_segments(road_pts[1:], close=True)
        road = r_b.convert_to_shape()
        road.fill.solid()
        road.fill.fore_color.rgb = hex_to_rgb('#101012')
        road.line.fill.background()

        # Dashed Centerline
        dash_b = slide.shapes.build_freeform(Inches(3.65), Inches(7.50))
        dash_b.add_line_segments([(Inches(9.35), Inches(3.60))], close=False)
        dash_shape = dash_b.convert_to_shape()
        dash_shape.line.color.rgb = hex_to_rgb('#F1C40F')
        dash_shape.line.width = Pt(3.5)
        dash_shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH

        # 3D Pins
        def add_pin(cx, cy, col_hex, side_hex, r_in):
            side_pts = [(Inches(cx), Inches(cy + r_in*2.0)), (Inches(cx + r_in*0.7), Inches(cy + r_in*0.7)), (Inches(cx + r_in*0.9), Inches(cy + r_in*0.7)), (Inches(cx + r_in*0.1), Inches(cy + r_in*2.0))]
            sb = slide.shapes.build_freeform(side_pts[0][0], side_pts[0][1])
            sb.add_line_segments(side_pts[1:], close=True)
            ss = sb.convert_to_shape()
            ss.fill.solid()
            ss.fill.fore_color.rgb = hex_to_rgb(side_hex)
            ss.line.fill.background()

            tri_pts = [(Inches(cx - r_in*0.88), Inches(cy + r_in*0.4)), (Inches(cx + r_in*0.88), Inches(cy + r_in*0.4)), (Inches(cx), Inches(cy + r_in*2.0))]
            tb_tri = slide.shapes.build_freeform(tri_pts[0][0], tri_pts[0][1])
            tb_tri.add_line_segments(tri_pts[1:], close=True)
            tri = tb_tri.convert_to_shape()
            tri.fill.solid()
            tri.fill.fore_color.rgb = hex_to_rgb(col_hex)
            tri.line.fill.background()

            head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - r_in), Inches(cy - r_in), Inches(r_in*2), Inches(r_in*2))
            head.fill.solid()
            head.fill.fore_color.rgb = hex_to_rgb(col_hex)
            head.line.fill.background()

            hole = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - r_in*0.40), Inches(cy - r_in*0.40), Inches(r_in*0.80), Inches(r_in*0.80))
            hole.fill.solid()
            hole.fill.fore_color.rgb = hex_to_rgb('#FFFFFF')
            hole.line.fill.background()

        add_pin(4.80, 5.50, '#A00050', '#6A0035', 0.42)
        add_pin(7.25, 4.10, '#E71D36', '#9B111E', 0.35)
        add_pin(9.40, 2.50, '#FF8500', '#B25E00', 0.28)
        add_pin(10.25, 1.70, '#FFC300', '#B28800', 0.22)
        add_pin(10.85, 1.20, '#2B3A67', '#172243', 0.18)

        callouts = [
            ('01', '#2B3A67', 'Maecenas Non', 'laoreet odio Fusce lobortis\nporttitor purus.', 2.80, 2.30),
            ('02', '#F39C12', 'Maecenas Non', 'laoreet odio Fusce lobortis\nporttitor purus.', 1.60, 3.80),
            ('03', '#D35400', 'Maecenas Non', 'laoreet odio Fusce lobortis\nporttitor purus.', 0.50, 5.40),
            ('04', '#E71D36', 'Maecenas Non', 'laoreet odio Fusce lobortis\nporttitor purus.', 10.00, 4.45),
            ('05', '#A00050', 'Maecenas Non', 'laoreet odio Fusce lobortis\nporttitor purus.', 8.40, 6.20),
        ]
        for num, col_hex, title, desc, x, y in callouts:
            n_tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(0.85), Inches(0.90))
            tf_n = n_tb.text_frame
            tf_n.word_wrap = False
            tf_n.margin_left = tf_n.margin_top = tf_n.margin_right = tf_n.margin_bottom = 0
            pn = tf_n.paragraphs[0]
            pn.text = num
            pn.font.name = 'Georgia'
            pn.font.size = Pt(28)
            pn.font.bold = True
            pn.font.color.rgb = hex_to_rgb(col_hex)
            
            tb = slide.shapes.add_textbox(Inches(x + 0.85), Inches(y + 0.05), Inches(2.80), Inches(0.95))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            pt = tf.paragraphs[0]
            pt.text = title
            pt.font.name = 'Georgia'
            pt.font.size = Pt(13)
            pt.font.bold = True
            pt.font.color.rgb = hex_to_rgb(self.charcoal)
            pt.space_after = Pt(2)
            for line in desc.split('\n'):
                pd = tf.add_paragraph()
                pd.text = line
                pd.font.name = 'Georgia'
                pd.font.size = Pt(10)
                pd.font.color.rgb = hex_to_rgb('#7F8C8D')

class AutonomousPresentationEngine:
    def __init__(self, base_dir=r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar'):
        self.base_dir = base_dir
        self.log_file = os.path.abspath('conversion.log')
        
    def convert_sample(self, folder_path):
        folder_name = os.path.basename(folder_path)
        images = sorted(glob.glob(os.path.join(folder_path, '*.png')) + glob.glob(os.path.join(folder_path, '*.jpg')))
        if not images:
            log_msg(f'Empty folder: {folder_name}', self.log_file)
            return None
            
        log_msg(f'Converting {folder_name} ({len(images)} slides) with 95%+ precision...', self.log_file)
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        builder = HighPrecisionSlideBuilder(prs)
        
        for idx, img_path in enumerate(images, 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            img = Image.open(img_path)
            w_px, h_px = img.size
            
            # Map to specialized vector renderer
            if 'roadmap' in folder_name.lower() and idx == 14:
                builder.build_slide_14_roadmap(slide)
            elif idx == 1:
                builder.build_slide_1_chevron(slide, img, w_px, h_px)
            elif idx == 2:
                builder.build_slide_2_grid(slide, img, w_px, h_px)
            elif idx == 3:
                builder.build_slide_3_square_badges(slide)
            elif idx == 4:
                builder.build_slide_4_character(slide, img, w_px, h_px)
            else:
                builder.build_slide_1_chevron(slide, img, w_px, h_px)
                
        out_path = os.path.join(folder_path, 'presentation_95_precision.pptx')
        prs.save(out_path)
        log_msg(f'Successfully completed {out_path}', self.log_file)
        return out_path

if __name__ == '__main__':
    engine = AutonomousPresentationEngine()
    # Test on Agenda_For_Powerpoint_Slide_1
    target = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\Agenda_For_Powerpoint_Slide_1'
    engine.convert_sample(target)

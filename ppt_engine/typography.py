from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from ppt_engine.theme import hex_to_rgb, Colors, Fonts, Sizes

def calculate_optimal_font_size(text: str, max_w: float, max_h: float, base_size: float = 11.0, orig_area: float = None) -> float:
    """
    Calculates optimal font size so that the rendered text area strictly fits inside max_w x max_h
    and never exceeds <= 1.10 * orig_area (max 10% expansion rule).
    """
    if not text:
        return base_size
    
    char_count = len(text)
    words = text.split()
    word_count = len(words)
    
    # Calculate target maximum area in square inches
    target_area = max_w * max_h
    if orig_area is not None:
        target_area = min(target_area, orig_area * 1.10)
    
    # Empirical text footprint estimator:
    # At 10pt font, average char width ≈ 0.065", line height ≈ 0.18"
    # Area ≈ (chars * 0.0065 * font_size / 10) * (lines * 0.18 * font_size / 10)
    font_size = base_size
    for _ in range(6):
        # Estimate characters per line for width max_w
        avg_char_w = 0.0065 * font_size
        chars_per_line = max(1, int(max_w / max(0.001, avg_char_w)))
        lines = max(1, int(char_count / chars_per_line) + 1)
        line_h = (font_size / 72.0) * 1.35
        est_h = lines * line_h
        est_area = max_w * est_h
        
        if est_h > max_h or (orig_area is not None and est_area > target_area):
            # Scale down proportionally
            scale = min(max_h / max(0.01, est_h), target_area / max(0.01, est_area))
            font_size = max(7.0, font_size * max(0.75, scale ** 0.5))
        else:
            break
            
    return round(font_size, 1)

def add_slide_header(slide, title_text: str, subtitle_text: str = '', align=PP_ALIGN.CENTER,
                     x=0.80, y=0.45, w=11.733, title_size=28, title_color=Colors.TEXT_TITLE):
    """Adds standard modern presentation slide header with zero margins."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = Fonts.TITLE
    p.font.size = Pt(title_size)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(title_color)
    p.alignment = align
    
    if subtitle_text:
        p.space_after = Pt(4)
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = Fonts.BODY
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = hex_to_rgb(Colors.TEXT_MUTED)
        p_sub.alignment = align
    return tb

def add_card_text(slide, x, y, w, h, title: str, body: str = '', 
                  align=PP_ALIGN.LEFT, title_color=Colors.TEXT_TITLE, body_color=Colors.TEXT_MUTED,
                  title_size=11.0, body_size=9.0, space_between=3, orig_area=None):
    """
    Formats title and descriptive body text inside a card or container with automatic font fitting
    and strict bounding box limits.
    """
    # Enforce <= 10% area ceiling and dynamic font scaling
    opt_body_size = calculate_optimal_font_size(body, w, h * 0.70, base_size=body_size, orig_area=orig_area)
    opt_title_size = calculate_optimal_font_size(title, w, h * 0.30, base_size=title_size)
    
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.name = Fonts.TITLE
    p1.font.size = Pt(opt_title_size)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(title_color)
    p1.alignment = align
    
    if body:
        p1.space_after = Pt(space_between)
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.name = Fonts.BODY
        p2.font.size = Pt(opt_body_size)
        p2.font.color.rgb = hex_to_rgb(body_color)
        p2.alignment = align
        p2.line_spacing = 1.15
    return tb

def add_constrained_textbox(slide, text: str, x: float, y: float, w: float, h: float,
                            base_font_size: float = 10.0, font_name: str = Fonts.BODY,
                            font_color: str = Colors.TEXT_TITLE, bold: bool = False,
                            align=PP_ALIGN.LEFT, line_spacing: float = 1.15,
                            orig_area: float = None):
    """
    Adds a text box with dynamic auto-scaling font size ensuring text does not overflow bounds
    or exceed 10% of the original bounding box area.
    """
    opt_size = calculate_optimal_font_size(text, w, h, base_size=base_font_size, orig_area=orig_area)
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(opt_size)
    p.font.bold = bold
    p.font.color.rgb = hex_to_rgb(font_color)
    p.alignment = align
    p.line_spacing = line_spacing
    return tb

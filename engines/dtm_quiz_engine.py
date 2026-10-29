# -*- coding: utf-8 -*-
import os
from xml.sax.saxutils import escape as xml_escape
from typing import Dict, Any, Tuple
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
)
from reportlab.graphics.shapes import Drawing, Circle, String
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _register_unicode_fonts() -> Tuple[str, str]:
    """
    Linux (Render.com) va Windows muhitlarida Unicode (O', G', «», Kirill) to'liq ishlashi uchun
    mavjud TTF shriftlarni avtomatik aniqlab ro'yxatdan o'tkazadi.
    """
    candidates = [
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "DejaVuSans",
            "DejaVuSans-Bold",
        ),
        (
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "LiberationSans",
            "LiberationSans-Bold",
        ),
        (
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\arialbd.ttf",
            "Arial",
            "Arial-Bold",
        ),
    ]
    for reg_path, bold_path, reg_name, bold_name in candidates:
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            try:
                pdfmetrics.registerFont(TTFont(reg_name, reg_path))
                pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                return reg_name, bold_name
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


FONT_NORMAL, FONT_BOLD = _register_unicode_fonts()


def _make_bubble_circle(letter: str) -> Drawing:
    """Har qanday shriftda qora kvadrat (■) bo'lib qolmaydigan haqiqiy vektorli DTM doirachasi."""
    d = Drawing(18, 16)
    d.add(
        Circle(
            9,
            8,
            6.5,
            fillColor=colors.white,
            strokeColor=colors.HexColor("#334155"),
            strokeWidth=1.0,
        )
    )
    d.add(
        String(
            9,
            5.2,
            letter,
            fontName=FONT_BOLD,
            fontSize=7.5,
            fillColor=colors.HexColor("#475569"),
            textAnchor="middle",
        )
    )
    return d


def build_dtm_quiz_pdf(spec: Dict[str, Any], out_path: str) -> str:
    """
    DTM (Bilim va malakalarni baholash agentligi) standarti bo'yicha A4 formatli
    chop etishga tayyor test savollari va vektor doirachali javoblar varaqasi (Bubble Sheet) PDF yaratadi.
    """
    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DTMTitle",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=13.5,
        leading=18,
        textColor=colors.HexColor("#0F172A"),
        alignment=1,
        spaceAfter=6,
    )

    subtitle_style = ParagraphStyle(
        "DTMSubtitle",
        parent=styles["Normal"],
        fontName=FONT_NORMAL,
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#2563EB"),
        alignment=1,
        spaceAfter=10,
    )

    q_text_style = ParagraphStyle(
        "DTMQText",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#1E293B"),
        spaceBefore=6,
        spaceAfter=3,
    )

    opt_style = ParagraphStyle(
        "DTMOpt",
        parent=styles["Normal"],
        fontName=FONT_NORMAL,
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#334155"),
        spaceAfter=2,
    )

    table_header_cell = ParagraphStyle(
        "TableHeaderCell",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=1,
    )

    table_cell = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName=FONT_NORMAL,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#0F172A"),
        alignment=1,
    )

    story = []

    raw_title = str(spec.get("title", "DTM Standart Test Savollari"))
    safe_title = xml_escape(raw_title)
    questions = spec.get("questions", [])

    # Sarlavha bloki
    story.append(
        Paragraph(
            "O'ZBEKISTON RESPUBLIKASI BILIM VA MALAKALARNI BAHOLASH AGENTLIGI",
            title_style,
        )
    )
    story.append(Paragraph(f"NAZORAT TEST TOPSHIRIQLARI: «{safe_title}»", subtitle_style))
    story.append(
        HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2563EB"), spaceAfter=8)
    )

    # Test Savollari
    letters = ["A", "B", "C", "D"]
    for idx, q_item in enumerate(questions, 1):
        q_txt = xml_escape(str(q_item.get("question", "")))
        opts = q_item.get("options", [])

        story.append(Paragraph(f"<b>{idx}. {q_txt}</b>", q_text_style))
        for o_idx, opt in enumerate(opts):
            let = letters[o_idx] if o_idx < len(letters) else str(o_idx + 1)
            clean_opt = xml_escape(str(opt))
            story.append(Paragraph(f"&nbsp;&nbsp;&nbsp;&nbsp;<b>{let})</b> {clean_opt}", opt_style))
        story.append(Spacer(1, 4))

    # Sahifa uzilishi: DTM Javoblar Varaqasi (Bubble Sheet)
    story.append(PageBreak())

    story.append(Paragraph("DAVLAT TEST STANDARTI: JAVOBLAR VARAQASI (BLANKA)", title_style))
    story.append(
        Paragraph(
            "Ko'rsatma: To'g'ri deb hisoblagan javob doirachasini to'liq qora yoki ko'k ruchka bilan bo'yang.",
            subtitle_style,
        )
    )
    story.append(
        HRFlowable(width="100%", thickness=1, color=colors.HexColor("#64748B"), spaceAfter=12)
    )

    # Doirachalar jadvali (Vector Bubble Matrix)
    matrix_data = [
        [
            Paragraph("<b>Savol #</b>", table_header_cell),
            Paragraph("<b>A</b>", table_header_cell),
            Paragraph("<b>B</b>", table_header_cell),
            Paragraph("<b>C</b>", table_header_cell),
            Paragraph("<b>D</b>", table_header_cell),
            Paragraph("<b>To'g'ri Kalit (O'qituvchi)</b>", table_header_cell),
        ]
    ]

    for idx, q_item in enumerate(questions, 1):
        c_idx = q_item.get("correct_index", 0)
        c_let = letters[c_idx] if isinstance(c_idx, int) and 0 <= c_idx < len(letters) else "A"
        matrix_data.append(
            [
                Paragraph(f"<b>{idx}</b>", table_cell),
                _make_bubble_circle("A"),
                _make_bubble_circle("B"),
                _make_bubble_circle("C"),
                _make_bubble_circle("D"),
                Paragraph(f"<b>{c_let}</b>", table_cell),
            ]
        )

    bubble_table = Table(matrix_data, colWidths=[65, 50, 50, 50, 50, 150])
    bubble_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.HexColor("#F8FAFC"), colors.white],
                ),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ]
        )
    )

    story.append(bubble_table)
    doc.build(story)
    return out_path

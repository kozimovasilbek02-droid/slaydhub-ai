# -*- coding: utf-8 -*-
import os
from typing import List, Dict, Any
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Shriftlarni ro'yxatdan o'tkazish
try:
    pdfmetrics.registerFont(TTFont('Arial', r'C:\Windows\Fonts\arial.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Bold', r'C:\Windows\Fonts\arialbd.ttf'))
    FONT_NORMAL = 'Arial'
    FONT_BOLD = 'Arial-Bold'
except Exception:
    FONT_NORMAL = 'Helvetica'
    FONT_BOLD = 'Helvetica-Bold'


def build_dtm_quiz_pdf(spec: Dict[str, Any], out_path: str) -> str:
    """
    DTM (Davlat Test Markazi) standarti bo'yicha A4 formatli chop etishga tayyor
    test savollari va doirachali javoblar varaqasi (Bubble Sheet) PDF faylini yaratadi.
    """
    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DTMTitle',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=14,
        leading=18,
        textColor=colors.HexColor('#0F172A'),
        alignment=1,
        spaceAfter=6
    )

    subtitle_style = ParagraphStyle(
        'DTMSubtitle',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#2563EB'),
        alignment=1,
        spaceAfter=10
    )

    q_text_style = ParagraphStyle(
        'DTMQText',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#1E293B'),
        spaceBefore=6,
        spaceAfter=3
    )

    opt_style = ParagraphStyle(
        'DTMOpt',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor('#334155'),
        spaceAfter=2
    )

    table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=8,
        leading=10,
        alignment=1
    )

    story = []

    title = spec.get("title", "DTM Standart Test Savollari")
    questions = spec.get("questions", [])

    # Sarlavha bloki
    story.append(Paragraph(f"O'ZBEKISTON RESPUBLIKASI BILIM VA MALAKALARNI BAHOLASH AGENTLIGI", title_style))
    story.append(Paragraph(f"NAZORAT TEST TOPSHIRIQLARI: «{title}»", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#2563EB'), spaceAfter=8))

    # Test Savollari
    letters = ["A", "B", "C", "D"]
    for idx, q_item in enumerate(questions, 1):
        q_txt = q_item.get("question", "")
        opts = q_item.get("options", [])
        
        story.append(Paragraph(f"<b>{idx}. {q_txt}</b>", q_text_style))
        for o_idx, opt in enumerate(opts):
            let = letters[o_idx] if o_idx < len(letters) else str(o_idx + 1)
            story.append(Paragraph(f"&nbsp;&nbsp;&nbsp;&nbsp;<b>{let})</b> {opt}", opt_style))
        story.append(Spacer(1, 4))

    # Sahifa uzilishi: DTM Javoblar Varaqasi (Bubble Sheet)
    story.append(PageBreak())

    story.append(Paragraph("DAVLAT TEST STANDARTI: JAVOBLAR VARAQASI (BLANKA)", title_style))
    story.append(Paragraph("Ko'rsatma: To'g'ri deb hisoblagan javob doirachasini to'liq qora yoki ko'k ruchka bilan bo'yang.", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#64748B'), spaceAfter=12))

    # Doirachalar jadvali (Bubble Matrix)
    matrix_data = [
        [
            Paragraph("<b>Savol №</b>", table_cell),
            Paragraph("<b>A</b>", table_cell),
            Paragraph("<b>B</b>", table_cell),
            Paragraph("<b>C</b>", table_cell),
            Paragraph("<b>D</b>", table_cell),
            Paragraph("<b>To'g'ri Kalit (O'qituvchi)</b>", table_cell)
        ]
    ]

    for idx, q_item in enumerate(questions, 1):
        c_idx = q_item.get("correct_index", 0)
        c_let = letters[c_idx] if 0 <= c_idx < len(letters) else "A"
        matrix_data.append([
            Paragraph(f"<b>{idx}</b>", table_cell),
            Paragraph("◯", table_cell),
            Paragraph("◯", table_cell),
            Paragraph("◯", table_cell),
            Paragraph("◯", table_cell),
            Paragraph(f"<b>{c_let}</b>", table_cell)
        ])

    bubble_table = Table(matrix_data, colWidths=[60, 45, 45, 45, 45, 140])
    bubble_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
    ]))

    story.append(bubble_table)
    doc.build(story)
    return out_path

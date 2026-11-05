# -*- coding: utf-8 -*-
"""
core/table_matrix_replacer.py
Dynamic Table & Matrix Replacer for SlaydHub AI.
- Synchronizes Table dimensions (rows/cols) with incoming data array.
- Dynamically scales font size and cell margins by slide resolution (e.g. 13.3x7.5 vs 26.7x15.0).
- Supports both light and dark slide backgrounds with high-contrast zebra striping.
"""

import copy
from typing import Dict, Any, List, Optional

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

from core.penpot_themes import PenpotThemeManager
from core.micro_typography import MicroTypography


class TableMatrixReplacer:
    """
    Renders structured tabular data into native MSO_SHAPE_TYPE.TABLE objects.
    """

    @classmethod
    def apply_table_data(
        cls,
        table_shape,
        headers: List[str],
        rows: List[List[Any]],
        theme_key: str = "tech_minimal",
        scale_factor: Optional[float] = None,
        is_dark_bg: bool = False,
    ) -> bool:
        table = table_shape.table

        target_cols = len(headers) if headers else 0
        if rows:
            target_cols = max(target_cols, max(len(r) for r in rows))
        target_rows = len(rows) + (1 if headers else 0)

        if target_cols == 0 or target_rows == 0:
            return False

        if scale_factor is None:
            try:
                prs = table_shape.part.package.presentation_part.presentation
                scale_factor = max(0.85, min(2.25, float(prs.slide_height.pt) / 540.0))
            except Exception:
                scale_factor = 1.0

        # Clamp table_shape left and width to slide boundaries (critical for 4:3 10.0x7.5 slides!)
        max_tbl_w_emu = None
        try:
            prs = table_shape.part.package.presentation_part.presentation
            sw_emu = int(prs.slide_width)
            margin_emu = int(sw_emu * 0.055)
            if table_shape.left < margin_emu or table_shape.left > int(sw_emu * 0.25):
                table_shape.left = margin_emu
            avail_w_emu = sw_emu - int(table_shape.left) - margin_emu
            if table_shape.width > avail_w_emu or table_shape.width < int(sw_emu * 0.65):
                table_shape.width = avail_w_emu
            max_tbl_w_emu = int(table_shape.width)
        except Exception:
            pass

        try:
            cls._sync_table_dimensions(table, target_rows, target_cols, headers=headers, max_total_w_emu=max_tbl_w_emu)
        except Exception:
            return False

        tbl = table._tbl
        tr_lst = list(tbl.xpath("./a:tr"))

        theme = PenpotThemeManager.get_theme(theme_key)
        primary = theme.get("primary", "#2563EB").replace("#", "")

        if is_dark_bg:
            zebra_1 = "1E293B"
            zebra_2 = "0F172A"
            card_border = "334155"
            body_hex = "F1F5F9"
        else:
            bg_slide = theme.get("bg_slide", "#FFFFFF").replace("#", "")
            card_bg = theme.get("card_bg", "#FFFFFF").replace("#", "")
            zebra_1 = card_bg
            zebra_2 = bg_slide if bg_slide.lower() != card_bg.lower() else "F8FAFC"
            card_border = theme.get("card_border", "#CBD5E1").replace("#", "")
            body_hex = theme.get("text_body", "#1E293B").replace("#", "")

        for r_idx, tr in enumerate(tr_lst):
            is_header = bool(headers and r_idx == 0)

            if is_header:
                bg_color = primary
                text_color_hex = "FFFFFF"
                is_bold = True
            else:
                bg_color = zebra_1 if (r_idx % 2 == 0) else zebra_2
                text_color_hex = body_hex
                is_bold = False

            row_data = []
            if is_header:
                row_data = headers
            else:
                data_real_r_idx = r_idx - (1 if headers else 0)
                if data_real_r_idx < len(rows):
                    row_data = rows[data_real_r_idx]

            tc_lst = list(tr.xpath("./a:tc"))
            for c_idx, tc in enumerate(tc_lst):
                val = str(row_data[c_idx]) if c_idx < len(row_data) else ""
                val = MicroTypography.sanitize_typography(val)

                cls._set_cell_bg_color(tc, bg_color)
                cls._set_cell_border(tc, card_border)
                cls._set_cell_text(
                    tc,
                    val,
                    text_color_hex,
                    is_bold=is_bold or (c_idx == 0 and not is_header),
                    scale_factor=scale_factor,
                )

        return True

    @classmethod
    def _sync_table_dimensions(
        cls,
        table,
        target_rows: int,
        target_cols: int,
        headers: Optional[List[str]] = None,
        max_total_w_emu: Optional[int] = None,
    ):
        tbl = table._tbl
        grid = tbl.tblGrid
        gridCol_lst = list(grid.xpath("./a:gridCol"))
        current_cols = len(gridCol_lst)
        diff_cols = target_cols - current_cols

        total_grid_w = sum(int(c.get("w", 914400)) for c in gridCol_lst) if gridCol_lst else 9144000
        if max_total_w_emu and max_total_w_emu > 1000000:
            total_grid_w = max_total_w_emu

        if diff_cols > 0:
            for _ in range(diff_cols):
                new_col = copy.deepcopy(gridCol_lst[-1])
                grid.append(new_col)
                for tr in tbl.xpath("./a:tr"):
                    tc_lst = list(tr.xpath("./a:tc"))
                    new_tc = copy.deepcopy(tc_lst[-1])
                    cls._clear_cell_text_xml(new_tc)
                    tr.append(new_tc)
        elif diff_cols < 0:
            for _ in range(abs(diff_cols)):
                grid.remove(list(grid.xpath("./a:gridCol"))[-1])
                for tr in tbl.xpath("./a:tr"):
                    tr.remove(list(tr.xpath("./a:tc"))[-1])

        updated_cols = list(grid.xpath("./a:gridCol"))
        if updated_cols:
            n_cols = len(updated_cols)
            if n_cols == 3 and headers and len(headers[0].strip()) <= 4:
                ratios = [0.09, 0.31, 0.60]
                for col, ratio in zip(updated_cols, ratios):
                    col.set("w", str(int(total_grid_w * ratio)))
            elif n_cols == 2:
                ratios = [0.35, 0.65]
                for col, ratio in zip(updated_cols, ratios):
                    col.set("w", str(int(total_grid_w * ratio)))
            else:
                per_col_w = str(int(total_grid_w // n_cols))
                for col in updated_cols:
                    col.set("w", per_col_w)

        tr_lst = list(tbl.xpath("./a:tr"))
        current_rows = len(tr_lst)
        diff_rows = target_rows - current_rows

        if diff_rows > 0:
            for _ in range(diff_rows):
                new_tr = copy.deepcopy(list(tbl.xpath("./a:tr"))[-1])
                for tc in new_tr.xpath(".//a:tc"):
                    cls._clear_cell_text_xml(tc)
                tbl.append(new_tr)
        elif diff_rows < 0:
            for _ in range(abs(diff_rows)):
                tbl.remove(list(tbl.xpath("./a:tr"))[-1])

    @classmethod
    def _clear_cell_text_xml(cls, tc):
        for t in tc.iter():
            if t.tag.endswith("}t"):
                t.text = ""

    @classmethod
    def _set_cell_bg_color(cls, tc, hex_color: str):
        tcPr_list = tc.xpath("./a:tcPr")
        if tcPr_list:
            tcPr = tcPr_list[0]
        else:
            tcPr = OxmlElement("a:tcPr")
            tc.insert(0, tcPr)

        for fill in list(tcPr):
            if fill.tag.endswith("Fill"):
                tcPr.remove(fill)

        solidFill = OxmlElement("a:solidFill")
        srgbClr = OxmlElement("a:srgbClr")
        srgbClr.set("val", hex_color)
        solidFill.append(srgbClr)
        tcPr.append(solidFill)

    @classmethod
    def _set_cell_border(cls, tc, hex_color: str):
        tcPr_list = tc.xpath("./a:tcPr")
        if tcPr_list:
            tcPr = tcPr_list[0]
        else:
            tcPr = OxmlElement("a:tcPr")
            tc.insert(0, tcPr)

        for edge in ["lnB", "lnT", "lnL", "lnR"]:
            ln_list = tcPr.xpath(f"./a:{edge}")
            if ln_list:
                ln = ln_list[0]
            else:
                ln = OxmlElement(f"a:{edge}")
                ln.set("w", "12700")
                tcPr.append(ln)

            for fill in list(ln):
                if fill.tag.endswith("Fill"):
                    ln.remove(fill)

            solidFill = OxmlElement("a:solidFill")
            srgbClr = OxmlElement("a:srgbClr")
            srgbClr.set("val", hex_color)
            solidFill.append(srgbClr)
            ln.append(solidFill)

    @classmethod
    def _set_cell_text(
        cls,
        tc,
        text: str,
        text_color_hex: str,
        is_bold: bool = False,
        scale_factor: float = 1.0,
    ):
        txBody_list = tc.xpath("./a:txBody")
        if txBody_list:
            txBody = txBody_list[0]
        else:
            txBody = OxmlElement("a:txBody")
            bodyPr = OxmlElement("a:bodyPr")
            lstStyle = OxmlElement("a:lstStyle")
            txBody.append(bodyPr)
            txBody.append(lstStyle)
            tc.append(txBody)

        bodyPr_list = txBody.xpath("./a:bodyPr")
        if bodyPr_list:
            bodyPr = bodyPr_list[0]
            bodyPr.set("wrap", "square")
            pad_x = str(int(91440 * scale_factor))
            pad_y = str(int(68580 * scale_factor))
            bodyPr.set("lIns", pad_x)
            bodyPr.set("rIns", pad_x)
            bodyPr.set("tIns", pad_y)
            bodyPr.set("bIns", pad_y)

        for p in txBody.xpath("./a:p"):
            txBody.remove(p)

        if not text:
            txBody.append(OxmlElement("a:p"))
            return

        char_count = len(text)
        if char_count > 120:
            base_pt = 13.0
        elif char_count > 60:
            base_pt = 14.5
        else:
            base_pt = 16.0

        sz = str(int(round(base_pt * 100 * scale_factor)))

        lines = text.split("\n")
        for line in lines:
            line = line.strip()
            p = OxmlElement("a:p")
            pPr = OxmlElement("a:pPr")
            pPr.set("algn", "c" if len(line) <= 4 else "l")
            pPr.set("latinLnBrk", "0")

            r = OxmlElement("a:r")
            rPr = OxmlElement("a:rPr")
            rPr.set("sz", sz)
            if is_bold:
                rPr.set("b", "1")

            solidFill = OxmlElement("a:solidFill")
            srgbClr = OxmlElement("a:srgbClr")
            srgbClr.set("val", text_color_hex)
            solidFill.append(srgbClr)
            rPr.append(solidFill)

            t = OxmlElement("a:t")
            t.text = line if line else " "

            r.append(rPr)
            r.append(t)
            p.append(pPr)
            p.append(r)
            txBody.append(p)

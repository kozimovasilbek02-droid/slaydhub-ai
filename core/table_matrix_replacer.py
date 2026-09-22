# -*- coding: utf-8 -*-
"""
core/table_matrix_replacer.py
Dynamic Table & Matrix Replacer for SlaydHub AI.
- Synchronizes Table dimensions (rows/cols) with incoming data array.
- Preserves internal cell layout while dynamically adjusting font size to prevent overflow.
- Harmonizes 'Zebra Striping' and header background colors with Penpot themes.
"""

import copy
from typing import Dict, Any, List

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
        theme_key: str = "tech_minimal"
    ) -> bool:
        """
        Adapts a pptx Table geometry to the data's rows/cols count, and sets the cell values
        along with Penpot theme colors.
        Returns True if successful, False otherwise.
        """
        table = table_shape.table

        # Calculate target dimensions
        target_cols = len(headers) if headers else 0
        if rows:
            target_cols = max(target_cols, max(len(r) for r in rows))
        target_rows = len(rows) + (1 if headers else 0)

        if target_cols == 0 or target_rows == 0:
            return False

        # 1. Synchronize Dimensions (XML-level Deep Copy/Prune)
        try:
            cls._sync_table_dimensions(table, target_rows, target_cols)
        except Exception as e:
            # If complex XML manipulation fails, return False gracefully
            return False

        # Refetch the updated XML table structure
        tbl = table._tbl
        tr_lst = list(tbl.xpath('./a:tr'))

        # 2. Extract Penpot Colors for Harmonization
        theme = PenpotThemeManager.get_theme(theme_key)
        bg_slide = theme.get("bg_slide", "#FFFFFF").replace("#", "")
        primary = theme.get("primary", "#2563EB").replace("#", "")
        card_bg = theme.get("card_bg", "#FFFFFF").replace("#", "")
        card_border = theme.get("card_border", "#E2E8F0").replace("#", "")

        # Zebra alternating backgrounds: card_bg vs very light variant
        zebra_1 = card_bg
        zebra_2 = bg_slide if bg_slide.lower() != card_bg.lower() else "F8FAFC" # fallback light slate

        title_hex = theme.get("text_title", "#0F172A").replace("#", "")
        body_hex = theme.get("text_body", "#334155").replace("#", "")

        # 3. Populate Cells
        for r_idx, tr in enumerate(tr_lst):
            is_header = (headers and r_idx == 0)

            # Stylistic Decisions
            if is_header:
                bg_color = primary
                text_color_hex = "FFFFFF"  # White text on primary header
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

            tc_lst = list(tr.xpath('./a:tc'))
            for c_idx, tc in enumerate(tc_lst):
                val = str(row_data[c_idx]) if c_idx < len(row_data) else ""
                val = MicroTypography.sanitize_typography(val)

                # Apply Harmonization
                cls._set_cell_bg_color(tc, bg_color)
                cls._set_cell_border(tc, card_border)

                # Insert Text
                cls._set_cell_text(tc, val, text_color_hex, is_bold=is_bold)

        return True

    @classmethod
    def _sync_table_dimensions(cls, table, target_rows: int, target_cols: int):
        tbl = table._tbl

        # 1. Sync Columns via <a:tblGrid>
        grid = tbl.tblGrid
        gridCol_lst = list(grid.xpath('./a:gridCol'))
        current_cols = len(gridCol_lst)
        diff_cols = target_cols - current_cols

        total_grid_w = sum(int(c.get('w', 914400)) for c in gridCol_lst) if gridCol_lst else 9144000

        if diff_cols > 0:
            for _ in range(diff_cols):
                new_col = copy.deepcopy(gridCol_lst[-1])
                grid.append(new_col)
                for tr in tbl.xpath('./a:tr'):
                    tc_lst = list(tr.xpath('./a:tc'))
                    new_tc = copy.deepcopy(tc_lst[-1])
                    cls._clear_cell_text_xml(new_tc)
                    tr.append(new_tc)
        elif diff_cols < 0:
            for _ in range(abs(diff_cols)):
                grid.remove(list(grid.xpath('./a:gridCol'))[-1])
                for tr in tbl.xpath('./a:tr'):
                    tr.remove(list(tr.xpath('./a:tc'))[-1])

        # Redistribute column widths proportionally to preserve total table geometry
        updated_cols = list(grid.xpath('./a:gridCol'))
        if updated_cols:
            per_col_w = str(int(total_grid_w // len(updated_cols)))
            for col in updated_cols:
                col.set('w', per_col_w)

        # 2. Sync Rows via <a:tr>
        tr_lst = list(tbl.xpath('./a:tr'))
        current_rows = len(tr_lst)
        diff_rows = target_rows - current_rows

        if diff_rows > 0:
            for _ in range(diff_rows):
                new_tr = copy.deepcopy(list(tbl.xpath('./a:tr'))[-1])
                for tc in new_tr.xpath('.//a:tc'):
                    cls._clear_cell_text_xml(tc)
                tbl.append(new_tr)
        elif diff_rows < 0:
            for _ in range(abs(diff_rows)):
                tbl.remove(list(tbl.xpath('./a:tr'))[-1])

    @classmethod
    def _clear_cell_text_xml(cls, tc):
        """Empties all text node elements <a:t> inside a table cell."""
        for t in tc.iter():
            if t.tag.endswith('}t'):
                t.text = ""

    @classmethod
    def _set_cell_bg_color(cls, tc, hex_color: str):
        """Writes solidFill to the cell properties."""
        tcPr_list = tc.xpath('./a:tcPr')
        if tcPr_list:
            tcPr = tcPr_list[0]
        else:
            tcPr = OxmlElement('a:tcPr')
            tc.insert(0, tcPr)

        for fill in list(tcPr):
            if fill.tag.endswith('Fill'):
                tcPr.remove(fill)

        solidFill = OxmlElement('a:solidFill')
        srgbClr = OxmlElement('a:srgbClr')
        srgbClr.set('val', hex_color)
        solidFill.append(srgbClr)
        tcPr.append(solidFill)

    @classmethod
    def _set_cell_border(cls, tc, hex_color: str):
        """Sets a clean bottom border for cell separation."""
        tcPr_list = tc.xpath('./a:tcPr')
        if tcPr_list:
            tcPr = tcPr_list[0]
        else:
            tcPr = OxmlElement('a:tcPr')
            tc.insert(0, tcPr)

        for edge in ['lnB']: # Bottom border
            ln_list = tcPr.xpath(f'./a:{edge}')
            if ln_list:
                ln = ln_list[0]
            else:
                ln = OxmlElement(f'a:{edge}')
                ln.set('w', '12700') # 1 pt approx
                tcPr.append(ln)

            for fill in list(ln):
                if fill.tag.endswith('Fill'):
                    ln.remove(fill)

            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            srgbClr.set('val', hex_color)
            solidFill.append(srgbClr)
            ln.append(solidFill)

    @classmethod
    def _set_cell_text(cls, tc, text: str, text_color_hex: str, is_bold: bool = False):
        """Reconstructs the cell text body completely to prevent legacy formatting interference."""
        txBody_list = tc.xpath('./a:txBody')
        if txBody_list:
            txBody = txBody_list[0]
        else:
            txBody = OxmlElement('a:txBody')
            bodyPr = OxmlElement('a:bodyPr')
            lstStyle = OxmlElement('a:lstStyle')
            txBody.append(bodyPr)
            txBody.append(lstStyle)
            tc.append(txBody)

        # Ensure words wrap inside narrow cells
        bodyPr_list = txBody.xpath('./a:bodyPr')
        if bodyPr_list:
            bodyPr = bodyPr_list[0]
            bodyPr.set('wrap', 'square')
            # Increase margins slightly for clarity inside table grids
            bodyPr.set('lIns', '91440') # 0.1 inch
            bodyPr.set('rIns', '91440')

        # Clean existing paragraphs
        for p in txBody.xpath('./a:p'):
            txBody.remove(p)

        if not text:
            # Need at least one empty paragraph to prevent PPTX repair mode
            txBody.append(OxmlElement('a:p'))
            return

        # Dynamic shrinking based on text volume
        char_count = len(text)
        if char_count > 120:
            sz = '1200' # 12 pt
        elif char_count > 60:
            sz = '1350' # 13.5 pt
        else:
            sz = '1500' # 15 pt

        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            p = OxmlElement('a:p')

            # Paragraph formatting - no spacing after for tables, keeps it tight
            pPr = OxmlElement('a:pPr')
            pPr.set('algn', 'l') # align left

            r = OxmlElement('a:r')
            rPr = OxmlElement('a:rPr')
            rPr.set('sz', sz)
            if is_bold:
                rPr.set('b', '1')

            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            srgbClr.set('val', text_color_hex)
            solidFill.append(srgbClr)
            rPr.append(solidFill)

            t = OxmlElement('a:t')
            t.text = line if line else " "

            r.append(rPr)
            r.append(t)
            p.append(pPr)
            p.append(r)
            txBody.append(p)

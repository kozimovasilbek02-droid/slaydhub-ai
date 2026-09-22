# -*- coding: utf-8 -*-
"""
core/chart_harmonizer.py
Native PowerPoint Chart & Embedded Data Replacer for SlaydHub AI.

Features:
1. Detects native PowerPoint chart shapes (MSO_SHAPE_TYPE.CHART).
2. Updates underlying Excel workbook data via CategoryChartData.
3. Dynamically reconciles category/series count mismatches.
4. Harmonizes series and slice colors to PenpotThemeManager palettes.
"""

import logging
from typing import Dict, Any, List, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor

from core.penpot_themes import PenpotThemeManager

logger = logging.getLogger("ChartHarmonizer")


class ChartHarmonizer:
    """
    Intelligent PowerPoint Chart Data and Palette Harmonizer.
    """

    @classmethod
    def hex_to_rgb(cls, hex_str: str) -> RGBColor:
        clean = hex_str.replace("#", "").strip()
        if len(clean) == 3:
            clean = "".join([c * 2 for c in clean])
        r = int(clean[0:2], 16)
        g = int(clean[2:4], 16)
        b = int(clean[4:6], 16)
        return RGBColor(r, g, b)

    @classmethod
    def apply_chart_data(
        cls,
        chart_shape,
        chart_data_dict: Dict[str, Any],
        theme_key: str = "tech_minimal"
    ) -> bool:
        """
        Replaces chart categories and series values, and recolors series to match Penpot theme.
        Returns True if successful, False otherwise.
        """
        if not hasattr(chart_shape, "has_chart") or not chart_shape.has_chart:
            return False

        chart = chart_shape.chart

        categories = chart_data_dict.get("categories", [])
        raw_series = chart_data_dict.get("series", [])

        if not categories or not raw_series:
            logger.warning("ChartHarmonizer: Incomplete chart data provided (empty categories or series).")
            return False

        # 1. Sanitize series: Ensure all series values match category length
        num_cats = len(categories)
        sanitized_series = []
        for s in raw_series:
            s_name = str(s.get("name", "Ko'rsatkich"))
            s_vals = list(s.get("values", []))
            # Pad or truncate
            if len(s_vals) < num_cats:
                s_vals.extend([0.0] * (num_cats - len(s_vals)))
            elif len(s_vals) > num_cats:
                s_vals = s_vals[:num_cats]
            # Convert to float/int
            cleaned_vals = []
            for v in s_vals:
                try:
                    cleaned_vals.append(float(v))
                except Exception:
                    cleaned_vals.append(0.0)
            sanitized_series.append({"name": s_name, "values": cleaned_vals})

        # 2. Build CategoryChartData and replace workbook data
        try:
            cdata = CategoryChartData()
            cdata.categories = [str(c) for c in categories]
            for s in sanitized_series:
                cdata.add_series(s["name"], tuple(s["values"]))

            chart.replace_data(cdata)
            logger.info(f"ChartHarmonizer: Successfully replaced chart data ({num_cats} categories, {len(sanitized_series)} series).")
        except Exception as e:
            logger.error(f"ChartHarmonizer: Failed to replace chart data: {e}")
            return False

        # 3. Apply Penpot Color Harmonization
        try:
            theme = PenpotThemeManager.get_theme(theme_key)
            palette = [
                cls.hex_to_rgb(theme.get("primary", "#2563EB")),
                cls.hex_to_rgb(theme.get("secondary", "#7C3AED")),
                cls.hex_to_rgb(theme.get("accent", "#06B6D4")),
                cls.hex_to_rgb("#10B981"), # Emerald
                cls.hex_to_rgb("#F59E0B"), # Amber
            ]

            # Color each series
            for s_idx, series in enumerate(chart.series):
                color = palette[s_idx % len(palette)]
                try:
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = color
                except Exception as e_fill:
                    # Pie/donut charts: color individual points instead
                    try:
                        for p_idx, pt in enumerate(series.points):
                            pt_color = palette[p_idx % len(palette)]
                            pt.format.fill.solid()
                            pt.format.fill.fore_color.rgb = pt_color
                    except Exception:
                        pass
        except Exception as e_color:
            logger.debug(f"ChartHarmonizer: Color harmonization note: {e_color}")

        return True

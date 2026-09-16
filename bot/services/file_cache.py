# -*- coding: utf-8 -*-
"""
File Cache for Telegram Callback Data
-------------------------------------
Telegram callback_data 64 baytdan oshmasligi uchun
fayllarga qisqa ID berib keshlovchi xizmat.
"""

from pathlib import Path
from typing import Dict, Optional
import uuid

_CACHE: Dict[str, Path] = {}


def register_file(file_path: Path) -> str:
    """Faylni keshga qo'shib, 8 belgili qisqa ID qaytaradi."""
    short_id = uuid.uuid4().hex[:8]
    _CACHE[short_id] = Path(file_path)
    return short_id


def get_file(short_id_or_name: str, output_dir: Path) -> Optional[Path]:
    """Qisqa ID yoki to'liq fayl nomi bo'yicha faylni topadi."""
    if short_id_or_name in _CACHE and _CACHE[short_id_or_name].exists():
        return _CACHE[short_id_or_name]
    
    direct_p = output_dir / short_id_or_name
    if direct_p.exists():
        return direct_p
    
    # ID bo'yicha qidirish
    matches = list(output_dir.glob(f"*{short_id_or_name}*.pptx"))
    if matches:
        return matches[0]
        
    return None

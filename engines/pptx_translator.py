# -*- coding: utf-8 -*-
import re
from typing import List, Dict, Any, Tuple
from pptx import Presentation
from pptx.oxml import parse_xml
from core.gemini_client import call_gemini_cascade
from core.memory_guard import run_ram_heavy_sync

JUNK_WATERMARK_PATTERNS = [
    re.compile(r"presentationgo\.com", re.I),
    re.compile(r"slidesgo\.com", re.I),
    re.compile(r"allppt\.com", re.I),
    re.compile(r"slidehunter\.com", re.I),
    re.compile(r"powerpointify\.com", re.I),
    re.compile(r"canva\.com", re.I),
    re.compile(r"fonts\s*&\s*colors", re.I),
    re.compile(r"instructions\s+for\s+use", re.I),
]


def _set_latin_ln_brk(p_element) -> None:
    """OpenXML latinLnBrk='0' orqali o'zbekcha so'zlarning o'rtasidan bo'linib ketishini to'xtatadi."""
    try:
        pPr = p_element.find("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")
        if pPr is None:
            pPr = parse_xml(
                '<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" latinLnBrk="0"/>'
            )
            p_element.insert(0, pPr)
        else:
            pPr.set("latinLnBrk", "0")
    except Exception:
        pass


def _is_junk_text(text: str) -> bool:
    if not text:
        return False
    return any(p.search(text) for p in JUNK_WATERMARK_PATTERNS)


def extract_texts_from_shape(shape, extracted: List[Tuple[Any, str]], depth: int = 0) -> None:
    """10 qatlamgacha guruhlangan shakllar, jadvallar va bloklardan matnlarni rekursiv yig'adi."""
    if depth > 10:
        return

    if shape.has_subshapes if hasattr(shape, "has_subshapes") else (shape.shape_type == 6):
        for sub_shape in shape.shapes:
            extract_texts_from_shape(sub_shape, extracted, depth + 1)
        return

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    t = p.text.strip()
                    if t and not _is_junk_text(t):
                        extracted.append((p, t))
        return

    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            t = p.text.strip()
            if t and not _is_junk_text(t):
                extracted.append((p, t))


def _extract_unique_texts_sync(in_path: str) -> List[str]:
    """Sinxron OpenXML o'qish: RAM_COMPUTE_LOCK ichida chaqiriladi."""
    prs = Presentation(in_path)
    all_extracted: List[Tuple[Any, str]] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            extract_texts_from_shape(shape, all_extracted)
    seen = set()
    unique_texts = []
    for _, t in all_extracted:
        if t not in seen:
            seen.add(t)
            unique_texts.append(t)
    return unique_texts


def _apply_translations_sync(in_path: str, out_path: str, translated_map: Dict[str, str]) -> str:
    """Sinxron OpenXML yozish: RAM_COMPUTE_LOCK ichida chaqiriladi."""
    prs = Presentation(in_path)
    all_extracted: List[Tuple[Any, str]] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            extract_texts_from_shape(shape, all_extracted)

    for p_elem, orig_text in all_extracted:
        if orig_text in translated_map:
            new_text = translated_map[orig_text]
            if p_elem.runs:
                first_run = p_elem.runs[0]
                first_run.text = new_text
                for r in p_elem.runs[1:]:
                    r.text = ""
            else:
                p_elem.text = new_text
            _set_latin_ln_brk(p_elem._p)

    prs.save(out_path)
    return out_path


async def translate_pptx_deck(in_path: str, out_path: str) -> str:
    """
    Taqdimotni dizaynini 100% saqlagan holda o'zbek tiliga o'giruvchi chuqur OpenXML dvigateli.
    1) Matnni yig'ish -> RAM_COMPUTE_LOCK
    2) Gemini API tarjima -> NETWORK_IO_SEMAPHORE (RAM qulfini band qilmaydi!)
    3) Tarjimani yozish -> RAM_COMPUTE_LOCK
    """
    texts_to_translate = await run_ram_heavy_sync(_extract_unique_texts_sync, in_path)

    if not texts_to_translate:
        return await run_ram_heavy_sync(_apply_translations_sync, in_path, out_path, {})

    batch_size = 25
    translated_map: Dict[str, str] = {}

    for i in range(0, len(texts_to_translate), batch_size):
        chunk = texts_to_translate[i : i + batch_size]
        prompt = (
            "Vazifa: Quyidagi taqdimot matnlari ro'yxatini ma'nosi va ilmiy kontekstini aniq "
            "saqlagan holda zamonaviy o'zbek adabiy tiliga (lotin yozuvida) tarjima qilib ber.\n"
            "Chiqish formati: Faqat JSON massiv (massiv elementlari soni aynan kiritilgan ro'yxat bilan teng bo'lsin!).\n\n"
            f"Matnlar: {chunk}"
        )
        schema = {"type": "ARRAY", "items": {"type": "STRING"}}
        res, _ = await call_gemini_cascade(prompt, schema=schema, max_output_tokens=4096)
        if res and isinstance(res, list) and len(res) == len(chunk):
            for orig, trans in zip(chunk, res):
                translated_map[orig] = str(trans)
        else:
            for orig in chunk:
                translated_map[orig] = orig

    return await run_ram_heavy_sync(_apply_translations_sync, in_path, out_path, translated_map)

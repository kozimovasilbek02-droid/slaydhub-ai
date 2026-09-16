# -*- coding: utf-8 -*-
import os
import sys
import uuid
import shutil
import glob
import re
import time
from typing import Dict, List, Any, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import smart_slide_engine
import slide_manager

app = FastAPI(
    title="SlideAI Studio - Precision Vector Slide Recreator & Translator",
    description="Professional AI PowerPoint Studio: OCR Vector Reconstruction & Translation",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

TEMP_STORAGE = os.path.abspath(os.path.join(parent_dir, "temp_sessions"))
os.makedirs(TEMP_STORAGE, exist_ok=True)
OUTPUT_BASE = os.path.join(parent_dir, "output", "presentations")
os.makedirs(OUTPUT_BASE, exist_ok=True)
UPLOAD_BASE = os.path.join(parent_dir, "output", "uploads")
os.makedirs(UPLOAD_BASE, exist_ok=True)

# ─────────────────────────────────────────────────────────
# 0. SYSTEM HEALTH & DIRECT API
# ─────────────────────────────────────────────────────────
@app.api_route("/", methods=["GET", "HEAD"])
@app.api_route("/health", methods=["GET", "HEAD"])
@app.api_route("/api/v1/health", methods=["GET", "HEAD"])
async def health_check():
    return {
        "status": "ok",
        "service": "SlideTranslate AI / SlideAI Studio",
        "version": "2.5.0",
        "uptime": "24/7 active"
    }

@app.post("/api/v1/translate")
async def translate_v1_direct(
    file: UploadFile = File(...),
    target_lang: str = Form("uz-Latn"),
    api_key: Optional[str] = Form(None),
    domain: str = Form("general"),
    auto_fit: bool = Form(True)
):
    if not file.filename or not file.filename.lower().endswith((".pptx", ".potx")):
        raise HTTPException(status_code=400, detail="Faqat .pptx formatidagi fayllar qabul qilinadi.")

    target_script = "cyrillic" if target_lang.lower() in ["uz-cyrl", "cyrillic", "kirill"] else "latin"
    session_id = str(uuid.uuid4())
    in_path = os.path.join(TEMP_STORAGE, f"{session_id}_{file.filename}")
    with open(in_path, "wb") as f:
        f.write(await file.read())

    # Extract presentation data
    extracted = PPTXProcessor.extract_presentation_data(in_path)
    all_items = [it for s in extracted.get("slides", []) for it in s.get("items", [])]

    translator = GeminiTranslator(api_key=api_key)
    stats: Dict[str, Any] = {}
    if all_items:
        translated_results = translator.translate_items_batch(
            items=all_items,
            target_script=target_script,
            domain=domain,
            stats=stats
        )
        trans_map = {r["id"]: r["translated_text"] for r in translated_results}
    else:
        trans_map = {}

    out_name = f"translated_{file.filename}"
    out_path = os.path.join(OUTPUT_BASE, out_name)
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=in_path,
        translations_map=trans_map,
        output_pptx_path=out_path,
        auto_fit=auto_fit,
        target_script=target_script
    )
    return FileResponse(out_path, filename=out_name)

# ─────────────────────────────────────────────────────────
# 1. FOLDER DISCOVERY & MANAGEMENT
# ─────────────────────────────────────────────────────────
@app.get("/api/folders")
async def list_folders(search: Optional[str] = None):
    """Returns all available presentation folders (597+ folders + uploaded folders)."""
    all_folders = slide_manager.get_cached_folders()
    result = []
    
    # Also check user uploads
    if os.path.exists(UPLOAD_BASE):
        for up_dir in os.listdir(UPLOAD_BASE):
            full_up = os.path.join(UPLOAD_BASE, up_dir)
            if os.path.isdir(full_up):
                all_folders[f"📤 {up_dir}"] = full_up
                
    q = search.lower().strip() if search else ""
    for label, fpath in all_folders.items():
        fname = os.path.basename(os.path.normpath(fpath))
        if q and (q not in label.lower() and q not in fname.lower()):
            continue
            
        images = smart_slide_engine.get_images_in_folder(fpath)
        pptx_path = slide_manager.get_folder_pptx_path(fpath)
        out_folder = os.path.join(OUTPUT_BASE, fname)
        previews = glob.glob(os.path.join(out_folder, "Generated_Slide_*.png"))
        
        result.append({
            "name": fname,
            "label": label,
            "path": fpath,
            "slide_count": len(images),
            "is_ready": pptx_path is not None and os.path.exists(pptx_path),
            "has_previews": len(previews) > 0,
            "pptx_path": pptx_path
        })
        
    return {"folders": result, "total": len(result)}

@app.get("/api/folder/{folder_name}/details")
async def get_folder_details(folder_name: str):
    """Returns slides, previews, and comparisons for a folder."""
    all_folders = slide_manager.get_cached_folders()
    target_path = None
    
    # Check uploads first
    up_cand = os.path.join(UPLOAD_BASE, folder_name)
    if os.path.exists(up_cand):
        target_path = up_cand
    else:
        for _, fpath in all_folders.items():
            if os.path.basename(os.path.normpath(fpath)) == folder_name:
                target_path = fpath
                break
                
    if not target_path or not os.path.exists(target_path):
        for s_dir in slide_manager.SOURCE_DIRS:
            candidate = os.path.join(s_dir, folder_name)
            if os.path.exists(candidate):
                target_path = candidate
                break
                
    if not target_path:
        raise HTTPException(status_code=404, detail=f"Papka topilmadi: {folder_name}")
        
    images = smart_slide_engine.get_images_in_folder(target_path)
    out_folder = os.path.join(OUTPUT_BASE, folder_name)
    pptx_path = slide_manager.get_folder_pptx_path(target_path)
    
    previews = sorted(glob.glob(os.path.join(out_folder, "Generated_Slide_*.png")), key=smart_slide_engine.natural_sort_key)
    comparisons = sorted(glob.glob(os.path.join(out_folder, "comparisons", "Comparison_Slide_*.png")), key=smart_slide_engine.natural_sort_key)
    
    slides_info = []
    for i in range(len(images)):
        orig_img = images[i]
        gen_img = previews[i] if i < len(previews) else None
        comp_img = comparisons[i] if i < len(comparisons) else None
        
        slides_info.append({
            "index": i + 1,
            "original_filename": os.path.basename(orig_img),
            "original_url": f"/api/image?path={orig_img}",
            "generated_url": f"/api/image?path={gen_img}" if gen_img else None,
            "comparison_url": f"/api/image?path={comp_img}" if comp_img else None,
        })
        
    return {
        "folder_name": folder_name,
        "folder_path": target_path,
        "total_slides": len(images),
        "pptx_ready": pptx_path is not None and os.path.exists(pptx_path),
        "pptx_download_url": f"/api/download_pptx/{folder_name}" if pptx_path else None,
        "slides": slides_info
    }

# ─────────────────────────────────────────────────────────
# 2. DIRECT DRAG & DROP UPLOAD (SINGLE IMAGE OR MULTIPLE)
# ─────────────────────────────────────────────────────────
@app.post("/api/upload_drop")
async def upload_drag_drop(
    files: List[UploadFile] = File(...),
    custom_name: Optional[str] = Form(None)
):
    """
    Accepts one or more dragged-and-dropped images (or folder contents).
    Automatically saves, runs OCR + Vector recreation, and returns full deck info.
    """
    if not files:
        raise HTTPException(status_code=400, detail="Fayl tanlanmadi!")
        
    # Generate folder name
    if custom_name and custom_name.strip():
        safe_name = slide_manager.sanitize_name(custom_name.strip())
    else:
        first_fn = os.path.splitext(files[0].filename)[0]
        safe_name = f"Drop_{slide_manager.sanitize_name(first_fn)}_{int(time.time())}"
        
    target_upload_dir = os.path.join(UPLOAD_BASE, safe_name)
    os.makedirs(target_upload_dir, exist_ok=True)
    
    saved_images = []
    for i, file in enumerate(files):
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in [".png", ".jpg", ".jpeg", ".webp"]:
            continue
            
        dest_filename = f"Slide_{i+1:02d}{ext}"
        dest_path = os.path.join(target_upload_dir, dest_filename)
        with open(dest_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        saved_images.append(dest_path)
        
    if not saved_images:
        raise HTTPException(status_code=400, detail="Tashlangan fayllar orasida rasm topilmadi (.png, .jpg, .jpeg, .webp)")
        
    # Immediately run Smart Vector Recreation
    out_folder = os.path.join(OUTPUT_BASE, safe_name)
    os.makedirs(out_folder, exist_ok=True)
    pptx_path = os.path.join(out_folder, "presentation_95_precision.pptx")
    
    success, msg = smart_slide_engine.build_multi_slide_presentation(target_upload_dir, pptx_path)
    if success:
        # Export previews and comparisons
        smart_slide_engine.export_deck_previews(pptx_path, out_folder)
        comp_dir = os.path.join(out_folder, "comparisons")
        smart_slide_engine.generate_deck_comparisons(target_upload_dir, out_folder, comp_dir)
        
    return {
        "status": "ok",
        "folder_name": safe_name,
        "total_slides": len(saved_images),
        "message": msg
    }

# ─────────────────────────────────────────────────────────
# 3. IMAGE SERVING
# ─────────────────────────────────────────────────────────
@app.get("/api/image")
async def serve_image(path: str):
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Rasm topilmadi")
    return FileResponse(path)

# ─────────────────────────────────────────────────────────
# 4. AI VECTOR GENERATION & RECREATION
# ─────────────────────────────────────────────────────────
class GenerateRequest(BaseModel):
    folder_name: str

@app.post("/api/generate")
async def generate_vector_presentation(req: GenerateRequest):
    all_folders = slide_manager.get_cached_folders()
    target_path = None
    
    up_cand = os.path.join(UPLOAD_BASE, req.folder_name)
    if os.path.exists(up_cand):
        target_path = up_cand
    else:
        for _, fpath in all_folders.items():
            if os.path.basename(os.path.normpath(fpath)) == req.folder_name:
                target_path = fpath
                break
                
    if not target_path:
        for s_dir in slide_manager.SOURCE_DIRS:
            cand = os.path.join(s_dir, req.folder_name)
            if os.path.exists(cand):
                target_path = cand
                break
                
    if not target_path:
        raise HTTPException(status_code=404, detail=f"Papka topilmadi: {req.folder_name}")
        
    success, msg = slide_manager.convert_presentation(target_path)
    if not success:
        raise HTTPException(status_code=500, detail=msg)
        
    return {"status": "ok", "message": msg, "folder_name": req.folder_name}

# ─────────────────────────────────────────────────────────
# 5. SINGLE SLIDE AI INSPECTOR
# ─────────────────────────────────────────────────────────
@app.get("/api/inspect_slide")
async def inspect_slide(image_path: str):
    if not os.path.exists(image_path):
        raise HTTPException(status_code=404, detail="Rasm topilmadi")
        
    img_pil = Image.open(image_path).convert("RGB")
    img_cv = cv2.imread(image_path)
    if img_cv is None:
        img_cv = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
        
    bg_color = smart_slide_engine.get_background_color(img_pil)
    cards = smart_slide_engine.detect_individual_cards(img_cv, bg_color)
    ocr_items = smart_slide_engine.extract_ocr_hierarchy(img_pil, bg_color, cards)
    
    cards_json = []
    for c in cards:
        cards_json.append({
            "box": c["box"],
            "color_rgb": c["color"],
            "color_hex": f"#{c['color'][0]:02x}{c['color'][1]:02x}{c['color'][2]:02x}",
            "is_dark": c["is_dark"],
            "area": c["area"]
        })
        
    items_json = []
    for item in ocr_items:
        items_json.append({
            "text": item["text"],
            "box": item["box"],
            "color_rgb": item["color"],
            "color_hex": f"#{item['color'][0]:02x}{item['color'][1]:02x}{item['color'][2]:02x}",
            "font_pt": item["font_pt"],
            "is_title": item["is_title"],
            "inside_card": item["inside_card"] is not None
        })
        
    return {
        "bg_color_rgb": bg_color,
        "bg_color_hex": f"#{bg_color[0]:02x}{bg_color[1]:02x}{bg_color[2]:02x}",
        "cards": cards_json,
        "text_items": items_json,
        "total_words": sum(len(x["text"].split()) for x in ocr_items)
    }

# ─────────────────────────────────────────────────────────
# 6. PPTX DOWNLOAD & LAUNCH
# ─────────────────────────────────────────────────────────
@app.get("/api/download_pptx/{folder_name}")
async def download_pptx(folder_name: str):
    out_folder = os.path.join(OUTPUT_BASE, folder_name)
    pptx_path = os.path.join(out_folder, "presentation_95_precision.pptx")
    if not os.path.exists(pptx_path):
        raise HTTPException(status_code=404, detail="PPTX fayli hali yaratilmagan!")
    return FileResponse(pptx_path, filename=f"{folder_name}_Vector_Presentation.pptx")

@app.post("/api/open_in_powerpoint")
async def open_in_powerpoint(req: GenerateRequest):
    out_folder = os.path.join(OUTPUT_BASE, req.folder_name)
    pptx_path = os.path.join(out_folder, "presentation_95_precision.pptx")
    if not os.path.exists(pptx_path):
        raise HTTPException(status_code=404, detail="PPTX fayli topilmadi!")
    os.startfile(pptx_path)
    return {"status": "ok", "message": "PowerPoint dasturida ochildi"}

# ─────────────────────────────────────────────────────────
# 7. AI PRESENTATION GENERATION API (GAMMA & VECTOR)
# ─────────────────────────────────────────────────────────
from backend.core.ai_presentation_generator import AIPresentationGenerator
from backend.core.deck_builder import DeckBuilder, PALETTES
from backend.core.gamma_generator import GammaGenerator

ai_gen_service = AIPresentationGenerator()

class PresentationGenerateRequest(BaseModel):
    topic: str
    slide_count: int = 10
    style: str = "vector"  # "gamma", "vector"
    theme: str = "uzbek_blue"  # "uzbek_blue", "emerald_teal", "dark_slate", "modern_purple", "crimson_ruby"
    provider: str = "auto"

@app.get("/api/generate/styles")
async def get_generation_styles():
    return {
        "styles": [
            {"id": "gamma", "name": "Gamma-Style (Ultra-HD)", "desc": "Glassmorphism, gradients and high-res cards"},
            {"id": "vector", "name": "Native Vector (Editable)", "desc": "100% editable PowerPoint vector shapes and layouts"}
        ],
        "themes": list(PALETTES.keys())
    }

@app.post("/api/generate/presentation")
async def generate_ai_presentation(req: PresentationGenerateRequest):
    if not req.topic or len(req.topic.strip()) < 3:
        raise HTTPException(status_code=400, detail="Mavzu kamida 3 ta belgidan iborat bo'lishi kerak!")

    spec = await ai_gen_service.generate_deck_spec(
        topic=req.topic,
        slide_count=req.slide_count,
        theme=req.theme,
        provider=req.provider
    )

    file_id = str(uuid.uuid4())[:8]
    clean_topic = "".join(c for c in req.topic if c.isalnum() or c in (" ", "_", "-")).strip()[:35]
    out_name = f"{clean_topic}_{req.style}_{file_id}.pptx"
    out_path = os.path.join(OUTPUT_BASE, "_generated", out_name)

    if req.style == "gamma":
        final_path = await GammaGenerator.build(spec, out_path)
    else:
        final_path = DeckBuilder.build(spec, out_path)

    return {
        "status": "success",
        "topic": req.topic,
        "style": req.style,
        "theme": req.theme,
        "slide_count": len(spec.get("slides", [])),
        "filename": out_name,
        "download_url": f"/api/download_generated/{out_name}"
    }

@app.get("/api/download_generated/{filename}")
async def download_generated_pptx(filename: str):
    f_path = os.path.join(OUTPUT_BASE, "_generated", filename)
    if not os.path.exists(f_path):
        raise HTTPException(status_code=404, detail="Fayl topilmadi!")
    return FileResponse(f_path, filename=filename)

# ─────────────────────────────────────────────────────────
# 7. SERVE COMPILED REACT FRONTEND
# ─────────────────────────────────────────────────────────
FRONTEND_DIST = os.path.join(parent_dir, "frontend", "dist")
if os.path.exists(FRONTEND_DIST):
    app.mount("/", StaticFiles(directory=FRONTEND_DIST, html=True), name="frontend")

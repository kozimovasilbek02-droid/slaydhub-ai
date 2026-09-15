import os
import sys
import glob
import re
import subprocess
import shutil

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import smart_slide_engine

output_dir = os.path.join(current_dir, "output", "presentations")
os.makedirs(output_dir, exist_ok=True)

SOURCE_DIRS = [
    r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar",
    r"C:\Users\user\Desktop\5M_Of_Advertising",
    r"G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar",
    r"C:\Users\user\Desktop"
]

_cached_folders = None

def get_available_folders():
    return get_cached_folders()

def get_cached_folders():
    global _cached_folders
    if _cached_folders is not None:
        return _cached_folders
        
    folders = {}
    
    # 1. Desktop 5M
    p_5m = r"C:\Users\user\Desktop\5M_Of_Advertising"
    if os.path.exists(p_5m):
        folders["🎯 5M Of Advertising (Desktop)"] = p_5m
        
    # 2. CrossDevice / Phone / Google Drive Presentations (597+ folders)
    for s_dir in SOURCE_DIRS:
        if os.path.exists(s_dir) and os.path.isdir(s_dir):
            try:
                subdirs = sorted([d for d in os.listdir(s_dir) if os.path.isdir(os.path.join(s_dir, d))])
                for sub in subdirs:
                    full_p = os.path.join(s_dir, sub)
                    imgs = smart_slide_engine.get_images_in_folder(full_p)
                    if imgs:
                        clean_title = sub.replace('_', ' ').title()
                        folders[f"📁 {clean_title} ({len(imgs)} slayd)"] = full_p
            except Exception:
                pass
                
    _cached_folders = folders
    return _cached_folders

def search_drive_folders(query, limit=50):
    if not query:
        return []
    q_lower = query.lower()
    results = []
    
    for s_dir in SOURCE_DIRS:
        if os.path.exists(s_dir) and os.path.isdir(s_dir):
            try:
                for d in os.listdir(s_dir):
                    if q_lower in d.lower() and os.path.isdir(os.path.join(s_dir, d)):
                        results.append(os.path.join(s_dir, d))
                        if len(results) >= limit:
                            return results
            except Exception:
                pass
    return results

def natural_sort_key(path):
    m = re.search(r'(?:Slide_|_s|slide_|page_)(\d+)', os.path.basename(path), re.IGNORECASE)
    return int(m.group(1)) if m else 999

def get_folder_slides(folder_path):
    if not folder_path or not os.path.exists(folder_path):
        return []
    return smart_slide_engine.get_images_in_folder(folder_path)

def get_folder_pptx_path(folder_path):
    if not folder_path:
        return None
        
    folder_name = os.path.basename(os.path.normpath(folder_path))
    
    # Check in output/presentations/{folder_name}/
    out_folder = os.path.join(output_dir, folder_name)
    p_out = os.path.join(out_folder, "presentation_95_precision.pptx")
    if os.path.exists(p_out):
        return p_out
        
    # Check inside folder itself
    p1 = os.path.join(folder_path, "presentation_95_precision.pptx")
    if os.path.exists(p1):
        return p1
        
    p2 = os.path.join(folder_path, "presentation.pptx")
    if os.path.exists(p2):
        return p2
        
    all_pptx = glob.glob(os.path.join(folder_path, "*.pptx"))
    if all_pptx:
        return all_pptx[0]
        
    return None

def sanitize_name(name):
    return re.sub(r'[^a-zA-Z0-9_-]', '_', name)

def export_pptx_previews(pptx_path, folder_name):
    if not pptx_path or not os.path.exists(pptx_path):
        return []
        
    safe_name = sanitize_name(folder_name)
    out_folder = os.path.join(output_dir, safe_name)
    
    # Check existing
    existing = sorted(glob.glob(os.path.join(out_folder, "Generated_Slide_*.png")), key=natural_sort_key)
    if existing:
        return existing
        
    return smart_slide_engine.export_deck_previews(pptx_path, out_folder)

def convert_presentation(folder_path):
    """
    Core Converter: Converts any folder of images into 100% editable vector PPTX
    using RapidOCR, OpenCV card segmentation, timeline nodes, and auto contrast.
    """
    if not folder_path or not os.path.exists(folder_path):
        return False, "Tanlangan papka topilmadi."
        
    folder_name = os.path.basename(os.path.normpath(folder_path))
    out_folder = os.path.join(output_dir, folder_name)
    os.makedirs(out_folder, exist_ok=True)
    
    pptx_path = os.path.join(out_folder, "presentation_95_precision.pptx")
    comp_dir = os.path.join(out_folder, "comparisons")
    
    print(f"[*] Smart vector recreation boshlandi: {folder_name}")
    success, msg = smart_slide_engine.build_multi_slide_presentation(folder_path, pptx_path)
    
    if not success:
        return False, msg
        
    # Export previews and comparisons
    previews = smart_slide_engine.export_deck_previews(pptx_path, out_folder)
    smart_slide_engine.generate_deck_comparisons(folder_path, out_folder, comp_dir)
    
    # Also copy PPTX to source folder if writable
    try:
        shutil.copy2(pptx_path, os.path.join(folder_path, "presentation_95_precision.pptx"))
    except Exception:
        pass
        
    return True, f"{msg} | Prevyular va taqqoslash kartalari tayyorlandi!"

"""
UNIFIED PRESENTATION ENGINE HUB
================================
This module integrates all 4 core PowerPoint engines:
1. PptxGenJS (Node.js engine)
2. Marp CLI (Markdown-to-PPTX engine)
3. Presenton (AI theme & template catalogue)
4. Python-pptx + OpenCV Smart Shapes Vector Engine
"""

import os
import subprocess
from core.smart_shapes import (
    draw_folded_tiered_cards,
    draw_interlocking_cards,
    add_header,
    hex_to_rgb
)
from core.templates import build_roadmap_timeline

WORKSPACE_ROOT = os.path.abspath(os.path.dirname(__file__))
TOOLS_DIR = os.path.join(WORKSPACE_ROOT, 'tools')
PRESENTON_DIR = os.path.join(TOOLS_DIR, 'presenton')
NODE_MODULES = os.path.join(WORKSPACE_ROOT, 'node_modules')

def verify_all_engines():
    status = {}
    
    # 1. PptxGenJS
    pptxgen_path = os.path.join(NODE_MODULES, 'pptxgenjs')
    status['PptxGenJS'] = {
        'installed': os.path.exists(pptxgen_path),
        'path': pptxgen_path,
        'description': 'JavaScript/TypeScript PowerPoint Vector Generator'
    }
    
    # 2. Marp CLI
    marp_path = os.path.join(NODE_MODULES, '@marp-team', 'marp-cli')
    status['Marp-CLI'] = {
        'installed': os.path.exists(marp_path),
        'path': marp_path,
        'description': 'Markdown to PPTX CLI Engine'
    }
    
    # 3. Presenton
    status['Presenton'] = {
        'installed': os.path.exists(PRESENTON_DIR),
        'path': PRESENTON_DIR,
        'description': 'Open-source AI Presentation Framework (Gamma/Decktopus analog)'
    }
    
    # 4. Smart Shapes Vector Engine
    smart_shapes_path = os.path.join(WORKSPACE_ROOT, 'core', 'smart_shapes.py')
    status['SmartShapes_OpenCV'] = {
        'installed': os.path.exists(smart_shapes_path),
        'path': smart_shapes_path,
        'description': 'Pixel-perfect Vector Geometry & Decompiler Engine'
    }
    
    return status

if __name__ == '__main__':
    print("==================================================")
    print("   UNIFIED PPTX AGENT ENGINE STATUS")
    print("==================================================")
    engines = verify_all_engines()
    for name, info in engines.items():
        state = "[OK] INSTALLED" if info['installed'] else "[MISSING]"
        print(f"{state} - {name}:")
        print(f"   Path: {info['path']}")
        print(f"   Role: {info['description']}\n")

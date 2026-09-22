import os
import sys
import traceback
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from backend.core.gemini_translator import GeminiTranslator
from backend.core.pptx_processor import PPTXProcessor
from scripts.fix_and_refine_uz_slides import refine_single_file

f = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ\Adabiyotshunoslik - Adabiyotda miforealizm yo'nalishi.pptx"
print(f"Testing on {f}...")

translator = GeminiTranslator()
try:
    res = refine_single_file(f, translator)
    print("Result:", res)
except Exception:
    traceback.print_exc()

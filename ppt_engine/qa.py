import os
import time
import pythoncom
import win32com.client
from typing import List

class VisionQA:
    """PowerPoint COM export and visual QA validation engine."""
    
    @staticmethod
    def kill_existing_powerpoint():
        """Terminates any hanging PowerPoint background instances."""
        try:
            os.system("powershell -Command \"Stop-Process -Name 'POWERPNT' -Force -ErrorAction SilentlyContinue\"")
            time.sleep(0.5)
        except Exception:
            pass

    @staticmethod
    def export_slides_to_png(pptx_path: str, output_dir: str = 'output/qa_previews', width: int = 1920, height: int = 1080) -> List[str]:
        """
        Exports all slides of a PPTX file to high-res 1920x1080 PNG images.
        Returns list of absolute PNG paths.
        """
        VisionQA.kill_existing_powerpoint()
        os.makedirs(output_dir, exist_ok=True)
        
        abs_pptx = os.path.abspath(pptx_path)
        if not os.path.exists(abs_pptx):
            raise FileNotFoundError(f"PPTX file not found: {abs_pptx}")
            
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch('PowerPoint.Application')
        
        exported_files = []
        try:
            pres = ppt.Presentations.Open(abs_pptx, WithWindow=False)
            base_name = os.path.splitext(os.path.basename(pptx_path))[0]
            
            for i in range(1, pres.Slides.Count + 1):
                out_png = os.path.abspath(os.path.join(output_dir, f"{base_name}_s{i}.png"))
                pres.Slides.Item(i).Export(out_png, 'PNG', width, height)
                exported_files.append(out_png)
                
            pres.Close()
        finally:
            ppt.Quit()
            pythoncom.CoUninitialize()
            
        return exported_files

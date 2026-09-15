import os
from ppt_engine.diff import VisualDiff

folder_path = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\15_Famous_Foods_Of_The_Usa'
scores = []

for i in range(1, 17):
    tmpl = os.path.join(folder_path, f'Slide_{i:02d}.png')
    gen = f'output/folder11_flawless_previews/presentation_95_precision_s{i}.png'
    if os.path.exists(tmpl) and os.path.exists(gen):
        ssim_val = VisualDiff.calculate_ssim(tmpl, gen)
        scores.append((i, ssim_val))
        print(f"Slide {i:02d}: SSIM = {ssim_val*100:.2f}%")

avg_ssim = sum(s[1] for s in scores) / len(scores) if scores else 0
print(f"\nAverage SSIM across 16 slides: {avg_ssim*100:.2f}%")

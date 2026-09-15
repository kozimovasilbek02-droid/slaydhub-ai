import os
from ppt_engine import VisualDiff

ref_img = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\100_Days_Action_Plan\Slide_02.png'
gen_img = r'C:\Users\user\Desktop\Antigravity\Power Point\output\action_plan_previews\action_plan_s2_perfect_wave_s1.png'
diff_out = r'C:\Users\user\Desktop\Antigravity\Power Point\output\action_plan_previews\diff_heatmap_s2.png'

ssim_score = VisualDiff.calculate_ssim(ref_img, gen_img)
print(f"Matematik SSIM O'xshashlik Bali: {ssim_score * 100:.2f}%")

VisualDiff.generate_diff_heatmap(ref_img, gen_img, diff_out)
print(f"Diff Heatmap saqlandi: {diff_out}")

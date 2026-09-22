import os
import shutil

src = os.path.abspath('output/action_plan_s2_flawless_chain.pptx')
dst = os.path.abspath(r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\100_Days_Action_Plan\presentation_95_precision.pptx')
shutil.copy(src, dst)
print('Master presentation for 100_Days_Action_Plan updated to:', dst)

"""
Find exact Drive paths safely
"""
import os, sys

for root_drive in ['G:', 'C:', 'D:']:
    if os.path.exists(root_drive):
        print(f"Drive {root_drive} exists")
        try:
            items = os.listdir(root_drive)
            for item in items:
                if 'диск' in item or 'Drive' in item or 'Slide' in item:
                    print(f"  Found matching folder: {item}")
                    sub_p = os.path.join(root_drive, item, "SlideEgg_Downloads", "Prezentatsiyalar")
                    if os.path.exists(sub_p):
                        print(f"  -> VALID PATH: {sub_p}")
        except Exception as e:
            print(f"  Error reading {root_drive}: {e}")

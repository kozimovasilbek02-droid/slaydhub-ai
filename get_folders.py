"""
Get exact folder names from Google Drive Prezentatsiyalar
"""
import os, json

drive_base = r"G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(drive_base) if os.path.isdir(os.path.join(drive_base, f))])

print(f"Total folders: {len(folders)}")
for i, f in enumerate(folders[:25], 1):
    print(f"[{i}] {f}")

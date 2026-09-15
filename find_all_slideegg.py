"""
Search for SlideEgg_Downloads across all drives
"""
import os, string

for letter in string.ascii_uppercase:
    drive = f"{letter}:\\"
    if os.path.exists(drive):
        print(f"Drive {drive} is mounted")
        try:
            for root, dirs, files in os.walk(drive):
                if "SlideEgg_Downloads" in dirs:
                    found = os.path.join(root, "SlideEgg_Downloads")
                    print(f"  FOUND SlideEgg_Downloads at: {found}")
                    break
                # Only check top 2 levels
                if root.count(os.sep) - drive.count(os.sep) >= 2:
                    dirs.clear()
        except Exception:
            pass

import string, os, glob

print("Checking all drive letters:")
for letter in string.ascii_uppercase:
    d = f"{letter}:\\"
    if os.path.exists(d):
        print(f"  Drive {d} exists")
        try:
            items = os.listdir(d)
            print(f"    Contents: {items[:8]}")
        except Exception as e:
            print(f"    Error: {e}")

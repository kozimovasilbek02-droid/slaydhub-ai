import os, glob, re

current_dir = os.path.dirname(os.path.abspath(__file__))
builder_files = glob.glob(os.path.join(current_dir, "build_*.py"))

print(f"Patching {len(builder_files)} builder scripts...")

for bpath in builder_files:
    try:
        with open(bpath, "r", encoding="utf-8") as f:
            code = f.read()
            
        # Ensure sys is imported
        if "import sys" not in code:
            code = "import sys\n" + code
            
        # Replace hardcoded save logic with sys.argv[1] priority
        pattern = r"prs\.save\([^\)]+\)"
        replacement = """
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
"""
        # Replace all prs.save calls
        new_code = re.sub(pattern, replacement.strip(), code)
        
        with open(bpath, "w", encoding="utf-8") as f:
            f.write(new_code)
        print(f"  [OK] Patched {os.path.basename(bpath)}")
    except Exception as e:
        print(f"  [ERROR] {os.path.basename(bpath)}: {e}")

print("All builder scripts successfully patched for dynamic output paths!")

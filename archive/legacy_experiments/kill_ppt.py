import subprocess
try:
    subprocess.run(["taskkill", "/f", "/im", "POWERPNT.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("PowerPoint terminated cleanly.")
except Exception as e:
    print(f"Error: {e}")

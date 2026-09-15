import sys
for pkg in ['fitz', 'pdf2image', 'pptx', 'PIL', 'win32com', 'comtypes']:
    try:
        __import__(pkg)
        print(f"Package {pkg}: AVAILABLE")
    except ImportError:
        print(f"Package {pkg}: NOT available")

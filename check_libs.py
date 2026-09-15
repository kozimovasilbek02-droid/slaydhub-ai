import sys

libs = ['rembg', 'onnxruntime', 'shapely', 'cairosvg', 'cv2', 'PIL', 'scipy', 'pptx', 'skia']
for lib in libs:
    try:
        __import__(lib)
        print(f"[INSTALLED] {lib}")
    except ImportError:
        print(f"[MISSING] {lib}")

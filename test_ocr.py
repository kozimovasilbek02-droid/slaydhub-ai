from rapidocr_onnxruntime import RapidOCR
from PIL import Image
import numpy as np

engine = RapidOCR()
img_path = r'C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\100_Day_New_Leadership_Position_Plan\Slide_01.png'
img = Image.open(img_path)
results, elapse = engine(np.array(img.convert('RGB')))
print(f'RapidOCR detected {len(results) if results else 0} text elements in {elapse}s:')
if results:
    for box, text, score in results[:15]:
        print(f'  [{score:.2f}] "{text}"')

import os
from PIL import Image
from rembg import remove

os.makedirs('output/f11_cutouts', exist_ok=True)
folder_path = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\15_Famous_Foods_Of_The_Usa'

print("Processing Slide 2 Burger cutout...")
im2 = Image.open(os.path.join(folder_path, 'Slide_02.png'))
w, h = im2.size
c2 = im2.crop((0, int(0.18*h), int(0.42*w), h))
out2 = remove(c2)
out2.save('output/f11_cutouts/s2_burger_cutout.png')
print("Burger cutout saved successfully!")

print("Processing Slide 8 Hot Dog cutout...")
im8 = Image.open(os.path.join(folder_path, 'Slide_08.png'))
w8, h8 = im8.size
c8 = im8.crop((int(0.45*w8), int(0.15*h8), int(0.95*w8), int(0.75*h8)))
out8 = remove(c8)
out8.save('output/f11_cutouts/s8_hotdog_cutout.png')
print("Hotdog cutout saved successfully!")

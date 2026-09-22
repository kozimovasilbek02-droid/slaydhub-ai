import sys
import os
import shutil
from PIL import Image
import numpy as np
from scipy.ndimage import binary_fill_holes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ppt_engine import Colors, Fonts, VisionQA, hex_to_rgb, add_slide_header, add_soft_shadow

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

folder_path = r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\15_Famous_Foods_Of_The_Usa'
os.makedirs('output/f11_cutouts', exist_ok=True)

def get_img(s_num):
    p = os.path.join(folder_path, f'Slide_{s_num:02d}.png')
    return Image.open(p)

def save_crop(im, box, name):
    w, h = im.size
    c = im.crop((int(box[0]*w), int(box[1]*h), int(box[2]*w), int(box[3]*h)))
    out = os.path.abspath(f'output/f11_cutouts/{name}.png')
    c.save(out)
    return out

def isolate_food_object(im, crop_box, bg_colors, tolerance=25):
    w, h = im.size
    crop_im = im.convert('RGBA').crop((int(crop_box[0]*w), int(crop_box[1]*h), int(crop_box[2]*w), int(crop_box[3]*h)))
    arr = np.array(crop_im, dtype=np.float32)
    rgb = arr[:, :, :3]
    
    is_bg = np.zeros((crop_im.height, crop_im.width), dtype=bool)
    for c in bg_colors:
        c_arr = np.array(c, dtype=np.float32)
        dist = np.linalg.norm(rgb - c_arr, axis=2)
        is_bg = is_bg | (dist < tolerance)
        
    fg_mask = binary_fill_holes(~is_bg)
    arr[:, :, 3] = (fg_mask.astype(np.uint8)) * 255
    return Image.fromarray(arr.astype(np.uint8), mode='RGBA')

# ==================== SLIDE 1: TITLE ====================
im1 = get_img(1)
p1_food = save_crop(im1, (0.48, 0, 1.0, 1.0), 's1_food')
s1 = prs.slides.add_slide(prs.slide_layouts[6])
s1.shapes.add_picture(p1_food, Inches(6.00), 0, width=Inches(7.33))
tb1 = s1.shapes.add_textbox(Inches(1.20), Inches(2.20), Inches(5.00), Inches(3.20))
p1 = tb1.text_frame.paragraphs[0]
p1.text = "Famous\nFoods Of\nUSA"
p1.font.name = Fonts.TITLE
p1.font.size = Pt(46)
p1.font.bold = True
p1.font.color.rgb = hex_to_rgb('#2D3748')

# ==================== SLIDE 2: HAMBURGER (SEPARATE YELLOW SHAPE + TRANSPARENT BURGER) ====================
im2 = get_img(2)
im2_cutout = isolate_food_object(im2, (0, 0.20, 0.40, 1.0), bg_colors=[[255, 192, 0], [255, 255, 255]], tolerance=25)
p2_burger_cutout = os.path.abspath('output/f11_cutouts/s2_burger_isolated.png')
im2_cutout.save(p2_burger_cutout)

s2 = prs.slides.add_slide(prs.slide_layouts[6])
# 1. Native yellow background shape (Left vertical banner)
y_banner2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(3.10), Inches(7.5))
y_banner2.fill.solid()
y_banner2.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_banner2.line.fill.background()

# 2. Transparent burger image placed independently on top
s2.shapes.add_picture(p2_burger_cutout, 0, Inches(2.25), width=Inches(4.90))

# 3. Right Text Frame
tb2 = s2.shapes.add_textbox(Inches(6.80), Inches(2.00), Inches(5.80), Inches(4.50))
tf2 = tb2.text_frame
tf2.word_wrap = True
p2_t = tf2.paragraphs[0]
p2_t.text = "Hamburger"
p2_t.font.name = Fonts.TITLE
p2_t.font.size = Pt(36)
p2_t.font.bold = True
p2_t.font.color.rgb = hex_to_rgb('#FFC000')
p2_t.space_after = Pt(18)
p2_b = tf2.add_paragraph()
p2_b.text = "The hamburger, a hallmark of American fast food, emerged in the US. It features a ground beef patty between two buns, often with veggies. This innovation streamlined the way we consume meat and bread, becoming an enduring icon of American cuisine."
p2_b.font.name = Fonts.BODY
p2_b.font.size = Pt(13.5)
p2_b.font.color.rgb = hex_to_rgb('#2D3748')
p2_b.line_spacing = 1.35

# ==================== SLIDE 3: APPLE PIE ====================
im3 = get_img(3)
p3_lf = save_crop(im3, (0.046, 0.078, 0.372, 0.943), 's3_lf')
p3_tr = save_crop(im3, (0.42, 0, 1.0, 0.43), 's3_tr')
s3 = prs.slides.add_slide(prs.slide_layouts[6])
y_b3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(4.50), Inches(2.75))
y_b3.fill.solid()
y_b3.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_b3.line.fill.background()
s3.shapes.add_picture(p3_lf, Inches(0.60), Inches(0.60), width=Inches(4.30))
s3.shapes.add_picture(p3_tr, Inches(5.60), 0, width=Inches(7.733))
tb3 = s3.shapes.add_textbox(Inches(5.70), Inches(3.70), Inches(6.80), Inches(3.20))
tf3 = tb3.text_frame
tf3.word_wrap = True
p3_t = tf3.paragraphs[0]
p3_t.text = "Apple Pie"
p3_t.font.name = Fonts.TITLE
p3_t.font.size = Pt(34)
p3_t.font.bold = True
p3_t.font.color.rgb = hex_to_rgb('#FFC000')
p3_t.space_after = Pt(16)
p3_b = tf3.add_paragraph()
p3_b.text = "Originating in England, this fruit pie, primarily apple based, gained American popularity in the 18th and 19th centuries. European immigrants introduced spices like nutmeg and cinnamon, shaping it into a cherished part of traditional American cuisine."
p3_b.font.name = Fonts.BODY
p3_b.font.size = Pt(13)
p3_b.font.color.rgb = hex_to_rgb('#2D3748')
p3_b.line_spacing = 1.3

# ==================== SLIDE 4: TATER TOTS ====================
im4 = get_img(4)
p4_top = save_crop(im4, (0, 0, 1.0, 0.27), 's4_top')
p4_bowl = save_crop(im4, (0.57, 0.40, 1.0, 1.0), 's4_bowl')
s4 = prs.slides.add_slide(prs.slide_layouts[6])
s4.shapes.add_picture(p4_top, 0, 0, width=Inches(13.333))
s4.shapes.add_picture(p4_bowl, Inches(7.60), Inches(3.00), width=Inches(5.733))
tb4 = s4.shapes.add_textbox(Inches(0.80), Inches(2.40), Inches(6.50), Inches(4.50))
tf4 = tb4.text_frame
tf4.word_wrap = True
p4_t = tf4.paragraphs[0]
p4_t.text = "Tater Tots"
p4_t.font.name = Fonts.TITLE
p4_t.font.size = Pt(36)
p4_t.font.bold = True
p4_t.font.color.rgb = hex_to_rgb('#FFC000')
p4_t.space_after = Pt(18)
p4_b = tf4.add_paragraph()
p4_b.text = "A famous American dish, the cylindrical, crispy potato patty, is enjoyed by all ages and pairs well with various dips. It's a popular item in the inventory of many frozen food companies."
p4_b.font.name = Fonts.BODY
p4_b.font.size = Pt(13.5)
p4_b.font.color.rgb = hex_to_rgb('#2D3748')
p4_b.line_spacing = 1.35

# ==================== SLIDE 5: JERKY ====================
im5 = get_img(5)
p5_pan = save_crop(im5, (0, 0, 0.27, 0.80), 's5_pan')
p5_plate = save_crop(im5, (0.69, 0.45, 1.0, 1.0), 's5_plate')
s5 = prs.slides.add_slide(prs.slide_layouts[6])
y_c5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.60), Inches(1.05), Inches(6.25), Inches(5.40))
y_c5.fill.solid()
y_c5.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c5.line.fill.background()
s5.shapes.add_picture(p5_pan, 0, 0, width=Inches(3.60))
s5.shapes.add_picture(p5_plate, Inches(9.20), Inches(3.40), width=Inches(4.133))
tb5_t = s5.shapes.add_textbox(Inches(10.50), Inches(1.20), Inches(2.50), Inches(1.00))
p5_t = tb5_t.text_frame.paragraphs[0]
p5_t.text = "Jerky"
p5_t.font.name = Fonts.TITLE
p5_t.font.size = Pt(36)
p5_t.font.bold = True
p5_t.font.color.rgb = hex_to_rgb('#FFC000')
tb5_b = s5.shapes.add_textbox(Inches(4.10), Inches(2.20), Inches(5.20), Inches(3.50))
tf5_b = tb5_b.text_frame
tf5_b.word_wrap = True
p5_b = tf5_b.paragraphs[0]
p5_b.text = "Jerky, a beloved American nomadic staple, is a dried and seasoned meat dish that requires no refrigeration. It's rich in protein and offers various flavors, making it a perfect side dish for a hearty American meal."
p5_b.font.name = Fonts.BODY
p5_b.font.size = Pt(13.5)
p5_b.font.color.rgb = hex_to_rgb('#1A202C')
p5_b.line_spacing = 1.35

# ==================== SLIDE 6: PASTA ====================
im6 = get_img(6)
p6_fork = save_crop(im6, (0.62, 0.15, 0.98, 0.98), 's6_fork')
s6 = prs.slides.add_slide(prs.slide_layouts[6])
y_c6 = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.12), Inches(7.33), Inches(6.38))
y_c6.fill.solid()
y_c6.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c6.line.fill.background()
s6.shapes.add_picture(p6_fork, Inches(8.20), Inches(1.12), width=Inches(4.80))
tb6 = s6.shapes.add_textbox(Inches(1.20), Inches(1.80), Inches(5.20), Inches(5.00))
tf6 = tb6.text_frame
tf6.word_wrap = True
p6_t = tf6.paragraphs[0]
p6_t.text = "Pasta"
p6_t.font.name = Fonts.TITLE
p6_t.font.size = Pt(32)
p6_t.font.bold = True
p6_t.font.color.rgb = hex_to_rgb('#1A202C')
p6_t.alignment = PP_ALIGN.CENTER
p6_t.space_after = Pt(16)
p6_b = tf6.add_paragraph()
p6_b.text = "Macaroni and cheese, an American take on pasta, is a beloved comfort food in the USA. Quick and satisfying, it's a staple in offices and homes. This dish combines macaroni with cheddar cheese, offering a mild yet cheesy delight. You'll find it on menus across the country, making it a must-try for cheese lovers."
p6_b.font.name = Fonts.BODY
p6_b.font.size = Pt(12)
p6_b.font.color.rgb = hex_to_rgb('#1A202C')
p6_b.alignment = PP_ALIGN.CENTER
p6_b.line_spacing = 1.3

# ==================== SLIDE 7: BUFFALO CHICKEN WINGS (SEPARATE YELLOW SHAPE + CLEAN PLATTER) ====================
im7 = get_img(7)
p7_wings = save_crop(im7, (0.238, 0, 0.612, 1.0), 's7_wings_pure')
s7 = prs.slides.add_slide(prs.slide_layouts[6])

# 1. Native Yellow Card Shape
y_c7 = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.95), Inches(4.20), Inches(3.80))
y_c7.fill.solid()
y_c7.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c7.line.fill.background()

tb7_t = s7.shapes.add_textbox(Inches(0.40), Inches(2.40), Inches(3.40), Inches(2.80))
tf7_t = tb7_t.text_frame
tf7_t.word_wrap = True
p7_t = tf7_t.paragraphs[0]
p7_t.text = "Buffalo\nChicken\nWings"
p7_t.font.name = Fonts.TITLE
p7_t.font.size = Pt(28)
p7_t.font.bold = True
p7_t.font.color.rgb = hex_to_rgb('#FFFFFF')
p7_t.alignment = PP_ALIGN.CENTER

# 2. Pure wings platter image
s7.shapes.add_picture(p7_wings, Inches(3.15), 0, width=Inches(5.00))

# 3. Right description text
tb7_b = s7.shapes.add_textbox(Inches(8.60), Inches(1.60), Inches(4.20), Inches(4.50))
tf7_b = tb7_b.text_frame
tf7_b.word_wrap = True
p7_b = tf7_b.paragraphs[0]
p7_b.text = "Buffalo wings are a cherished American favorite - crispy chicken wings drenched in cayenne pepper hot sauce and butter. They're served with veggies and dips like ranch or blue cheese. Restaurants often offer a baked, healthier version. A true USA classic."
p7_b.font.name = Fonts.BODY
p7_b.font.size = Pt(13)
p7_b.font.color.rgb = hex_to_rgb('#2D3748')
p7_b.line_spacing = 1.35

# ==================== SLIDE 8: HOT DOG (SEPARATE NATIVE YELLOW BLOCKS + CLEAN HOTDOG) ====================
im8 = get_img(8)
im8_cutout = isolate_food_object(im8, (0.45, 0.15, 0.95, 0.75), bg_colors=[[255, 255, 255], [255, 192, 0]], tolerance=20)
p8_hotdog_isolated = os.path.abspath('output/f11_cutouts/s8_hotdog_isolated.png')
im8_cutout.save(p8_hotdog_isolated)

s8 = prs.slides.add_slide(prs.slide_layouts[6])
# 1. Native Yellow Left Rectangle
y_c8 = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(6.05), Inches(7.5))
y_c8.fill.solid()
y_c8.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c8.line.fill.background()

# 2. Native Yellow Bottom-Right Accent Shape
y_dec8 = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.50), Inches(3.00), Inches(5.833), Inches(4.50))
y_dec8.fill.solid()
y_dec8.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_dec8.line.fill.background()

# 3. Transparent Plate Hotdog
s8.shapes.add_picture(p8_hotdog_isolated, Inches(5.80), Inches(1.30), width=Inches(6.50))

# 4. Left Text Frame
tb8 = s8.shapes.add_textbox(Inches(0.60), Inches(1.20), Inches(5.00), Inches(5.50))
tf8 = tb8.text_frame
tf8.word_wrap = True
p8_t = tf8.paragraphs[0]
p8_t.text = "Hot Dog"
p8_t.font.name = Fonts.TITLE
p8_t.font.size = Pt(34)
p8_t.font.bold = True
p8_t.font.color.rgb = hex_to_rgb('#1A202C')
p8_t.space_after = Pt(16)
p8_b = tf8.add_paragraph()
p8_b.text = "The hot dog, rooted in Germany, is now an iconic American snack. It's a sausage in a bun with toppings like ketchup and mustard. Introduced by German immigrants in the late 1800s, it's a national favorite available at stands throughout the USA."
p8_b.font.name = Fonts.BODY
p8_b.font.size = Pt(13)
p8_b.font.color.rgb = hex_to_rgb('#1A202C')
p8_b.line_spacing = 1.35

# ==================== SLIDE 9: MEATLOAF ====================
im9 = get_img(9)
p9_photo = save_crop(im9, (0.02, 0.03, 0.54, 0.54), 's9_photo')
s9 = prs.slides.add_slide(prs.slide_layouts[6])
y_c9 = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.20), Inches(1.25), Inches(6.133), Inches(3.50))
y_c9.fill.solid()
y_c9.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c9.line.fill.background()
s9.shapes.add_picture(p9_photo, Inches(0.30), Inches(0.25), width=Inches(6.90))
tb9_t = s9.shapes.add_textbox(Inches(7.60), Inches(1.80), Inches(5.40), Inches(2.50))
tf9_t = tb9_t.text_frame
tf9_t.word_wrap = True
p9_t = tf9_t.paragraphs[0]
p9_t.text = "Meatloaf | Cooked Meat\nFor A Wholesome Meal"
p9_t.font.name = Fonts.TITLE
p9_t.font.size = Pt(28)
p9_t.font.bold = True
p9_t.font.color.rgb = hex_to_rgb('#FFFFFF')

def add_bullet_item(slide, x, y, w, text):
    bd = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y + 0.05), Inches(0.32), Inches(0.32))
    bd.fill.solid()
    bd.fill.fore_color.rgb = hex_to_rgb('#FFC000')
    bd.line.fill.background()
    p_ic = bd.text_frame.paragraphs[0]
    p_ic.text = ">"
    p_ic.font.bold = True
    p_ic.font.color.rgb = hex_to_rgb('#FFFFFF')
    p_ic.alignment = PP_ALIGN.CENTER
    tb = slide.shapes.add_textbox(Inches(x + 0.45), Inches(y), Inches(w - 0.45), Inches(1.40))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = Fonts.BODY
    p.font.size = Pt(9.5)
    p.font.color.rgb = hex_to_rgb('#4A5568')

add_bullet_item(s9, 0.40, 5.60, 5.80, "Meatloaf, a simple and versatile dish, features spiced minced meat baked or smoked in a loaf shape. It often includes bread crumbs, eggs, cheese, and vegetables for moisture.")
add_bullet_item(s9, 6.80, 5.60, 5.80, "This American take on a traditional European dish, with roots in Scrapple, is a nutritious lunch option. Enjoy it with a side of tomato ketchup.")

# ==================== SLIDE 10: GRITS ====================
im10 = get_img(10)
p10_bowl = save_crop(im10, (0.365, 0, 1.0, 0.62), 's10_bowl')
s10 = prs.slides.add_slide(prs.slide_layouts[6])
y_c10 = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.35), Inches(1.00), Inches(4.50), Inches(2.70))
y_c10.fill.solid()
y_c10.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c10.line.fill.background()
tb10_t = s10.shapes.add_textbox(Inches(0.70), Inches(1.25), Inches(3.80), Inches(2.20))
tf10_t = tb10_t.text_frame
tf10_t.word_wrap = True
p10_t = tf10_t.paragraphs[0]
p10_t.text = "Grits | Porridge\nFor A Healthy\nStart"
p10_t.font.name = Fonts.TITLE
p10_t.font.size = Pt(26)
p10_t.font.bold = True
p10_t.font.color.rgb = hex_to_rgb('#FFFFFF')
s10.shapes.add_picture(p10_bowl, Inches(4.85), 0, width=Inches(8.48))
tb10_b = s10.shapes.add_textbox(Inches(1.00), Inches(5.40), Inches(11.333), Inches(1.50))
tf10_b = tb10_b.text_frame
tf10_b.word_wrap = True
p10_b = tf10_b.paragraphs[0]
p10_b.text = "'Grits,' originating from Old English, refers to boiled cornmeal porridge, a beloved American classic, particularly in coastal communities. Commonly enjoyed for breakfast or dinner, it pairs with ingredients like eggs, bacon, fish, and country ham, offering both sweet and savory variations."
p10_b.font.name = Fonts.BODY
p10_b.font.size = Pt(11)
p10_b.font.color.rgb = hex_to_rgb('#4A5568')
p10_b.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 11: BARBECUE RIBS ====================
im11 = get_img(11)
p11_ribs = save_crop(im11, (0.415, 0.34, 0.925, 0.86), 's11_ribs')
s11 = prs.slides.add_slide(prs.slide_layouts[6])
y_bar11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.40), Inches(7.5))
y_bar11.fill.solid()
y_bar11.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_bar11.line.fill.background()
y_r11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.10), Inches(3.50), Inches(4.233), Inches(4.00))
y_r11.fill.solid()
y_r11.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_r11.line.fill.background()
tb11_t = s11.shapes.add_textbox(Inches(5.70), Inches(1.10), Inches(6.80), Inches(1.50))
tf11_t = tb11_t.text_frame
tf11_t.word_wrap = True
p11_t = tf11_t.paragraphs[0]
p11_t.text = "Barbecue Ribs Spicy Ribs For A\nDelicious Surprise"
p11_t.font.name = Fonts.TITLE
p11_t.font.size = Pt(25)
p11_t.font.bold = True
p11_t.font.color.rgb = hex_to_rgb('#2D3748')
s11.shapes.add_picture(p11_ribs, Inches(5.50), Inches(2.70), width=Inches(6.80))
add_bullet_item(s11, 0.80, 1.10, 4.40, "A beloved American classic, barbecue ribs, made from lamb and pig ribs, are prepared using various methods and paired with diverse barbecue sauces.")
add_bullet_item(s11, 0.80, 3.20, 4.40, "These savory ribs promise to lift your spirits and tantalize your palate, holding a cherished spot in US culinary tradition.")
add_bullet_item(s11, 0.80, 5.20, 4.40, "Beyond hamburgers and apple pie, barbecue ribs deliver a mouthwatering, mood-enhancing experience, celebrating the rich flavors of American cuisine.")

# ==================== SLIDE 12: CUBAN SANDWICH ====================
im12 = get_img(12)
p12_sand = save_crop(im12, (0.058, 0.33, 0.545, 0.82), 's12_sand')
s12 = prs.slides.add_slide(prs.slide_layouts[6])
y_b12 = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.90), Inches(5.65), Inches(3.60))
y_b12.fill.solid()
y_b12.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_b12.line.fill.background()
tb12_t = s12.shapes.add_textbox(Inches(0.80), Inches(0.80), Inches(11.733), Inches(1.00))
p12_t = tb12_t.text_frame.paragraphs[0]
p12_t.text = "Cuban Sandwich | Healthy Snack In The US"
p12_t.font.name = Fonts.TITLE
p12_t.font.size = Pt(28)
p12_t.font.bold = True
p12_t.font.color.rgb = hex_to_rgb('#2D3748')
p12_t.alignment = PP_ALIGN.CENTER
s12.shapes.add_picture(p12_sand, Inches(0.77), Inches(2.50), width=Inches(6.50))
tb12_b = s12.shapes.add_textbox(Inches(7.95), Inches(2.60), Inches(4.80), Inches(4.00))
tf12_b = tb12_b.text_frame
tf12_b.word_wrap = True
p12_b = tf12_b.paragraphs[0]
p12_b.text = "A Cuban sandwich, featuring buttered Cuban bread, mustard, Swiss cheese, pickles, and meats like roast pork, salami, or glazed ham, hails from Cuba but became a Tampa specialty in Florida. Perfect for snack time, it's available at chains like Publix."
p12_b.font.name = Fonts.BODY
p12_b.font.size = Pt(13)
p12_b.font.color.rgb = hex_to_rgb('#4A5568')
p12_b.line_spacing = 1.35

# ==================== SLIDE 13: CHOP SUEY ====================
im13 = get_img(13)
p13_top = save_crop(im13, (0.225, 0, 0.775, 0.53), 's13_top')
s13 = prs.slides.add_slide(prs.slide_layouts[6])
y_c13 = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.50), Inches(13.333), Inches(3.90))
y_c13.fill.solid()
y_c13.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c13.line.fill.background()
s13.shapes.add_picture(p13_top, Inches(3.00), 0, width=Inches(7.33))
tb13_t = s13.shapes.add_textbox(Inches(0.80), Inches(4.30), Inches(11.733), Inches(1.00))
p13_t = tb13_t.text_frame.paragraphs[0]
p13_t.text = "Chop Suey | Soupy Dish For Starters Or Snacks"
p13_t.font.name = Fonts.TITLE
p13_t.font.size = Pt(28)
p13_t.font.bold = True
p13_t.font.color.rgb = hex_to_rgb('#FFFFFF')
p13_t.alignment = PP_ALIGN.CENTER
add_bullet_item(s13, 0.60, 5.80, 5.80, "Despite its strong presence in Chinese cuisine, this dish actually emerged in America when the country became home to a sizable Chinese population. The dish's creation is shrouded in various intriguing tales.")
add_bullet_item(s13, 6.90, 5.80, 5.80, "This flavorful concoction unites meat, eggs, a medley of sautéed vegetables (cabbage, celery, bean sprouts, and more), all served in a delectable sauce. Often enjoyed with rice, it boasts a mild spiciness and is an excellent option as a main course.")

# ==================== SLIDE 14: AVOCADO TOAST ====================
im14 = get_img(14)
p14_toast = save_crop(im14, (0.375, 0.325, 0.965, 0.92), 's14_toast')
s14 = prs.slides.add_slide(prs.slide_layouts[6])
y_r14 = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.65), Inches(3.50), Inches(4.68), Inches(4.00))
y_r14.fill.solid()
y_r14.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_r14.line.fill.background()
tb14_t = s14.shapes.add_textbox(Inches(0.60), Inches(0.80), Inches(10.00), Inches(1.50))
tf14_t = tb14_t.text_frame
tf14_t.word_wrap = True
p14_t = tf14_t.paragraphs[0]
p14_t.text = "Avocado Toast | USA's Traditional Food That\nEveryone Is Hyping About"
p14_t.font.name = Fonts.TITLE
p14_t.font.size = Pt(25)
p14_t.font.bold = True
p14_t.font.color.rgb = hex_to_rgb('#2D3748')
s14.shapes.add_picture(p14_toast, Inches(4.90), Inches(2.40), width=Inches(7.80))
add_bullet_item(s14, 0.60, 2.50, 4.00, "Avocado, the nation's native fruit, is passionately embraced. Avocado toast, endorsed by celebrities, offers a wholesome snack choice.")
add_bullet_item(s14, 0.60, 4.10, 4.00, "This dish includes mashed avocados seasoned with olive oil, hummus, tomatoes, and lemon juice, ideal for an American breakfast.")
add_bullet_item(s14, 0.60, 5.70, 4.00, "Don't miss this delectable treat - relish it for breakfast with a cup of American-style tea for a flavorful and nutritious start.")

# ==================== SLIDE 15: BAKED ZITI ====================
im15 = get_img(15)
p15_photo = save_crop(im15, (0, 0, 0.64, 0.63), 's15_photo')
s15 = prs.slides.add_slide(prs.slide_layouts[6])
y_c15 = s15.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.533), Inches(0.80), Inches(4.80), Inches(3.10))
y_c15.fill.solid()
y_c15.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_c15.line.fill.background()
s15.shapes.add_picture(p15_photo, 0, 0, width=Inches(8.533))
tb15_t = s15.shapes.add_textbox(Inches(8.80), Inches(1.10), Inches(4.20), Inches(2.50))
tf15_t = tb15_t.text_frame
tf15_t.word_wrap = True
p15_t = tf15_t.paragraphs[0]
p15_t.text = "Baked Ziti |\nClassical Dish\nFrom\nItalian-American\nCuisine"
p15_t.font.name = Fonts.TITLE
p15_t.font.size = Pt(22)
p15_t.font.bold = True
p15_t.font.color.rgb = hex_to_rgb('#FFFFFF')
tb15_b = s15.shapes.add_textbox(Inches(1.00), Inches(5.30), Inches(11.333), Inches(1.50))
tf15_b = tb15_b.text_frame
tf15_b.word_wrap = True
p15_b = tf15_b.paragraphs[0]
p15_b.text = "While pasta is typically associated with Italy, baked ziti has American roots, specifically in South Florida. It's a delightful fusion of ziti pasta, a rich tomato-based sauce, and a medley of cheeses (ricotta, parmesan, and mozzarella). It's often personalized with additions like bell peppers, meat, mushrooms, and onions. If you're a cheese enthusiast, this dish is a must-try."
p15_b.font.name = Fonts.BODY
p15_b.font.size = Pt(11)
p15_b.font.color.rgb = hex_to_rgb('#4A5568')
p15_b.alignment = PP_ALIGN.CENTER

# ==================== SLIDE 16: CHICKEN FRIED STEAK ====================
im16 = get_img(16)
p16_steak = save_crop(im16, (0.475, 0.24, 0.99, 0.75), 's16_steak')
s16 = prs.slides.add_slide(prs.slide_layouts[6])
y_r16 = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.70), 0, Inches(4.633), Inches(7.5))
y_r16.fill.solid()
y_r16.fill.fore_color.rgb = hex_to_rgb('#FFC000')
y_r16.line.fill.background()
s16.shapes.add_picture(p16_steak, Inches(6.30), Inches(1.80), width=Inches(6.80))
tb16 = s16.shapes.add_textbox(Inches(0.80), Inches(1.80), Inches(5.20), Inches(5.00))
tf16 = tb16.text_frame
tf16.word_wrap = True
p16_t = tf16.paragraphs[0]
p16_t.text = "Chicken Fried Steak |\nClassical American\nCutlet"
p16_t.font.name = Fonts.TITLE
p16_t.font.size = Pt(28)
p16_t.font.bold = True
p16_t.font.color.rgb = hex_to_rgb('#2D3748')
p16_t.space_after = Pt(16)
p16_b = tf16.add_paragraph()
p16_b.text = "Indulge in this American twist on Austrian and Italian classics—a tenderized beef steak dipped in batter, seasoned flour, and deep-fried, known as country-fried steak. Enjoy it with gravy, veggies, hashbrowns, and toast. Explore the USA's culinary legacy. Adotrip offers budget-friendly travel assistance and itineraries for your memorable trip."
p16_b.font.name = Fonts.BODY
p16_b.font.size = Pt(12)
p16_b.font.color.rgb = hex_to_rgb('#4A5568')
p16_b.line_spacing = 1.3

out_local = os.path.abspath('output/f11_presentation_precision.pptx')
out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('Saved local presentation to:', out_local)

# Try copying to destination if not locked by PowerPoint
dst_master = os.path.abspath(r'G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar\15_Famous_Foods_Of_The_Usa\presentation_95_precision.pptx')
try:
    shutil.copy(out_local, dst_master)
    print('Master presentation updated at:', dst_master)
except Exception as e:
    print('Note: destination file is open in PowerPoint. Saved to local output successfully.')

# Export previews
exported = VisionQA.export_slides_to_png(out_local, output_dir='output/folder11_perfect_previews')
print(f'Exported {len(exported)} perfect preview slides for Folder 11.')

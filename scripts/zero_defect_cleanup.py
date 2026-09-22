import os
import glob
import re
import sys
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"

replacements_map = {
    "PERVЫY ETAP": "BIRINCHI BOSQICH",
    "PERVY ETAP": "BIRINCHI BOSQICH",
    "PERVY ETAP": "BIRINCHI BOSQICH",
    "PERVYY ETAP": "BIRINCHI BOSQICH",
    "ot elektrona do pozitrona (1897-1932 gg.)": "elektrondan pozitroncha (1897-1932 yy.)",
    "ITOGI UROKA": "DARS XULOSASI",
    "ITOG UROKA": "DARS XULOSASI",
    "Podvedenie itogov uroka:": "Dars xulosalarini chiqarish:",
    "DOMAShNEE ZADANIE": "UYGA VAZIFA",
    "DOMASHNEE ZADANIE": "UYGA VAZIFA",
    "1. Chto my izuchili v 8-om klasse?": "1. 8-sinfda nimani o'rgangan edik?",
    "1. Chto my izuchili v 7-om klasse?": "1. 7-sinfda nimani o'rgangan edik?",
    "a) teplovye yavleniya": "a) issiqlik hodisalari",
    "b) elektricheskie i magnitnye yavleniya": "b) elektr va magnit hodisalari",
    "b) elektricheskie i mag": "b) elektr va magnit",
    "a) dvijenie": "a) harakat",
    "Vpervye vesomost vozduxa privela lyudey v zameshatelstvo v 1638 godu, kogda ne udalas popыtka ustroit fontan": "1638-yilda Florensiyadagi saroy bog'ida favvora qurish urinishi muvaffaqiyatsiz tugagach, havoning og'irligi birinchi marta olimlarni hayratda qoldirdi",
    "Kak obrazuetsya vysokoe i nizkoe atmosfernoe davlenie.": "Yuqori va past atmosfera bosimi qanday hosil bo'ladi.",
    "Oblast vysokogo atmosfernogo davleniya obrazuetsya pri nisxodyashix potokax vozduxa.": "Yuqori atmosfera bosimi sohasi havoning pastga yo'nalgan oqimlari natijasida hosil bo'ladi.",
    "Xolodnyy vozdux u poverxnosti Zemli ne mojet skaplivatsya v odnom meste. On nachinaet rastekatsya v storony.": "Yer yuzasidagi sovuq havo bir joyda to'planib qola olmaydi va har tomonga tarqala boshlaydi.",
    "Fizicheskiy\xa0– eto silnyy xolod ili nevynosimaya jara, ponijenie ili povyshenie atmosfernogo davleniya": "Jismoniy stress – kuchli sovuq yoki chidab bo'lmas issiqlik, atmosfera bosimining pasayishi yoki ko'tarilishi",
    "Fizicheskiy\xa0 eto silnyy xolod ili nevynosimaya jara, ponijenie ili povyshenie atmosfernogo davleniya": "Jismoniy stress – kuchli sovuq yoki chidab bo'lmas issiqlik, atmosfera bosimining pasayishi yoki ko'tarilishi",
    "Faza trevogi\xa0– sostoyanie neopredelennosti i straxa v svyazi s priblijayusheysya ugrozoy": "Tashvish fazasi – yaqinlashib kelayotgan xavf tufayli noaniqlik va qo'rquv holati",
    "Faza trevogi\xa0 sostoyanie neopredelennosti i straxa v svyazi s priblijayusheysya ugrozoy": "Tashvish fazasi – yaqinlashib kelayotgan xavf tufayli noaniqlik va qo'rquv holati",
    "Kakie byvayut fazy stressa?": "Stressning qanday fazalari bor?",
    "Snijaetsya kontsentratsiya vnimaniya, chto vlechet za soboy uxudshenie pamyati;": "Diqqat konsentratsiyasi pasayadi, bu esa xotiraning yomonlashishiga olib keladi;",
    "Negativnye posledstviya stressa": "Stressning salbiy oqibatlari",
    "Krolik\xa0– passivnaya reaktsiya na stressovuyu situatsiyu. Stress lishaet vozmojnosti soprotivlyatsya, chelovek stanovitsya bespomoshnыm.": "Quyon (Krolik) – stress holatiga passiv reaksiya. Stress qarshilik ko'rsatish qobiliyatidan mahrum qiladi, inson nochor bo'lib qoladi.",
    "Krolik\xa0 passivnaya reaktsiya na stressovuyu situatsiyu. Stress lishaet vozmojnosti soprotivlyatsya, chelovek stanovitsya bespomoshnыm.": "Quyon (Krolik) – stress holatiga passiv reaksiya. Stress qarshilik ko'rsatish qobiliyatidan mahrum qiladi, inson nochor bo'lib qoladi.",
    "Kakova vzaimosvyaz konflikta i stressa?": "Nizo va stress o'rtasida qanday bog'liqlik bor?",
    "O'tkazgichlarni ketma-ket ulash bayram galarida": "O'tkazgichlarni ketma-ket ulash bayram gulchambarlarida (girlyandalarida)",
    "Srednee vremya jizni": "O'rtacha yashash vaqti",
    "Vremya, v techenie kotorogo jivet chastitsa": "Zarracha yashaydigan vaqt",
    "Mezonы": "Mezonlar",
    "Nuklonы": "Nuklonlar",
    "Myuonы": "Myuonlar",
    "Internet-resursы:": "Internet manbalari:",
    "Ot chego zavisit vnutrennyaya energiya?": "Ichki energiya nimaga bog'liq?",
    "Po kakoy formule rasschityvaetsya ?": "Qaysi formula bo'yicha hisoblanadi?",
    "§ 75\nZnat opredeleniya": "§ 75\nTa'riflarni yod olish"
}

files = sorted(glob.glob(os.path.join(folder, "*.pptx")))

for fpath in files:
    fname = os.path.basename(fpath)
    prs = Presentation(fpath)
    modified = False
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text
                    new_t = t
                    for k, v in replacements_map.items():
                        if k in new_t:
                            new_t = new_t.replace(k, v)
                    # General regex cleanup for broken characters
                    new_t = new_t.replace('З', 'Z').replace('з', 'z').replace('ы', 'y').replace('', "'")
                    if new_t != t:
                        p.text = new_t
                        modified = True
                        
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for p in cell.text_frame.paragraphs:
                                t = p.text
                                new_t = t
                                for k, v in replacements_map.items():
                                    if k in new_t:
                                        new_t = new_t.replace(k, v)
                                new_t = new_t.replace('З', 'Z').replace('з', 'z').replace('ы', 'y').replace('', "'")
                                if new_t != t:
                                    p.text = new_t
                                    modified = True
                                    
    if modified:
        prs.save(fpath)
        print(f"Cleaned & Saved: {fname}")

print("\nZero-defect cleanup completed!")

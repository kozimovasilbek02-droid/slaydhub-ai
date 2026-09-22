import os
from pptx import Presentation

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"

# 1. Fizika Slide 3
f1 = os.path.join(folder, "Fizika - Moddiy nuqta va sanoq sistemasi (9-sinf).pptx")
prs1 = Presentation(f1)
s3 = prs1.slides[2]
for shape in s3.shapes:
    if shape.has_text_frame:
        if "7-sinfda" in shape.text_frame.text:
            shape.text_frame.text = "1. 7-sinfda nimani o'rgangan edik?\na) harakat\nb) kuchlar\nv) bosim\ng) ish va energiya\n2. Bu bilimlarning deyarli barchasi bizga kerak bo'ladi\n"
prs1.save(f1)

# 2. Menejment
f2 = os.path.join(folder, "Menejment - Nizolar va stresslarni boshqarish.pptx")
prs2 = Presentation(f2)

# Slide 8
s8 = prs2.slides[7]
for shape in s8.shapes:
    if shape.has_text_frame:
        if "Jismoniy stress" in shape.text_frame.text:
            shape.text_frame.text = (
                "Jismoniy stress – kuchli sovuq yoki chidab bo'lmas issiqlik, atmosfera bosimining pasayishi yoki ko'tarilishi.\n"
                "Kimyoviy – zaharli moddalar ta'siri.\n"
                "Ruhiy – kuchli salbiy yoki ijobiy hissiyotlar.\n"
                "Biologik – jarohatlar, virusli kasalliklar, mushak yuklamalari.\n"
                "Natijasiga ko'ra psixologiyada quyidagi stress turlari ajratiladi:\n"
                "Eustresslar (ijobiy stress) – har birimizning muvaffaqiyatli faoliyatimiz uchun ma'lum miqdorda stress zarur. Aynan u bizning rivojlanishimizning harakatlantiruvchi kuchi hisoblanadi. Bu holatni «uyg'onish reaksiyasi» deb atash mumkin.\n"
                "Distresslar (zararli stress) – kritik darajadagi zo'riqishda paydo bo'ladi. Aynan shu holat stress haqidagi barcha salbiy tushunchalarni o'zida aks ettiradi.\n"
            )
        elif "Vidy stressa" in shape.text_frame.text:
            shape.text_frame.text = "Stress turlari"

# Slide 11
s11 = prs2.slides[10]
for shape in s11.shapes:
    if shape.has_text_frame:
        if "Diqqat" in shape.text_frame.text:
            shape.text_frame.text = (
                "Diqqat konsentratsiyasi pasayadi, bu esa xotiraning yomonlashishiga olib keladi;\n"
                "Beparvolik va pala-partishlik paydo bo'ladi, bu esa o'ylamasdan qaror qabul qilish xavfini oshiradi;\n"
                "Mehnat qobiliyatining pasayishi va tez charchash bosh miya yarimsharlari po'stlog'idagi neyronlararo aloqalarning buzilishi oqibati bo'lishi mumkin;\n"
                "Salbiy his-tuyg'ular ustunlik qiladi – o'z holatidan, ishidan, sherigidan, tashqi ko'rinishidan umumiy norozilik depressiya rivojlanish xavfini oshiradi;\n"
                "Atrofdagilar bilan muloqotni qiyinlashtiradigan va nizoli vaziyatni cho'zadigan asabiylashish va tajovuzkorlik;\n"
                "Alkogol, antidepressantlar va dori vositalari orqali holatni yengillashtirishga urinish;\n"
                "O'z-o'zini baholashning pasayishi, o'z kuchiga ishonchsizlik;\n"
                "Oilaviy va shaxsiy hayotdagi muammolar;\n"
            )

# Slide 14
s14 = prs2.slides[13]
for shape in s14.shapes:
    if shape.has_text_frame:
        if "Quyon" in shape.text_frame.text or "Krolik" in shape.text_frame.text:
            shape.text_frame.text = (
                "\nQuyon (Krolik) – stress holatiga passiv reaksiya. Stress qarshilik ko'rsatish imkoniyatidan mahrum qiladi, inson o'zini nochor his qiladi.\n"
                "Arslon (Lev) – stress qisqa vaqt ichida organizmning barcha zaxiralarini ishga solishga majbur qiladi. Inson vaziyatga shiddatli va hissiy munosabat bildirib, uni hal qilish uchun keskin «sakrash» qiladi.\n"
                "Ho'kiz (Vol) – inson o'z aqliy va psixologik resurslarini oqilona boshqaradi, shuning uchun stress holatida ham uzoq vaqt unumli yashashi va ishlashi mumkin. Bu strategiya neyrofiziologiya nuqtai nazaridan eng asosli va eng samaralisidir.\n"
            )
        elif "Vydelyayut" in shape.text_frame.text:
            shape.text_frame.text = "Stressga munosabat bildirishning uchta strategiyasi ajratiladi"

prs2.save(f2)
print("Updated Fizika and Menejment to 100% clean Uzbek!")

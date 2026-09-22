# SlaydHub AI — Agent Ishlab Chiqish va Qat'iy Sifat Nazorati Standartlari (GEMINI.md)

Ushbu qoidalar SlaydHub AI loyihasida ishlovchi har qanday sun'iy intellekt agenti uchun MAJBURIY QONUNDIR.
Agent har qanday taqdimot yoki kod yaratganidan so'ng, foydalanuvchiga "Tayyor", "Bajarildi" yoki natija hisobotini yozishdan oldin quyidagi 5 ta filtrdan o'tishi SHART.

---

## 🛑 PROTOKOL: FOYDALANUVCHIGA TOPSHIRASHDAN OLDINGI VIZUAL VA TEXNIK NAZORAT

### 1. Vizual Audit (Visual Inspection via `view_file`):
1. **Ko'r-ko'rona "Ishladi" Deyish Taqiq**: Faqat kodning `exit code 0` bergani tizim ishlaganini anglatmaydi.
2. **Prevyularni O'zing Ochib Ko'r**: Yaratilgan taqdimot slaydlarining JPG prevyularini (`output/preview_cache/.../slide_<N>.jpg`) agent o'zining `view_file` asbobi orqali kamida 3-4 ta slaydini o'z ko'zi bilan ko'rib chiqishi SHART.
3. **Chegaradan Chiqmaslik (Zero Overflow)**:
   - Sarlavha yoki matn qutilari slaydning o'ng chekkasidan (width > 960pt) yoki pastki chekkasidan (height > 540pt) 1 piksel ham chiqib ketmasligi kerak.
   - Hech bir so'z bo'linib, qator tashlab buzilmasligi kerak.
4. **Shakllar To'qnashuvi (Spatial Collision)**:
   - Matn qutilari rasm, 3D vektor elementi yoki boshqa karta ustiga chiqib ketmasligi shart.
5. **Kvadrat Qutilar Taqiqi (No Narrow Square Slivers)**:
   - Matnni tor kvadrat qutilarga siqish qat'iyan man etiladi.
   - Kartalar keng to'g'ri to'rtburchak (aspect ratio >= 3:1 yoki 4:1, balandligi kamida 100–135pt) bo'lishi va matn qulay nafas olishi kerak.

---

### 2. Matn va Nusxa Sifati (McKinsey Copywriting Standarti):
1. **"Bold Lead-In" (Claim + Evidence) Qoidasi**:
   - Har bir tezis, karta izohi yoki punkt quruq oddiy matn bo'lmasligi shart.
   - U albatta **[Qalin Kalit So'z / Aniq Hukm]:** bilan boshlanishi, undan so'ng oddiy shriftda aniq ilmiy dalil, parametr va raqamlar kelishi lozim.
2. **Action-Oriented Headlines**:
   - Slayd sarlavhasi shunchaki umumiy so'z ("Reja", "Kriptografiya") emas, to'liq ilmiy xulosaviy jumla bo'lishi shart (masalan: "Post-Kvant Standartlari RSA Zaifliklarini To'liq Bartaraf Etadi").
3. **Sun'iy Kesilishlar Yo'qligi (No Truncation)**:
   - Matnlar 12-14 so'zdan keyin kesilib, oxiriga `...` qo'yib tashlanishi qat'iyan man etiladi. Fikr to'liq (18–28 so'z) ifodalanishi lozim.
4. **Zero AI Fluff**:
   - "Ta'kidlash joizki", "Ushbu slayd doirasida", "Dolzarb ahamiyatga ega" kabi sun'iy intellekt klishelari butunlay taqiqlanadi. Aniq raqamlar (99.9%, 2,500 kubit, 1.2 ms) ishlatiladi.

---

### 3. Shablon va XML Tozaligi:
1. **Reklama va Suvbelgilar Tozalanganligi**:
   - "PresentationGO", "SlideEgg", "SlidesCarnival", "Logotype™" kabi begona belgilar master va layoutlardan tozalangan bo'lishi shart.
2. **Bo'sh Ghost Pleysxolderlar Yo'qligi**:
   - PowerPointda "Click to edit title" yoki "Текст слайда" kabi qoldiq shakllar XML daraxtidan jismonan o'chirilgan bo'lishi shart.
3. **Autentik Shablonning 3D Dizayni Saqlanishi**:
   - Tizim hech qachon oq bo'sh slaydga tushib qolmasligi, asl shablonning barcha fon va grafikalari saqlanishi shart.

---

### 4. Texnik Tayyorgarlik:
1. Diskda haqiqiy `.pptx` va `.pdf` fayllari yaratilganligi va hajmi > 50KB ekanligi tasdiqlanishi kerak.
2. Streamlit ilovasida Step 5 (yuklab olish va ochish tugmalari) muvaffaqiyatli render bo'lishi lozim.
3. Tizim testlari (`test_all_functions.py` va `test_real_pptx_button_workflow.py`) 100% muvaffaqiyat bilan o'tgan bo'lishi shart.

---

## ⚡ QAT'IY QOIDANING KUCHI:
Agar yuqoridagi qoidalardan bittasi bo'yicha ham kamchilik aniqlansa — agent foydalanuvchiga "Hammasi tayyor" deb javob qaytara olmaydi.
Agent avval xatoni kodda tuzatadi, slaydni qayta chizadi, o'zi `view_file` orqali ko'rib 100% qoniqqanidan keyingina foydalanuvchiga topshiradi!

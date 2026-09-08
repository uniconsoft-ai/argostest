# 🏛️ ARGOS.UZ Web Vakansiyalar Tahlilchisi va Word Ko'ruvchi (In-Browser Word Viewer)

`vacancy.argos.uz` platformasidagi barcha vakansiyalarni tahlil qilish, aniq filtrlash va hosil qilingan rasmiy **Microsoft Word (.docx)** hujjatlarini to'g'ridan-to'g'ri brauzer ichida ko'rish uchun mo'ljallangan zamonaviy mahalliy web ilova.

---

## 🌟 Asosiy Imkoniyatlar

1. **In-Browser Word Document Viewer (Vebda to'g'ridan-to'g'ri Word ko'rish):**
   - Hech qanday MS Word yoki uchinchi tomon dasturlari o'rnatilmagan bo'lsa ham, Word (`.docx`) faylini to'liq sahifalari, jadvallari, nishonlari va ranglari bilan to'g'ridan-to'g'ri brauzer ichida ko'rish (`docx-preview.js` yordamida).
   - **Ikki xil ko'rish rejimi:**
     - 📄 **Sahifali ko'rinish (Page View):** Asl MS Word A4 sahifasi ko'rinishi (chegaralar, sahifa soyalari, bo'linishlar).
     - 📖 **Toza o'qish rejimi (Reading Mode):** Matn va jadvallarni keng ekranli o'qish formati (`mammoth`).
   - **Boshqaruv vositalari:** Kattalashtirish / Kichiklashtirish (Zoom 40%-200%), Kenglikka moslash (Fit Width), To'liq ekran, Vebdan to'g'ridan-to'g'ri chop etish (Print) va `.docx` yuklab olish.
   - **Oflayn ishlaydi:** Rendering kutubxonalari (`docx-preview.min.js`, `jszip.min.js`) to'liq mahalliy `static/js/` jildida saqlangan.

2. **Dinamik Filtrlash va Skanerlash:**
   - **URL orqali:** Umumiy ro'yxat (`https://vacancy.argos.uz/hrm-vacancy-list`) yoki bitta vakansiya sahifasi (`.../detail/12345`).
   - **Viloyat va Tuman:** O'zbekistonning barcha 14 ta hududi va tumanlar bo'yicha dinamik kaskad filtr.
   - **Test turi:** 
     - ⭐ *Davlat fuqarolik xizmatchisi (Boshqaruv xodimi)*
     - *Davlat fuqarolik xizmatchisi (Mutaxassis)*
     - *Hamshiralik ishi*
     - *Shifokorlar uchun*
     - *Test talab etilmaydi*
   - **Jonli Progress (SSE):** Skanerlash jarayoni real vaqtda progress bar va topilgan vakansiyalar kartalari ko'rinishida oqim bo'lib keladi.

3. **Hujjatlar Tarixi (History):**
   - Avval yaratilgan barcha Word hujjatlari `exports/` jildida saqlanadi.
   - O'ng burchakdagi **"Hujjatlar tarixi"** tugmasi orqali istalgan avvalgi Word faylini bir zumda vebda qayta ochish va ko'rish mumkin.

4. **Kompyuterdan Word ochish:**
   - "Fayl ochish" tugmasi orqali kompyuterdagi istalgan `.docx` faylini ham brauzer ichida ko'rish imkoniyati mavjud.

5. **Zamonaviy Dizayn va Qulaylik:**
   - Kunduzgi (Light) va Tungi (Dark) rejimlar.
   - 1-bosishda ishga tushirish (`ishga_tushirish_web.bat`).

---

## 🚀 Ishga Tushirish

### 1-usul (Eng osoni):
Fayl menejerida `C:\NBRC Com\argos_web\ishga_tushirish_web.bat` fayliga sichqoncha bilan ikki marta bosing.
Server avtomatik ishga tushadi va brauzeringizda quyidagi manzil ochiladi:
```
http://127.0.0.1:5000
```

### 2-usul (Terminal orqali):
```bash
cd "C:\NBRC Com\argos_web"
pip install -r requirements.txt
python app.py
```

---

## 📁 Jildlar Tuzilishi

```
C:\NBRC Com\argos_web\
│
├── app.py                     # Flask asosiy web server va Argos API tahlilchi
├── ishga_tushirish_web.bat    # Windows uchun 1-bosishda ishga tushiruvchi fayl
├── requirements.txt           # Kerakli Python kutubxonalari
├── firebase.json              # Firebase Hosting sozlamalari (argostest.web.app)
├── .firebaserc                # Firebase loyiha bog'lanishi
├── README.md                  # Qo'llanma va hujjatlar
│
├── templates/
│   └── index.html             # Argos 2-ustunli zamonaviy Web UI va Word Viewer
│
├── static/
│   ├── css/
│   │   └── style.css          # Maxsus uslublar va A4 sahifalar soyasi
│   └── js/
│       ├── docx-preview.min.js # Mahalliy brauzer ichida Word chizuvchi
│       └── jszip.min.js       # Mahalliy arxiv ochuvchi
│
└── exports/                   # Yaratilgan barcha Word (.docx) hujjatlari saqlanadigan joy
```

---

## ☁️ Firebase Hosting ga Deploy Qilish (`http://argostest.web.app/`)

Loyiha Google Cloud Run va Firebase Hosting uchun to'liq tayyorlangan.

### 1-usul: Avtomatlashtirilgan Batch orqali (Eng oson)
Papka ichidagi `deploy_firebase.bat` faylini ishga tushiring:
1. `[1]` ni tanlab Firebase/Google akkauntingizga kiring.
2. `[4]` ni tanlab bir martada Cloud Run va `argostest.web.app` ga to'liq yuklang.

### 2-usul: Buyruqlar qatori orqali
```bash
# 1. Google Cloud Run ga konteynerni yuklash
gcloud run deploy argos-web --project=argostest --source . --region=us-central1 --allow-unauthenticated

# 2. Firebase Hosting ga deploy qilish
npx firebase-tools deploy --only hosting --project argostest
```

Saytingiz `https://argostest.web.app` manzilida avtomatik ishga tushadi!

---

## ⚙️ Sozlamalar va Port
- **Lokalda:** Server standart `http://127.0.0.1:5000` portida ishga tushadi.
- **Bulutda (Cloud Run / Firebase):** Standart `PORT=8080` orqali dinamik ishlaydi.


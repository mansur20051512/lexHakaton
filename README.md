# ⚖️ ADOLAT.AI — O‘zbekiston Respublikasi Huquqiy Xizmatlari Milliy Portali (LegalTech)

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Django 5.x](https://img.shields.io/badge/Django-5.2-green.svg)](https://www.djangoproject.com/)
[![VoiceLab Aisha Orbit](https://img.shields.io/badge/AI-VoiceLab%20Aisha%20Orbit%20755B-orange.svg)](https://voicelab.uz)
[![Lex.uz 2026 Standart](https://img.shields.io/badge/Legislation-Lex.uz%202026%20BHM%20440k-red.svg)](https://lex.uz)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**ADOLAT.AI** — O‘zbekiston Respublikasi qonunchiligi (Lex.uz) va Adliya vazirligi standartlari asosida fuqarolarga sudgacha tinch, tez va qonuniy yechim beruvchi innovatsion sun'iy intellekt (LegalTech) milliy platformasi.

---

## 🌟 Asosiy Imkoniyatlar va Modullar

1. **🏛️ Yuridik Tahlil & Davlat Standartidagi Hujjat Generatori (.docx):**
   - Muammoni oddiy xalq tilida tahlil qilib, amaldagi kodekslar (FK, MK, MJtK, SK, OK) bo'yicha moddama-modda tushuntiradi.
   - **2026-yilgi eng yangi qonunchilik:** 1 BHM = **440 000 so'm** va MJtK 332-1-moddasi bo'yicha ma'muriy jarimalarni **15 kun ichida to'laganda 50% chegirma** imtiyozi avtomatik hisoblanadi.
   - O'zDSt / GOST davlat ish yuritish talablariga 100% mos rasmiy Word (.docx) ariza, da'vo yoki da'voni tan olmaslik hujjati yaratadi.

2. **⚖️ Mediasiya va Foyda Kalkulyatori:**
   - Sud jarayoni (2–6 oy, 4% davlat boji, advokat xarajatlari) va Mediasiya (3–7 kun, 80%+ ijobiy kelishuv, tejamkorlik) xarajatlari hamda muddatlarini taqqoslab beradi.

3. **🏢 Davlat Idoralariga Murojaat Konstruktori (O‘RQ-445):**
   - «Jismoniy va yuridik shaxslarning murojaatlari to‘g‘risida»gi Qonun asosida portal, vazirlik, hokimiyat va prokuraturaga rasmiy murojaat loyihalarini shakllantiradi.

4. **📑 Notariat Gidi va Bojlar Portali (My.gov.uz & E-notarius):**
   - Meros, oldi-sotdi, hadya va ishonchnomalar bo'yicha talab qilinadigan hujjatlar ro'yxati va davlat boji stavkalarini aniq hisoblab beradi.

5. **🌐 100% Ko‘p tillilik (Trilingual i18n):**
   - **O'zbekcha (Lotin)**, **Ўзбекча (Кирилл)** va **Русский** tillari o'rtasida bir klikda to'liq sinxron almashadi. Hech qanday kalit tarjimasiz qolib ketmaydi.

6. **🎙️ Multimodal Kiritish (Universal Hujjat O'quvchi & Ovoz):**
   - Fuqarolar yuridik muammolarini klaviaturada yozishi, **Web Speech API** orqali ovozli aytishi yoki **PDF**, **Rasm (JPG/PNG)**, **DOCX**, **TXT** hujjatlarini yuklab, tizimga tahlil qildirishi mumkin.

---

## 🧠 Sun'iy Intellekt va Neyrotizim Arxitekturasi

Platforma mustaqil va milliy AI tizimlari bilan integratsiyalashgan:
- **VoiceLab Aisha Orbit (755B) & Aisha Halo Vision:** O'zbek tili grammatikasi, huquqiy terminologiyasi va Lex.uz bazasi bo'yicha chuqur o'qitilgan milliy generativ model.
- **Qwen 2.5 / OpenAI-compatible REST Client:** Sof Python `requests` orqali ishlaydi (tashqi og'ir SDK'larsiz).
- **Oflayn Huquqiy Dvijok:** Internet yoki API bo'lmagan taqdirda ham deterministik Lex.uz bilimlar bazasi orqali to'liq ishlaydi.

---

## 🔒 Maxfiylik va Xavfsizlik

> [!IMPORTANT]
> Ushbu ochiq kodli repozitoriyda hech qanday shaxsiy API kalitlar yoki foydalanuvchi ma'lumotlari saqlanmaydi. Barcha maxfiy kalitlar `.env` fayli orqali boshqariladi va `.gitignore` orqali himoyalangan.

---

## 🚀 O'rnatish va Ishga Tushirish

### 1. Repozitoriyadan nusxa olish:
```bash
git clone https://github.com/mansur20051512/lexHakaton.git
cd lexHakaton
```

### 2. Virtual muhit yaratish va kutubxonalarni o'rnatish:
```bash
python -m venv venv
# Windows uchun:
venv\Scripts\activate
# Linux/macOS uchun:
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Konfiguratsiya faylini tayyorlash:
`.env.example` faylidan nusxa olib, `.env` faylini yarating:
```bash
copy .env.example .env
```
O'zingizning VoiceLab yoki Qwen API kalitingizni `.env` fayliga kiriting:
```env
VOICELAB_API_KEY=your_actual_key_here
VOICELAB_LLM_MODEL=aisha-orbit
```

### 4. Ma'lumotlar bazasi migratsiyasi:
```bash
python manage.py makemigrations legal
python manage.py migrate
```

### 5. Serverni ishga tushirish:
```bash
python manage.py runserver
```

Brauzeringizda oching: **`http://127.0.0.1:8000/`**

---

## 👨‍💻 Muallif

- **Mansur Akboyev** — [GitHub Profili](https://github.com/mansur20051512)
- TATU (Toshkent Axborot Texnologiyalari Universiteti)
- Loyiha: O'zbekiston LegalTech Milliy Xakatoni uchun maxsus ishlab chiqilgan.

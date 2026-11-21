<div align="center">

<!-- Header Banner -->
<img src="https://capsule-render.vercel.app/api?type=waving&color=gradient&customColorList=10,24,42,100&height=220&section=header&text=ADOLAT.AI%20%E2%9A%96%EF%B8%8F&fontSize=42&fontColor=ffffff&animation=fadeIn&subtext=O%E2%80%98zbekiston%20Respublikasi%20Huquqiy%20Xizmatlari%20Milliy%20Portali%20(LegalTech)&subfontSize=16&subfontColor=d1d5db" width="100%" />

<p align="center">
  <a href="https://github.com/mansur20051512/lexHakaton">
    <img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.10+" />
  </a>
  <a href="https://github.com/mansur20051512/lexHakaton">
    <img src="https://img.shields.io/badge/Django-5.2-092E20.svg?style=for-the-badge&logo=django&logoColor=white" alt="Django 5.2" />
  </a>
  <a href="https://voicelab.uz">
    <img src="https://img.shields.io/badge/AI-VoiceLab%20Aisha%20Orbit%20755B-FF6F00.svg?style=for-the-badge&logo=openai&logoColor=white" alt="VoiceLab Aisha" />
  </a>
  <a href="https://lex.uz">
    <img src="https://img.shields.io/badge/Lex.uz%202026-BHM%20440%2C000%20so%27m-critical.svg?style=for-the-badge&logo=googledocs&logoColor=white" alt="Lex.uz 2026" />
  </a>
  <a href="https://github.com/mansur20051512/lexHakaton/blob/main/LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge" alt="MIT License" />
  </a>
</p>

<p align="center">
  <b>⚖️ O‘zbekiston Respublikasi Adliya vazirligi standartlari va amaldagi Lex.uz qonunchiligi asosida fuqarolarga sudgacha tinch, tez, arzon va qonuniy huquqiy yordam ko‘rsatuvchi milliy sun'iy intellekt (LegalTech) tizimi.</b>
</p>

</div>

---

## 🌟 Asosiy Imkoniyatlar va Modullar

| Modul | Funktsional Tavsifi | Asosiy Qonunchilik / Standart |
|:---|:---|:---|
| 🏛️ **Yuridik Tahlil & Hujjat Generatori** | Fuqaroning muammosini tahlil qilib, amaldagi kodekslar (FK, MK, MJtK, SK, OK) bo'yicha batafsil huquqiy xulosa va davlat ish yuritish talabidagi Word (.docx) tayyorlaydi. | O‘zDSt / GOST standartlari, Fuqarolik va Mehnat kodekslari |
| ⚖️ **Mediasiya & Foyda Kalkulyatori** | Sud (2–6 oy, 4% davlat boji, advokat xarajatlari) va Mediasiya (3–7 kun, 80%+ ijobiy kelishuv, 0 boj) taqqoslamasini beradi. | O‘zR «Mediasiya to‘g‘risida»gi Qonuni |
| 🏢 **Murojaat Konstruktori** | Vazirliklar, hokimiyatlar, kommunal idoralar va Xalq qabulxonasiga rasmiy ariza va shikoyat loyihalarini shakllantiradi. | O‘zR «Murojaatlar to‘g‘risida»gi Qonuni (O‘RQ-445) |
| 📑 **Notariat Gidi va Bojlar** | Meros, oldi-sotdi, hadya, ishonchnomalar bo'yicha talab etiladigan hujjatlar va davlat boji stavkalarini aniq hisoblab beradi. | «Davlat boji to‘g‘risida»gi Qonun, My.gov.uz & E-notarius |
| 🎙️ **Multimodal Qabul (Vision & Ovoz)** | PDF hujjatlar, fotosuratlar (JPG/PNG), Word fayllar hamda real vaqtli ovozli xabarlarni tahlil qilish imkoniyati. | Web Speech API, PyMuPDF, EasyOCR |
| 🌐 **100% Ko‘p Tillilik (Trilingual i18n)** | O‘zbekcha (Lotin), Ўзбекча (Кирилл) va Русский tillari o'rtasida to'liq sinxron va mukammal moslashuv. | 0 ta tarjimasiz qolgan matn |

---

## 📅 2026-yilgi Qonunchilik Standartlari

Platforma O'zbekiston Respublikasining 2026-yilgi eng so'nggi normativ-huquqiy hujjatlari asosida ishlaydi:
- **1 BHM (Bazaviy hisoblash miqdori):** **440 000 so‘m** (barcha bojlar, jarimalar va to'lovlar ushbu ko'rsatkichda hisoblanadi).
- **15 kunlik 50% imtiyozli chegirma:** MJtK 332-1-moddasiga binoan barcha ma'muriy jarimalar (masalan, 440 000 so'mlik jarima 15 kun ichida 220 000 so'm) avtomatik hisoblanib, fuqaroga tejamkor yo'l ko'rsatiladi.
- **Yangi Mehnat Kodeksi:** 253 va 269-moddalar bo'yicha kechiktirilgan oylik maoshlar uchun kompensatsiya (MB qayta moliyalash stavkasi) hisobi.

---

## 🏗️ Tizim Arxitekturasi

```mermaid
flowchart TD
    User([👤 Fuqaro]) --> Input{Kiritish Shakli}
    
    Input -->|Klaviaturada yozish| TextIn[Oddiy xalq tilidagi matn]
    Input -->|Ovozli aytish| VoiceIn[Web Speech API / Whisper STT]
    Input -->|Hujjat yuklash| DocIn[Universal Document Reader: PDF, Image, Word]
    
    TextIn --> LegalPipeline[⚡ Legal Processing Pipeline]
    VoiceIn --> LegalPipeline
    DocIn --> LegalPipeline
    
    LegalPipeline --> AI_Engine{🧠 AI Dvijogi}
    AI_Engine -->|Milliy LLM & Vision| VoiceLab[VoiceLab Aisha Orbit 755B]
    AI_Engine -->|Open-Source REST| Qwen[Qwen 2.5 72B / Ollama]
    AI_Engine -->|Deterministik Fallback| RuleEngine[Lex.uz 2026 Knowledge Base]
    
    VoiceLab --> OutputGen[Hujjat va Tahlil Konstruktori]
    Qwen --> OutputGen
    RuleEngine --> OutputGen
    
    OutputGen --> AnalysisUI[🖥️ Interaktiv Huquqiy Tahlil]
    OutputGen --> DocxExport[📄 O‘zDSt Davlat Standartidagi Word .docx]
```

---

## 🔒 Xavfsizlik va Maxfiylik Kafolati

> [!IMPORTANT]
> - Ushbu repozitoriyda hech qanday shaxsiy API kalitlar yoki xavfsizlik tokenlari saqlanmaydi.
> - Barcha API kalitlar faqat lokal `.env` faylida saqlanadi va `.gitignore` orqali to'liq himoyalangan.
> - Fuqarolarning shaxsiy murojaat ma'lumotlari serverda uchinchi shaxslarga uzatilmaydi.

---

## 🚀 O'rnatish va Ishga Tushirish

### 1. Repozitoriydan nusxa olish:
```bash
git clone https://github.com/mansur20051512/lexHakaton.git
cd lexHakaton
```

### 2. Virtual muhit yaratish va bog'liqliklarni o'rnatish:
```bash
python -m venv venv

# Windows:
venv\Scripts\activate

# Linux / MacOS:
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Konfiguratsiya faylini sozlash:
`.env.example` dan nusxa olib, `.env` faylini hosil qiling va o'z kalitingizni kiriting:
```bash
copy .env.example .env
```

```env
SECRET_KEY=your_django_secret_key
DEBUG=True
VOICE_MODE=voicelab
VOICELAB_API_KEY=your_voicelab_key_here
VOICELAB_LLM_MODEL=aisha-orbit
```

### 4. Ma'lumotlar bazasini yangilash:
```bash
python manage.py makemigrations legal
python manage.py migrate
```

### 5. Serverni ishga tushirish:
```bash
python manage.py runserver
```

Brauzerda oching: **`http://127.0.0.1:8000/`**

---

## 👨‍💻 Muallif

<div align="center">

**Mansur Akboyev** — [GitHub Profili](https://github.com/mansur20051512)  
🏛️ Toshkent Axborot Texnologiyalari Universiteti (TATU / TUIT)  
🏆 *O‘zbekiston Respublikasi Milliy LegalTech Xakatoni uchun maxsus ishlab chiqilgan.*

<br/>

<!-- Footer Waving Banner -->
<img src="https://capsule-render.vercel.app/api?type=waving&color=gradient&customColorList=10,24,42,100&height=100&section=footer" width="100%" />

</div>

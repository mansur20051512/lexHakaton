import os
import re
import json
import time
import io
try:
    import json_repair
except ImportError:
    json_repair = None
from pathlib import Path
from typing import Dict, Any, Tuple
import requests

from django.conf import settings
from docx import Document
from docx.shared import Pt, Inches, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from .prompts import SYSTEM_PROMPT, build_llm_messages


class JSONSanitizer:
    """
    Qwen va boshqa Open-Source LLM-lar uchun kuchli Regex va JSON tozalovchi.
    Hech qachon JSONDecodeError yuzaga kelishiga yo‘l qo‘ymaydi.
    """

    @staticmethod
    def clean_and_parse(raw_text: str) -> Dict[str, Any]:
        if not raw_text or not isinstance(raw_text, str):
            raise ValueError("Kiritilgan matn bo‘sh yoki noto‘g‘ri formatda")

        text = raw_text.strip()
        text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)
        text = re.sub(r'```(?:json)?\s*', '', text)
        text = re.sub(r'```', '', text)

        first_brace = text.find('{')
        last_brace = text.rfind('}')

        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            json_str = text[first_brace:last_brace + 1]
        else:
            json_str = text

        json_str = re.sub(r',\s*([}\]])', r'\1', json_str)

        try:
            return json.loads(json_str)
        except Exception:
            pass

        if json_repair:
            try:
                repaired = json_repair.loads(json_str)
                if isinstance(repaired, dict) and 'executive_summary' in repaired:
                    return repaired
            except Exception:
                pass

        try:
            escaped_str = re.sub(r'(?<!\\)\n', r'\\n', json_str)
            escaped_str = re.sub(r'(?<!\\)\r', '', escaped_str)
            escaped_str = re.sub(r'(?<!\\)\t', r'\\t', escaped_str)
            return json.loads(escaped_str)
        except Exception:
            pass

        if json_repair:
            try:
                repaired = json_repair.loads(text)
                if isinstance(repaired, dict):
                    return repaired
            except Exception:
                pass

        return JSONSanitizer._regex_fallback_extract(text)

    @staticmethod
    def _regex_fallback_extract(text: str) -> Dict[str, Any]:
        headline_match = re.search(r'"headline"\s*:\s*"([^"]+)"', text)
        headline = headline_match.group(1) if headline_match else "Yuridik tahlil yakunlandi"

        advice_match = re.search(r'"detailed_advice"\s*:\s*"([^"]+)"', text)
        detailed_advice = advice_match.group(1) if advice_match else "Mazkur masala bo‘yicha sudgacha mediasiya tartibida kelishuvga erishish tavsiya etiladi."

        doc_title_match = re.search(r'"document_title"\s*:\s*"([^"]+)"', text)
        doc_title = doc_title_match.group(1) if doc_title_match else "TALABNOMA"

        doc_body_match = re.search(r'"document_body"\s*:\s*"([^"]+)"', text)
        doc_body = doc_body_match.group(1) if doc_body_match else "O‘zbekiston Respublikasi qonunchiligi asosida rasmiy murojaat."

        return {
            "meta": {
                "detected_language": "uz_latn",
                "primary_intent": "mediation",
                "secondary_intent": "appeal",
                "urgency_level": "o‘rtacha",
                "case_category": "Fuqarolik-huquqiy nizo"
            },
            "executive_summary": {
                "headline": headline,
                "detailed_advice": detailed_advice,
                "legal_articles": [
                    {
                        "code_name": "O‘zR Fuqarolik kodeksi",
                        "article": "324-modda. Majburiyatlarni buzish oqibatlari",
                        "relevance": "Qarzdor majburiyatni lozim darajada bajarmaganda yetkazilgan zararlarni to‘lashini kafolatlaydi."
                    },
                    {
                        "code_name": "O‘zR «Mediasiya to‘g‘risida»gi Qonuni",
                        "article": "15-modda. Mediasiya kelishuvi",
                        "relevance": "Nizoni ixtiyoriy ravishda o‘zaro manfaatli shartlarda sudgacha hal etish huquqini beradi."
                    }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True,
                    "court_duration": "2 – 4 oy",
                    "court_cost_estimate": "Da'vo summasining 4% davlat boji + advokat",
                    "mediation_duration": "3 – 7 kun",
                    "mediation_cost_estimate": "Minimal / Kelishuv asosida",
                    "peaceful_resolution_chance": 85,
                    "why_mediation_better": "Mediasiya ortiqcha sud xarajatlari va asab buzilishisiz nizoni tinch hal qiladi."
                },
                "warnings": [
                    "Barcha mavjud yozma dalillar, to‘lov kvitansiyalari va yozishmalarni ehtiyot qiling."
                ]
            },
            "action_roadmap": [
                {
                    "step_number": 1,
                    "title": "Rasmiy talabnoma yuborish",
                    "description": "Tayyorlangan hujjatni ikkinchi tarafga topshiring.",
                    "deadline_hint": "Bugun / 2 kun ichida"
                },
                {
                    "step_number": 2,
                    "title": "Muzokara va Mediasiya",
                    "description": "Professional mediator jalb etib tinch kelishuv imzolang.",
                    "deadline_hint": "7 kun ichida"
                }
            ],
            "notary_module": {
                "is_relevant": False,
                "service_title": "Notarial harakatlar",
                "required_documents": [],
                "estimated_state_duty": "Belgilanmagan",
                "pro_tip": "Elektron notarius tizimidan foydalaning."
            },
            "official_document": {
                "is_generated": True,
                "document_type": "TALABNOMA",
                "recipient_title": "[Javobgar shaxs yoki tashkilotga]",
                "sender_placeholder": "[F.I.SH.], [Manzil], [Telefon]",
                "document_title": doc_title,
                "document_body": doc_body.replace('\\n', '\n'),
                "attachments_list": [
                    "Asoslovchi hujjatlar nusxasi",
                    "Shaxsni tasdiqlovchi hujjat nusxasi"
                ]
            }
        }


class VoiceLabClient:
    """
    VoiceLab Aisha Orbit (va boshqa Aisha modellari) uchun rasmiy yuqori unumli REST mijozi.
    https://api.voicelab.uz/v1/chat/completions
    """

    def __init__(self, override_key=None, override_url=None, override_model=None):
        self.base_url = override_url or getattr(settings, 'VOICELAB_BASE_URL', 'https://api.voicelab.uz')
        self.api_url = f"{self.base_url.rstrip('/')}/v1/chat/completions"
        self.api_key = override_key if override_key is not None else getattr(settings, 'VOICELAB_API_KEY', '')
        self.model_name = override_model or getattr(settings, 'VOICELAB_LLM_MODEL', 'aisha-orbit')
        self.timeout = getattr(settings, 'LLM_TIMEOUT', 60)

    def is_available(self) -> bool:
        return bool(self.api_key)

    def generate(self, messages: list) -> str:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        payload = {
            "model": self.model_name,
            "messages": messages,
            "reasoning_effort": "low",
            "temperature": 0.1,
            "max_tokens": 8192
        }

        response = requests.post(
            self.api_url,
            headers=headers,
            json=payload,
            timeout=self.timeout
        )

        response.raise_for_status()
        result_json = response.json()
        choice = result_json['choices'][0]
        msg = choice.get('message', {})
        content = msg.get('content')
        if not content and 'reasoning_content' in msg:
            content = msg['reasoning_content']
        return content or ""


class QwenRESTClient:
    """
    OpenAI yoki Gemini SDK ishlatmasdan, sof requests kutubxonasi orqali
    Qwen REST API (Ollama localhost yoki OpenRouter/vLLM) bilan ishlovchi klient.
    """

    def __init__(self, override_key=None, override_url=None):
        self.api_url = override_url or getattr(settings, 'LLM_API_URL', 'http://localhost:11434/v1/chat/completions')
        self.api_key = override_key if override_key is not None else getattr(settings, 'LLM_API_KEY', '')
        self.model_name = getattr(settings, 'LLM_MODEL_NAME', 'qwen2.5:7b')
        self.timeout = getattr(settings, 'LLM_TIMEOUT', 45)

    def is_available(self) -> bool:
        # Agar lokal Ollama bo'lsa yoki API kalit berilgan bo'lsa
        if "localhost" in self.api_url or "127.0.0.1" in self.api_url:
            return True
        return bool(self.api_key)

    def generate(self, messages: list) -> str:
        headers = { "Content-Type": "application/json" }
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        payload = {
            "model": self.model_name,
            "messages": messages,
            "temperature": 0.2,
            "top_p": 0.9,
            "max_tokens": 3000
        }

        response = requests.post(
            self.api_url,
            headers=headers,
            json=payload,
            timeout=self.timeout
        )

        response.raise_for_status()
        result_json = response.json()
        return result_json['choices'][0]['message']['content']


class DynamicLegalBrain:
    """
    KUCHLI VA DINAMIK HUQUQIY TAHLIL DVIJOGI:
    Har qanday yangi kiritilgan erkin savolni (mehnat, oila, aliment, iste'molchi, ijara,
    avariya, shartnoma, jarima, kommunal, meros, qarz) tahlil qilib,
    aynan fuqaro kiritgan vaziyat bo'yicha maxsus yuridik xulosa, tegishli qonun moddalari
    va 100% mos rasmiy hujjat yaratadi!
    """

    @staticmethod
    def detect_lang(text: str, forced_lang: str = None) -> str:
        if forced_lang in ['ru', 'uz_cyrl', 'uz_latn']:
            return forced_lang

        cyrillic_chars = len(re.findall(r'[\u0400-\u04FF]', text))
        total_letters = len(re.findall(r'[a-zA-Z\u0400-\u04FF]', text))

        if total_letters > 0 and (cyrillic_chars / total_letters) > 0.35:
            uz_cyrl_markers = ['қ', 'ғ', 'ҳ', 'ў', 'бўлиб', 'керак', 'билан', 'қилинган', 'берган', 'учун', 'қилиш']
            text_lower = text.lower()
            if any(m in text_lower for m in uz_cyrl_markers):
                return 'uz_cyrl'
            return 'ru'
        return 'uz_latn'

    @classmethod
    def get_sanctions_and_liability(cls, category_key: str, lang: str) -> Dict[str, str]:
        """
        O'zbekiston Respublikasi qonunchiligi (Lex.uz / uzbek-legal-corpus) asosidagi
        aniq huquqbuzarlik alomatlari, qonuniy jazolar va kompensatsiya huquqlari.
        """
        if category_key == 'labor':
            if lang == 'ru':
                return {
                    "infringement_elements": "Нарушение работодателем статей 253 и 269 ТК РУз: задержка выплаты заработной платы или неправомерное расторжение трудового договора.",
                    "applicable_sanctions": "Штраф на должностное лицо от 5 до 10 БРВ по ст. 49 КоАО РУз (до 15 БРВ при повторном нарушении). При умышленном незаконном увольнении — уголовная ответственность по ст. 148 УК РУз.",
                    "compensation_rights": "Полное взыскание среднего заработка за все дни вынужденного прогула и денежная компенсация морального вреда (ст. 563, 564 ТК РУз)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Иш берувчи томонидан ЎзР Меҳнат кодексининг 253 ва 269-моддалари талаблари бузилиб, ойлик маош асоссиз кечиктирилган ёки ноқонуний ишдан бўшатиш содир этилган.",
                    "applicable_sanctions": "ЎзР МЖтК 49-моддаси: мансабдор шахсга БҲМнинг 5 бараваридан 10 бараваригача жарима (қайта содир этилса — 15 бараваргача). Қасддан ишдан бўшатиш ҳолатида — ЖК 148-моддаси билан жиноий жавобгарлик хавфи.",
                    "compensation_rights": "МК 563 ва 564-моддалари: барча мажбурий прогул кунлари учун тўлиқ ўртача иш ҳақи ва маънавий зарар учун моддий компенсация тўлиқ ундирилади."
                }
            else:
                return {
                    "infringement_elements": "Ish beruvchi tomonidan O'zR Mehnat kodeksining 253 va 269-moddalari talablari buzilib, oylik maosh asossiz kechiktirilgan yoki noqonuniy ishdan bo'shatish sodir etilgan.",
                    "applicable_sanctions": "O'zR MJtK 49-moddasi: mansabdor shaxsga BHMning 5 baravaridan 10 baravarigacha jarima (qayta sodir etilsa — 15 baravargacha). Qasddan ishdan bo'shatish holatida — JK 148-moddasi bilan jinoiy javobgarlik xavfi.",
                    "compensation_rights": "MK 563 va 564-moddalari: barcha majburiy progul kunlari uchun to'liq o'rtacha ish haqi va ma'naviy zarar uchun moddiy kompensatsiya to'liq undiriladi."
                }

        elif category_key == 'family':
            if lang == 'ru':
                return {
                    "infringement_elements": "Уклонение от содержания несовершеннолетних детей в нарушение статьи 96 Семейного кодекса РУз.",
                    "applicable_sanctions": "Статья 47-4 КоАО РУз: административный арест до 15 суток или штраф до 20 БРВ. При задолженности более 2 месяцев — уголовная ответственность по ст. 122 УК РУз (лишение свободы до 1 года).",
                    "compensation_rights": "Взыскание алиментов в твердой доле (до 1/2 на 3+ детей) и неустойки в размере 0.1% от суммы долга за каждый день просрочки (ст. 140 СК РУз)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Оила кодексининг 96-моддасига зид равишда вояга етмаган фарзандлар таъминоти учун алимент тўлашдан қасддан бўйин товлаш.",
                    "applicable_sanctions": "ЎзР МЖтК 47-4-моддаси: 15 суткагача маъмурий қамоқ ёки БҲМнинг 20 бараваригача жарима. 2 ойдан ортиқ тўламаслик ҳолатида: ЖК 122-моддаси бўйича 2 йилгача ахлоқ тузатиш ёки 1 йилгача озодликдан маҳрум қилиш жазоси.",
                    "compensation_rights": "Оила кодекси 99 ва 140-моддалари: алимент қарздорлигининг ҳар бир кечиктирилган куни учун 0.1% миқдорида пеня ундириш ҳуқуқи."
                }
            else:
                return {
                    "infringement_elements": "Oila kodeksining 96-moddasiga zid ravishda voyaga yetmagan farzandlar ta'minoti uchun aliment to'lashdan qasddan bo'yin tovlash.",
                    "applicable_sanctions": "O'zR MJtK 47-4-moddasi: 15 sutkagacha ma'muriy qamoq yoki BHMning 20 baravarigacha jarima. 2 oydan ortiq to'lamaslik holatida: JK 122-moddasi bo'yicha 2 yilgacha axloq tuzatish yoki 1 yilgacha ozodlikdan mahrum qilish jazosi.",
                    "compensation_rights": "Oila kodeksi 99 va 140-moddalari: aliment qarzdorligining har bir kechiktirilgan kuni uchun 0.1% miqdorida penya undirish huquqi."
                }

        elif category_key == 'consumer':
            if lang == 'ru':
                return {
                    "infringement_elements": "Отказ продавца в замене дефектного товара или возврате денежных средств в нарушение ст. 13, 14 Закона «О защите прав потребителей».",
                    "applicable_sanctions": "Штраф по ст. 178 КоАО РУз от 5 до 10 БРВ на должностных лиц и наложение финансовых санкций уполномоченным органом по защите прав потребителей.",
                    "compensation_rights": "Взыскание неустойки в размере 1% от стоимости товара за каждый день просрочки удовлетворения требований и компенсация морального вреда (ст. 18, 22 Закона)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "«Истеъмолчилар ҳуқуқларини ҳимоя қилиш тўғрисида»ги Қонуннинг 13, 14-моддалари бузилиб, нуқсонли товар учун маблағни қайтаришдан ёки алмаштиришдан бош тортиш.",
                    "applicable_sanctions": "ЎзР МЖтК 178-моддаси: истеъмолчилар ҳуқуқларини бузганлик учун БҲМнинг 5 бараваридан 10 бараваригача жарима ва Рақобат қўмитаси томонидан корхонага нисбатан молиявий санкция қўллаш.",
                    "compensation_rights": "Қонуннинг 18 ва 22-моддалари: ҳар бир кечиктирилган кун учун 1% пеня ва истеъмолчига етказилган маънавий зарарни пул шаклида ундириш."
                }
            else:
                return {
                    "infringement_elements": "«Iste'molchilar huquqlarini himoya qilish to'g'risida»gi Qonunning 13, 14-moddalari buzilib, nuqsonli tovar uchun mablag'ni qaytarishdan yoki almashtirishdan bosh tortish.",
                    "applicable_sanctions": "O'zR MJtK 178-moddasi: iste'molchilar huquqlarini buzganlik uchun BHMning 5 baravaridan 10 baravarigacha jarima va Raqobat qo'mitasi tomonidan korxonaga nisbatan moliyaviy sanksiya qo'llash.",
                    "compensation_rights": "Qonunning 18 va 22-moddalari: har bir kechiktirilgan kun uchun 1% penya va iste'molchiga yetkazilgan ma'naviy zararni pul shaklida undirish."
                }

        elif category_key == 'debt':
            if lang == 'ru':
                return {
                    "infringement_elements": "Невозврат суммы займа в срок по расписке или договору в нарушение статей 732, 735 Гражданского кодекса РУз.",
                    "applicable_sanctions": "Статья 327 ГК РУз: начисление законных процентов по ставке рефинансирования ЦБ РУз. При наличии признаков умышленного завладения имуществом — ст. 168 УК РУз (Мошенничество).",
                    "compensation_rights": "Полное взыскание основного долга, процентов, упущенной выгоды и возмещение всех судебных и почтовых расходов (ст. 14 ГК РУз)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "ЎзР Фуқаролик кодексининг 732, 735-моддалари бузилиб, тилхат ёки шартнома бўйича олинган қарз суммасини муддатида қайтармаслик.",
                    "applicable_sanctions": "ФК 327-моддаси: ЎзР Марказий банки қайта молиялаш ставкаси бўйича фоизлар. Агар қарз дастлаб алдов йўли билан олинган бўлса — ЖК 168-моддаси (Фирибгарлик — 3 йилдан 5 йилгача озодликдан маҳрум қилиш) хавфи.",
                    "compensation_rights": "ФК 14-моддаси: асосий қарз, ҳисобланган банк фоизлари, бой берилган фойда ва барча суд-почта харажатларини тўлиқ ундириб олиш ҳуқуқи."
                }
            else:
                return {
                    "infringement_elements": "O'zR Fuqarolik kodeksining 732, 735-moddalari buzilib, tilxat yoki shartnoma bo'yicha olingan qarz summasini muddatida qaytarmaslik.",
                    "applicable_sanctions": "FK 327-moddasi: O'zR Markaziy banki qayta moliyalash stavkasi bo'yicha foizlar. Agar qarz dastlab aldov yo'li bilan olingan bo'lsa — JK 168-moddasi (Firibgarlik — 3 yildan 5 yilgacha ozodlikdan mahrum qilish) xavfi.",
                    "compensation_rights": "FK 14-moddasi: asosiy qarz, hisoblangan bank foizlari, boy berilgan foyda va barcha sud-pochta xarajatlarini to'liq undirib olish huquqi."
                }

        elif category_key == 'damages':
            if lang == 'ru':
                return {
                    "infringement_elements": "Причинение материального ущерба имуществу гражданина (ДТП / затопление) и уклонение от добровольного возмещения по ст. 985 ГК РУз.",
                    "applicable_sanctions": "Статья 133 КоАО РУз: лишение водительских прав до 1 года или штраф до 4 БРВ; принудительный арест счетов и имущества виновника через БПИ.",
                    "compensation_rights": "Полное возмещение восстановительного ремонта, стоимости независимой оценки и морального вреда (ст. 985, 1021 ГК РУз)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Ўзганинг мулкига (автотранспорт ёки кўчмас мулкка) бевосита моддий зарар етказиш ва ЎзР ФК 985-моддаси бўйича тўлашдан бош тортиш.",
                    "applicable_sanctions": "МЖтК 133-моддаси: йўл ҳаракати хавфсизлигини бузиб моддий зарар етказганлик учун БҲМ 2-4 баравари жарима ёки транспорт бошқариш ҳуқуқидан 1 йилгача маҳрум қилиш; МИБ орқали мол-мулкни хатлаш.",
                    "compensation_rights": "ФК 985 ва 1021-моддалар: эҳтиёт қисмлар, таъмирлаш, мустақил баҳолаш харажатлари ва маънавий зарар қийматини тўлиқ ундириш."
                }
            else:
                return {
                    "infringement_elements": "O'zganing mulkiga (avtotransport yoki ko'chmas mulkka) bevosita moddiy zarar yetkazish va O'zR FK 985-moddasi bo'yicha to'lashdan bosh tortish.",
                    "applicable_sanctions": "O'zR MJtK 133-moddasi: yo'l harakati xavfsizligini buzib moddiy zarar yetkazganlik uchun BHM 2-4 baravari jarima yoki transport boshqarish huquqidan 1 yilgacha mahrum qilish; MIB orqali mol-mulkni xatlash.",
                    "compensation_rights": "FK 985 va 1021-moddalar: ehtiyot qismlar, ta'mirlash, mustaqil baholash xarajatlari va ma'naviy zarar qiymatini to'liq undirish."
                }

        elif category_key == 'rent':
            if lang == 'ru':
                return {
                    "infringement_elements": "Нарушение условий найма жилья, неуплата платежей более 3 месяцев и незаконное удержание помещения (ст. 535 ГК РУз).",
                    "applicable_sanctions": "Статья 134 Жилищного кодекса РУз: принудительное выселение на основании решения суда через судебных исполнителей БПИ.",
                    "compensation_rights": "Взыскание всей суммы задолженности по арендной плате, компенсация коммунальных расходов и пени за нарушение договора."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Ижара шартномаси шартларига зид равишда тўловларни кечиктириш ва турар жойни бўшатмасдан ноқонуний эгаллаб туриш.",
                    "applicable_sanctions": "Уй-жой кодекси 134-моддаси: суд қарори асосида мажбурий кўчириб чиқариш, қарздорнинг банк карталари ва даромадларига тақиқ қўйиш.",
                    "compensation_rights": "Тўланмаган ижара қарздорлиги, коммунал хизмат тўловлари ва ноқонуний яшалган ҳар бир кун учун бозор баҳосидаги ижара ҳақини ундириш."
                }
            else:
                return {
                    "infringement_elements": "Ijara shartnomasi shartlariga zid ravishda to'lovlarni kechiktirish va turar joyni bo'shatmasdan noqonuniy egallab turish.",
                    "applicable_sanctions": "Uy-joy kodeksi 134-moddasi: sud qarori asosida majburiy ko'chirib chiqarish, qarzdorning bank kartalari va daromadlariga taqiq qo'yish.",
                    "compensation_rights": "To'lanmagan ijara qarzdorligi, kommunal xizmat to'lovlari va noqonuniy yashalgan har bir kun uchun bozor bahosidagi ijara haqini undirish."
                }

        elif category_key == 'admin_fine':
            if lang == 'ru':
                return {
                    "infringement_elements": "Вынесение необоснованного постановления о наложении административного штрафа без состава правонарушения (ст. 307 КоАО РУз).",
                    "applicable_sanctions": "Прекращение производства по делу по ст. 271 п. 1 КоАО РУз и отмена постановления о штрафе в судебном порядке (ст. 315 КоАО РУз).",
                    "compensation_rights": "Возврат неправомерно удержанных средств и снятие всех ограничений с транспортного средства."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "ЎзР МЖтК нормаларига зид равишда маъмурий ҳуқуқбузарлик аломати мавжуд бўлмаган ҳолда ноҳақ жарима тўғрисида қарор чиқарилиши.",
                    "applicable_sanctions": "МЖтК 271-моддаси 1-банди: иш юритишни тугатиш ва қарорни бекор қилиш. МЖтК 315-моддаси: суд орқали қарорни ҳақиқий эмас деб топиш.",
                    "compensation_rights": "Тўланган жаримани тўлиқ қайтариб олиш, автомобиль ва ҳисоб рақамлардан барча тақиқларни ечдириш ҳуқуқи."
                }
            else:
                return {
                    "infringement_elements": "O'zR MJtK normalariga zid ravishda ma'muriy huquqbuzarlik alomati mavjud bo'lmagan holda nohaq jarima to'g'risida qaror chiqarilishi.",
                    "applicable_sanctions": "O'zR MJtK 271-moddasi 1-bandi: ish yuritishni tugatish va qarorni bekor qilish. MJtK 315-moddasi: sud orqali qarorni haqiqiy emas deb topish.",
                    "compensation_rights": "To'langan jarimani to'liq qaytarib olish, avtomobil va hisob raqamlardan barcha taqiqlarni yechdirish huquqi."
                }

        elif category_key == 'communal':
            if lang == 'ru':
                return {
                    "infringement_elements": "Незаконное отключение коммунальных услуг (газ, свет, вода) и бездействие должностных лиц по обращениям граждан (ст. 19 Закона ЗРУ-445).",
                    "applicable_sanctions": "Штраф по ст. 43 КоАО РУз на должностных лиц до 3 БРВ и привлечение к служебной дисциплинарной ответственности.",
                    "compensation_rights": "Немедленное восстановление энергоснабжения в течение 24 часов и обязательный перерасчет платы за период перебоев."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "ЎРҚ-445 Қонуни ва коммунал хизмат кўрсатиш қоидаларини бузиб, таъминотни асоссиз узиш ва фуқаро мурожаатларини эътиборсиз қолдириш.",
                    "applicable_sanctions": "МЖтК 43-моддаси: мурожаатларни кўриб чиқиш тартибини бузганлик учун мансабдор шахсга БҲМ 1-3 баравари жарима ва хизмат текшируви тайинлаш.",
                    "compensation_rights": "Таъминотни 24 соат ичида зудлик билан тиклаш ва узилиш даври учун тўловларни қайта ҳисоб-китоб (компенсация) қилдириш ҳуқуқи."
                }
            else:
                return {
                    "infringement_elements": "O'RQ-445 Qonuni va kommunal xizmat ko'rsatish qoidalarini buzib, ta'minotni asossiz uzish va fuqaro murojaatlarini e'tiborsiz qoldirish.",
                    "applicable_sanctions": "O'zR MJtK 43-moddasi: murojaatlarni ko'rib chiqish tartibini buzganlik uchun mansabdor shaxsga BHM 1-3 baravari jarima va xizmat tekshiruvi tayinlash.",
                    "compensation_rights": "Ta'minotni 24 soat ichida zudlik bilan tiklash va uzilish davri uchun to'lovlarni qayta hisob-kitob (kompensatsiya) qildirish huquqi."
                }

        elif category_key == 'notary':
            if lang == 'ru':
                return {
                    "infringement_elements": "Риск пропуска 6-месячного установленного законом срока принятия наследства по статье 1146 ГК РУз.",
                    "applicable_sanctions": "Утрата преимущественного права на наследственное имущество и необходимость судебного восстановления срока.",
                    "compensation_rights": "Оформление свидетельства о праве на наследство с освобождением от лишних пошлин по Закону «О госпошлине»."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Мерос ёки нотариал ҳаракатларни расмийлаштиришда қонуний муддатларни ўтказиб юбориш ёки ҳуқуқларнинг чекланиши хавфи.",
                    "applicable_sanctions": "ЎзР Фуқаролик кодексининг 1146-моддасига кўра мерос мулкидан маҳрум бўлиш ёки суд тартибида муддатни тиклаш мажбурияти юзага келиши.",
                    "compensation_rights": "ЎзР «Давлат божи тўғрисида»ги Қонуни имтиёзларидан фойдаланган ҳолда минимал харажат билан мерос гувоҳномасини олиш кафолати."
                }
            else:
                return {
                    "infringement_elements": "Meros yoki notarial harakatlarni rasmiylashtirishda qonuniy muddatlarni o'tkazib yuborish yoki huquqlarning cheklanishi xavfi.",
                    "applicable_sanctions": "O'zR Fuqarolik kodeksining 1146-moddasiga ko'ra meros mulkidan mahrum bo'lish yoki sud tartibida muddatni tiklash majburiyati yuzaga kelishi.",
                    "compensation_rights": "O'zR «Davlat boji to'g'risida»gi Qonuni imtiyozlaridan foydalangan holda minimal xarajat bilan meros guvohnomasini olish kafolati."
                }

        else: # universal
            if lang == 'ru':
                return {
                    "infringement_elements": "Нарушение договорных обязательств и ущемление законных прав гражданина по статье 333 ГК РУз.",
                    "applicable_sanctions": "Статьи 324, 333 ГК РУз: взыскание неустойки, законных процентов и принудительное исполнение через суд и органы БПИ.",
                    "compensation_rights": "Полное возмещение прямых убытков, неполученных доходов и компенсация морального вреда (ст. 14 ГК РУз)."
                }
            elif lang == 'uz_cyrl':
                return {
                    "infringement_elements": "Ўзбекистон Республикаси қонунчилиги ёки шартномавий мажбуриятларни бажармаслик, ўзганинг қонуний ҳуқуқ ва манфаатларини чеклаш.",
                    "applicable_sanctions": "ФК 324 ва 333-моддалари: мажбуриятни бажармаган тарафга нисбатан неустойка, банк фоизлари ва мажбурий суд ижроси чоралари қўлланилиши.",
                    "compensation_rights": "ФК 14-моддаси асосида барча етказилган зарарлар, маънавий зиён ва харажатларни айбдор шахс ҳисобидан тўлиқ қоплатиш ҳуқуқи."
                }
            else:
                return {
                    "infringement_elements": "O'zbekiston Respublikasi qonunchiligi yoki shartnomaviy majburiyatlarni bajarmaslik, o'zganing qonuniy huquq va manfaatlarini cheklash.",
                    "applicable_sanctions": "O'zR FK 324 va 333-moddalari: majburiyatni bajarmagan tarafga nisbatan neustoyka, bank foizlari va majburiy sud ijrosi choralari qo'llanilishi.",
                    "compensation_rights": "FK 14-moddasi asosida barcha yetkazilgan zararlar, ma'naviy ziyon va xarajatlarni aybdor shaxs hisobidan to'liq qoplatish huquqi."
                }

    @classmethod
    def analyze_custom_query(cls, query: str, forced_lang: str = None) -> Dict[str, Any]:
        lang = cls.detect_lang(query, forced_lang)
        q = query.lower()

        # Toza qisqartirilgan matn (hujjat ichiga qo'yish uchun)
        clean_user_facts = query.strip()
        if len(clean_user_facts) > 400:
            clean_user_facts = clean_user_facts[:400] + "..."

        # 1. NOTARIAT VA MEROS (Meros, vasiyatnoma, hadya, oldi-sotdi, ishonchnoma, notarius)
        if any(w in q for w in ['meros', 'vasiyat', 'hadya', 'notarius', 'ishonchnoma', 'oldi-sotdi', 'наследств', 'нотариус', 'дарствен', 'завещан', 'доверенност', 'мерос']):
            category_key = 'notary'
            res = cls._generate_notary_case(lang, clean_user_facts)

        # 2. MA'MURIY VA JARIMALAR (Radar, jarima, YPX, GAI, soliq qarori)
        elif any(w in q for w in ['radar', 'jarima', 'ypx', 'gai', 'protokol', 'штраф', 'радар', 'протокол', 'гаи', 'жарима']):
            category_key = 'admin_fine'
            res = cls._generate_admin_fine_case(lang, clean_user_facts)

        # 3. MEHNAT NIZOSI (Ishdan bo'shatish, oylik to'lamaslik, kompensatsiya)
        elif any(re.search(rf'(?:\b|_){re.escape(w)}', q) for w in ['ishdan', 'maosh', 'oylik', 'kompensatsiya', 'mehnat', 'ish beruvchi', 'bo\'shat', 'hayda', 'увольн', 'зарплат', 'оклад', 'трудов', 'маош', 'ойлик', 'бўшат']):
            category_key = 'labor'
            res = cls._generate_labor_case(lang, clean_user_facts)

        # 4. OILA VA ALIMENT (Aliment undirish, ajrashish, bola ta'minoti, mulk)
        elif any(w in q for w in ['aliment', 'ajrash', 'xotinim', 'farzand', 'nikoh', 'qaynona', 'алимент', 'развод', 'супруг', 'ребен', 'алиментлар', 'ажраш']) or re.search(r'(?<!y)\berim\b', q):
            category_key = 'family'
            res = cls._generate_family_case(lang, clean_user_facts)

        # 5. ISTE'MOLCHILAR HUQUQLARI (Sifatsiz tovar, telefon buzildi, magazin qaytarmayapti, kafolat)
        elif any(w in q for w in ['magazin', 'dokon', 'tovar', 'mahsulot', 'telefon', 'texnika', 'kafolat', 'brak', 'nuqson', 'almashtir', 'покупк', 'товар', 'магазин', 'брак', 'возврат', 'гаранти', 'дўкон', 'маҳсулот', 'нуқсон']):
            category_key = 'consumer'
            res = cls._generate_consumer_case(lang, clean_user_facts)

        # 6. YETKAZILGAN ZARAR VA YTH (Avtoavariya, mashina urib ketdi, suv bostirdi, zarar)
        elif any(w in q for w in ['avariya', 'urib', 'haydovchi', 'suv bos', 'tamir', 'yetkazilgan zarar', 'straxovka', 'дтп', 'авария', 'затоп', 'ущерб', 'страховк', 'зарар']):
            category_key = 'damages'
            res = cls._generate_damages_case(lang, clean_user_facts)

        # 7. IJARA NIZOSI (Ijara haqi to'lanmasligi, uydan chiqarish, kvartirant)
        elif any(w in q for w in ['ijara', 'ijarachi', 'kvartirant', 'arendator', 'arenda', 'uy egasi', 'chiqarib', 'аренд', 'съемщик', 'квартирант', 'ижара', 'ижарачи']):
            category_key = 'rent'
            res = cls._generate_rent_case(lang, clean_user_facts)

        # 8. DAVLATGA SHIKOYAT VA KOMMUNAL (Elektr, gaz, suv, hokimiyat, obodonlashtirish)
        elif any(w in q for w in ['elektr', 'gaz', 'suv', 'svet', 'hokim', 'hokimiyat', 'kommunal', 'shikoyat', 'хўжалик', 'электр', 'газ', 'жалоб', 'свет', 'жкх', 'хокимият']):
            category_key = 'communal'
            res = cls._generate_communal_case(lang, clean_user_facts)

        # 9. QARZ VA TILXAT (Pul berdim, qaytarmayapti, foizlar)
        elif any(w in q for w in ['qarz', 'tilxat', 'pulim', 'qaytarmayapti', 'bergan edim', 'долг', 'расписк', 'заем', 'қарз', 'тилхат']):
            category_key = 'debt'
            res = cls._generate_debt_case(lang, clean_user_facts)

        # 10. UNIVERSAL ERKIN MUAMMO (Har qanday boshqa fuqarolik yoki shartnoma nizosi)
        else:
            category_key = 'universal'
            res = cls._generate_universal_custom_case(lang, clean_user_facts)

        # Huquqbuzarlik alomatlari va jazolarni kafolatli qo'shish
        res['legal_sanctions_and_liability'] = cls.get_sanctions_and_liability(category_key, lang)
        return res

    # ---------------- 1. MEHNAT NIZOSI ----------------
    @classmethod
    def _generate_labor_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        if lang == 'ru':
            return {
                "meta": { "detected_language": "ru", "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "высокий", "case_category": "Трудовой спор (Незаконное увольнение / Невыплата зарплаты)" },
                "executive_summary": {
                    "headline": "Трудовой кодекс РУз гарантирует обязательную выплату задолженности по зарплате и компенсацию за незаконное увольнение.",
                    "detailed_advice": f"По вашей ситуации («{user_facts[:120]}...»): работодатель грубо нарушает требования Трудового кодекса РУз. Согласно закону, невыплата зарплаты влечет начисление пени и административную ответственность руководителя. Рекомендуется сначала направить официальную Претензию работодателю с требованием мирного расчета в течение 3 дней, а при отказе — обратиться в Государственную инспекцию труда или суд с возложением компенсации вынужденного прогула.",
                    "legal_articles": [
                        { "code_name": "Трудовой кодекс Республики Узбекистан", "article": "Статья 253. Сроки выплаты заработной платы", "relevance": "Обязывает работодателя производить своевременную выплату заработной платы в полном объеме." },
                        { "code_name": "Трудовой кодекс Республики Узбекистан", "article": "Статья 563. Возмещение материального и морального вреда", "relevance": "Предусматривает выплату среднего заработка за все время вынужденного прогула и компенсацию морального вреда." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "1 – 2 месяца", "court_cost_estimate": "Работник полностью освобожден от госпошлины в суде",
                        "mediation_duration": "3 – 5 дней", "mediation_cost_estimate": "Бесплатно", "peaceful_resolution_chance": 88,
                        "why_mediation_better": "Досудебная претензия побуждает работодателя выплатить деньги добровольно, чтобы избежать проверки трудовой инспекции и штрафов."
                    },
                    "warnings": ["Срок обращения в суд по спорам об увольнении составляет 3 месяца со дня вручения приказа (ст. 560 ТК РУз). Не затягивайте!"]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Направление досудебной претензии работодателю", "description": "Вручите подготовленное требование в отдел кадров/канцелярию под роспись.", "deadline_hint": "Сегодня" },
                    { "step_number": 2, "title": "Обращение в Инспекцию труда", "description": "При невыплате подайте электронную жалобу на mehnat.uz.", "deadline_hint": "В течение 3 дней" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Не требуется", "required_documents": [], "estimated_state_duty": "0 сум", "pro_tip": "По трудовым спорам госпошлина не взимается." },
                "official_document": {
                    "is_generated": True, "document_type": "ПРЕТЕНЗИЯ",
                    "recipient_title": "Руководителю организации / Работодателю\nАдрес: [Адрес места работы]",
                    "sender_placeholder": "От работника: [Ф.И.О.]\nДолжность: [Занимаемая должность]\nТел: [Номер телефона]",
                    "document_title": "ПРЕТЕНЗИЯ (ДОСУДЕБНОЕ ТРЕБОВАНИЕ)\nО ВЫПЛАТЕ ЗАДОЛЖЕННОСТИ ПО ЗАРАБОТНОЙ ПЛАТЕ И УРЕГУЛИРОВАНИИ СПОРА",
                    "document_body": f"Я осуществлял(а) трудовую деятельность в вашей организации. В связи со сложившейся ситуацией:\n\n«{user_facts}»\n\nВ соответствии со статьями 253 и 269 Трудового кодекса Республики Узбекистан, работодатель обязан производить оплату труда в установленные сроки, а при прекращении трудового договора — произвести полный расчет в день увольнения.\n\nНарушение данных требований влечет ответственность, установленную статьей 49 КоАО РУз, а также обязанность возмещения процентов за каждый день задержки.\n\nВ целях мирного урегулирования трудового спора без привлечения органов прокуратуры, Государственной трудовой инспекции и суда, ТРЕБУЮ:\n\n1. В течение 3 (трех) банковских дней произвести полный расчет и выплатить причитающуюся сумму заработной платы и компенсаций.\n2. В случае отказа я буду вынужден(а) обратиться в суд с требованием взыскания среднего заработка за вынужденный прогул и морального вреда.",
                    "attachments_list": ["Копия трудового договора или выписка из трудовой книжки", "Банковская выписка (при наличии)"]
                }
            }
        elif lang == 'uz_cyrl':
            return {
                "meta": { "detected_language": "uz_cyrl", "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "юқори", "case_category": "Меҳнат низоси (Ишдан бўшатиш / Ойлик маошни ундириш)" },
                "executive_summary": {
                    "headline": "ЎзР Меҳнат кодекси иш берувчининг асоссиз бўшатиши ва ойликни ушлаб қолишини қатъий таъқиқлайди.",
                    "detailed_advice": f"Сизнинг вазиятингиз бўйича («{user_facts[:120]}...»): Меҳнат кодексининг 253 ва 269-моддаларига кўра, иш берувчи меҳнат шартномаси бекор қилинганда барча ҳисоб-китобларни ўша куннинг ўзида амалга ошириши шарт. Сизга зудлик билан иш берувчига расмий Талабнома топшириш ва агар 3 кунда тўламаса, Давлат меҳнат инспекцияси ҳамда фуқаролик суди орқали мажбурий ундириш тавсия этилади. Меҳнат низолари бўйича ходим давлат божидан тўлиқ озод этилган.",
                    "legal_articles": [
                        { "code_name": "Ўзбекистон Республикаси Меҳнат кодекси", "article": "253-модда. Иш ҳақи тўлаш муддатлари", "relevance": "Иш ҳақини тўлаш муддатларини бузганлик учун иш берувчининг молиявий ва маъмурий жавобгарлигини белгилайди." },
                        { "code_name": "Ўзбекистон Республикаси Меҳнат кодекси", "article": "563-модда. Мажбурий прогул вақти учун ҳақ тўлаш", "relevance": "Ходимни қонунга хилоф равишда бўшатганда барча мажбурий юрган кунлари учун ўртача ойлик тўланишини кафолатлайди." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "1 – 2 ой", "court_cost_estimate": "Ходим судда давлат божидан 100% озод қилинган",
                        "mediation_duration": "3 – 5 кун", "mediation_cost_estimate": "Мутлақо бепул", "peaceful_resolution_chance": 88,
                        "why_mediation_better": "Талабнома олган иш берувчи текширув ва жарималардан қўрқиб, пулни судсиз тўлаб беришни маъқул кўради."
                    },
                    "warnings": ["Ишдан бўшатиш бўйича даъво муддати буйруқ берилган кундан 3 ойни ташкил этади (МК 560-модда). Вақтни ўтказиб юборманг!"]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Иш берувчига расмий талабнома топшириш", "description": "Тайёр талабномани қабулхона ёки кадрлар бўлимига топшириб, нусхасига имзо қўйдириб олинг.", "deadline_hint": "Бугуноқ" },
                    { "step_number": 2, "title": "Давлат меҳнат инспекциясига мурожаат", "description": "Агар 3 кунда пул берилмаса, mehnat.uz ёки 1176 орқали шикоят юборинг.", "deadline_hint": "3 кун ичида" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Зарурият йўқ", "required_documents": [], "estimated_state_duty": "0 сўм", "pro_tip": "Меҳнат низолари учун давлат божи тўланмайди." },
                "official_document": {
                    "is_generated": True, "document_type": "ТАЛАБНОМА",
                    "recipient_title": "«[Корхона/Ташкилот номи]» раҳбарига\nМанзил: [Иш жойи манзили]",
                    "sender_placeholder": "Ходим: [Ф.И.Ш.]дан\nЛавозими: [Эгаллаган лавозими]\nТелефон: [Телефон рақами]",
                    "document_title": "ИШ ҲАҚИ ВА БАРЧА ТЎЛОВЛАРНИ АМАЛГА ОШИРИШ ТЎҒРИСИДА\nРАСМИЙ ТАЛАБНОМА (ПРЕТЕНЗИЯ)",
                    "document_body": f"Мен мазкур ташкилотда ўз меҳнат фаолиятимни амалга ошириб келганман. Бироқ қуйидаги ҳолатлар бўйича меҳнат ҳуқуқларим қўпол равишда бузилди:\n\n«{user_facts}»\n\nЎзбекистон Республикаси Меҳнат кодексининг 253 ва 269-моддаларига мувофиқ, иш берувчи ходим билан меҳнат шартномаси бекор қилинганда ёки ойлик муддати келганда барча ҳақларни тўлиқ тўлаши шарт. Ушбу талабни бузиш МЖтКнинг 49-моддасига асосан жавобгарликка сабаб бўлади.\n\nНизони суд ва Давлат меҳнат инспекцияси аралашувисиз, тинч йўл билан ҳал этиш мақсадида ТАЛАБ ҚИЛАМАН:\n\n1. Мазкур талабнома олинган кундан бошлаб 3 (уч) банк куни ичида тегишли барча иш ҳақи ва компенсация пулларимни тўлаб беришингизни;\n2. Акс ҳолда фуқаролик ишлари бўйича судга даъво киритиб, барча мажбурий прогул ва маънавий зарар пулларини Сиздан ундириб оламан.",
                    "attachments_list": ["Меҳнат шартномаси ёки буйруқ нусхаси", "Пластик карта ҳисоб рақами кўчирмаси"]
                }
            }
        else:
            # uz_latn
            return {
                "meta": { "detected_language": "uz_latn", "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "yuqori", "case_category": "Mehnat nizosi (Ishdan bo'shatish va Ish haqini undirish)" },
                "executive_summary": {
                    "headline": "O'zR Mehnat kodeksi ish beruvchining noqonuniy bo'shatishi va oylikni bermasligini taqiqlaydi hamda to'liq himoya kafolatini beradi.",
                    "detailed_advice": f"Sizning holatingiz bo'yicha («{user_facts[:120]}...»): Mehnat kodeksining 253 va 269-moddalariga muvofiq, ish beruvchi xodim bilan barcha hisob-kitoblarni o'z vaqtida amalga oshirishga majbur. Hozir sudga bormasdan, birinchi navbatda ish beruvchiga 3 kunlik rasmiy Talabnoma topshirish eng to'g'ri yo'ldir. Ish beruvchi Davlat mehnat inspeksiyasi jarimalari va sud xarajatlaridan qochish uchun mablag'ni darhol to'lab berish ehtimoli 88% dan yuqori. Mehnat nizolarida xodim sudga davlat boji to'lamaydi.",
                    "legal_articles": [
                        { "code_name": "O'zbekiston Respublikasi Mehnat kodeksi", "article": "253-modda. Ish haqi to'lash muddatlari", "relevance": "Ish haqini o'z vaqtida to'lamaganlik uchun ish beruvchining qat'iy javobgarligini belgilaydi." },
                        { "code_name": "O'zbekiston Respublikasi Mehnat kodeksi", "article": "563-modda. Majburiy progul vaqti uchun haq to'lash", "relevance": "Noqonuniy bo'shatilgan xodimga barcha o'tgan kunlar uchun o'rtacha ish haqi va ma'naviy zarar to'latilishini kafolatlaydi." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "1 – 2 oy", "court_cost_estimate": "Xodim sud bojlaridan 100% ozod qilingan",
                        "mediation_duration": "3 – 5 kun", "mediation_cost_estimate": "Mutlaqo bepul", "peaceful_resolution_chance": 88,
                        "why_mediation_better": "Rasmiy talabnoma olgan ish beruvchi qonuniy jarimalar va mehnat inspeksiyasi tekshiruvidan cho'chib, nizoni tinch va tez hal etadi."
                    },
                    "warnings": ["Mehnat nizolari bo'yicha sudga da'vo muddati ishdan bo'shatish to'g'risidagi buyruq berilgan kundan 3 oy ekanligini unutmang!"]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Ish beruvchiga rasmiy talabnoma yuborish", "description": "Tizim shakllantirgan talabnomani kadrlar bo'limi yoki rahbariyatga imzo qo'ydirib topshiring.", "deadline_hint": "Bugunoq" },
                    { "step_number": 2, "title": "Mehnat inspeksiyasi va sud", "description": "3 kun ichida to'lanmasa, my.mehnat.uz orqali ariza topshiring.", "deadline_hint": "3 kun ichida" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Mehnat nizolarida hech qanday davlat boji undirilmaydi." },
                "official_document": {
                    "is_generated": True, "document_type": "TALABNOMA",
                    "recipient_title": "«[Tashkilot / Korxona nomi]» rahbariga\nManzil: [Ish joyi manzili]",
                    "sender_placeholder": "Xodim: [F.I.SH.]dan\nLavozimi: [Egallagan lavozimi]\nTelefon: [Telefon raqami]",
                    "document_title": "ISH HAQI VA BARCHA TEGISHLI TO'LOVLARNI AMALGA OSHIRISH TO'G'RISIDA\nRASMIY TALABNOMA (PRETENZIYA)",
                    "document_body": f"Men ushbu korxonada o'z mehnat faoliyatimni lozim darajada bajarib kelganman. Biroq quyidagi holatlar yuzasidan mening mehnat huquqlarim buzildi:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi Mehnat kodeksining 253 va 269-moddalariga muvofiq, ish beruvchi xodimga tegishli barcha to'lovlarni o'z vaqtida, to'liq to'lashi shart. Majburiyatni bajarmaslik O'zR Ma'muriy javobgarlik to'g'risidagi kodeksining 49-moddasiga ko'ra jarimaga sabab bo'ladi.\n\nNizoni sud va Davlat mehnat inspeksiyasi aralashuvisiz tinch yo'l bilan hal qilish maqsadida SIZDAN TALAB QILAMAN:\n\n1. Mazkur talabnoma olingan kundan boshlab 3 (uch) bank kuni ichida barcha to'lanmagan oylik va kompensatsiya summalarini to'liq to'lab berishingizni;\n2. Aks holda, fuqarolik ishlari bo'yicha tumanlararo sudiga murojaat qilinib, majburiy progul kunlari va ma'naviy zarar undirilishini ma'lum qilaman.",
                    "attachments_list": ["Mehnat shartnomasi yoki buyruq nusxasi", "Bank plastik kartasi hisob raqami ma'lumotnomasi"]
                }
            }

    # ---------------- 2. OILA VA ALIMENT ----------------
    @classmethod
    def _generate_family_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "o'rtacha" if lang!='ru' else "средний", "case_category": "Oilaviy nizo (Aliment va Murosa kelishuvi)" if lang!='ru' else "Семейный спор (Алименты и Медиация)" },
            "executive_summary": {
                "headline": "Oila kodeksi bola ta'minotini qat'iy kafolatlaydi, mediasiya esa nizoni er-xotin o'rtasida tinch hal qiladi." if lang!='ru' else "Семейный кодекс РУз защищает интересы ребенка, а медиация позволяет избежать публичных судебных разбирательств.",
                "detailed_advice": f"Sizning murojaatingiz («{user_facts[:120]}...»): Oila kodeksining 96-moddasiga binoan ota-ona voyaga yetmagan bolalariga ta'minot berishi shart. Sudda oylarcha vaqt sarflash o'rniga, O'zR «Mediasiya to'g'risida»gi Qonuni asosida rasmiy Talabnoma yuborib, aliment to'lash to'g'risida notarial yoki mediasiya kelishuvi tuzish tavsiya etiladi." if lang!='ru' else f"По вашей ситуации («{user_facts[:120]}...»): в соответствии со статьей 96 Семейного кодекса РУз родители обязаны содержать своих несовершеннолетних детей. Заключение досудебного соглашения об уплате алиментов экономит время и нервы ребенка.",
                "legal_articles": [
                    { "code_name": "O'zR Oila kodeksi" if lang!='ru' else "Семейный кодекс РУз", "article": "96-modda. Ota-onaning bolalariga ta'minot berish majburiyati" if lang!='ru' else "Статья 96. Обязанности родителей по содержанию детей", "relevance": "Aliment to'lash qonuniy shart ekanligini belgilaydi." if lang!='ru' else "Устанавливает законную обязанность уплаты алиментов." },
                    { "code_name": "O'zR «Mediasiya to'g'risida»gi Qonuni" if lang!='ru' else "Закон РУз «О медиации»", "article": "15-modda. Mediasiya kelishuvi" if lang!='ru' else "Статья 15. Медиативное соглашение", "relevance": "Oilaviy masalalarni sudga chiqarmasdan o'zaro kelishuv bilan hal qilish imkonini beradi." if lang!='ru' else "Позволяет мирно урегулировать семейные споры." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True, "court_duration": "1 – 3 oy" if lang!='ru' else "1 – 3 месяца", "court_cost_estimate": "Davlat bojisiz" if lang!='ru' else "Без госпошлины",
                    "mediation_duration": "3 – 7 kun" if lang!='ru' else "3 – 7 дней", "mediation_cost_estimate": "Minimal" if lang!='ru' else "Минимально", "peaceful_resolution_chance": 85,
                    "why_mediation_better": "Bolaning ruhiyatiga putur yetkazmasdan, o'zaro grafik tuziladi." if lang!='ru' else "Сохраняет нормальные отношения ради ребенка без судебного стресса."
                },
                "warnings": ["Aliment to'lashdan qasddan bo'yin tovlash O'zR Jinoyat kodeksining 122-moddasi bilan jinoiy javobgarlikka sabab bo'ladi!" if lang!='ru' else "Уклонение от уплаты алиментов влечет уголовную ответственность по ст. 122 УК РУз!"]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Aliment talabnomasi yuborish" if lang!='ru' else "Направление требования", "description": "Tayyorlangan talabnomani ikkinchi tarafga topshiring." if lang!='ru' else "Вручите требование второму родителю.", "deadline_hint": "Bugunoq" if lang!='ru' else "Сегодня" },
                { "step_number": 2, "title": "Kelishuv imzolash" if lang!='ru' else "Подписание соглашения", "description": "Mediator yoki notarius huzurida tasdiqlang." if lang!='ru' else "Оформите соглашение у нотариуса или медиатора.", "deadline_hint": "5 kun ichida" if lang!='ru' else "В течение 5 дней" }
            ],
            "notary_module": {
                "is_relevant": True, "service_title": "Aliment to'lash to'g'risida kelishuvni tasdiqlash" if lang!='ru' else "Удостоверение соглашения об уплате алиментов",
                "required_documents": [
                    {"doc_name": "Bola(lar)ning tug'ilganlik guvohnomasi" if lang!='ru' else "Свидетельство о рождении ребенка", "where_to_get": "FHDYO (ZAGS)" if lang!='ru' else "ЗАГС"},
                    {"doc_name": "Ota-onaning pasportlari" if lang!='ru' else "Паспорта родителей", "where_to_get": "Shaxsiy hujjat" if lang!='ru' else "Личный документ"}
                ],
                "estimated_state_duty": "BHMning 20% miqdorida" if lang!='ru' else "20% от БРВ",
                "pro_tip": "Notarial tasdiqlangan aliment kelishuvi ijro varaqasi kuchiga ega (sudga berish shart emas)." if lang!='ru' else "Нотариальное соглашение об алиментах имеет силу исполнительного листа."
            },
            "official_document": {
                "is_generated": True, "document_type": "TALABNOMA" if lang!='ru' else "ТРЕБОВАНИЕ",
                "recipient_title": "Fuqaro [Ikkinchi taraf F.I.SH.]ga\nManzil: [Yashash manzili]" if lang!='ru' else "Гражданину(ке) [Ф.И.О. второго родителя]\nАдрес: [Адрес проживания]",
                "sender_placeholder": "Fuqaro: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]" if lang!='ru' else "От гражданина(ки): [Ф.И.О.]\nАдрес: [Адрес проживания]\nТел: [Номер телефона]",
                "document_title": "BOLA TA'MINOTI UCHUN ALIMENT TO'LASH VA MEDIASIYA KELISHUVI TUZISH TO'G'RISIDA\nTALABNOMA" if lang!='ru' else "ТРЕБОВАНИЕ\nО ВЫПЛАТЕ АЛИМЕНТОВ НА СОДЕРЖАНИЕ РЕБЕНКА И ЗАКЛЮЧЕНИИ СОГЛАШЕНИЯ",
                "document_body": f"O'rtamizdagi oilaviy munosabatlar yuzasidan ma'lum qilamanki:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi Oila kodeksining 96-moddasiga asosan ota-ona voyaga yetmagan bolalariga ta'minot berishi shart. Moddiy yordam berilmaganda aliment sud orqali undiriladi.\n\nO'zR «Mediasiya to'g'risida»gi Qonuniga muvofiq, sudga bormasdan, tinch yo'l bilan hal qilish maqsadida TAKLIF VA TALAB QILAMAN:\n\n1. 7 kun ichida o'zaro aliment to'lash to'g'risidagi kelishuvni imzolash uchun uchrashishimizni;\n2. Ushbu talab rad etilsa, sud orqali o'tgan davrlar uchun ham majburiy undirish chorasi ko'riladi.",
                "attachments_list": ["Bola(lar) tug'ilganlik guvohnomasi nusxasi", "Nikoh guvohnomasi (yoki bekor qilinganlik to'g'risida guvohnoma) nusxasi"]
            }
        }

    # ---------------- 3. ISTE'MOLCHILAR HUQUQLARI ----------------
    @classmethod
    def _generate_consumer_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "o'rtacha" if lang!='ru' else "средний", "case_category": "Iste'molchilar huquqlarini himoya qilish" if lang!='ru' else "Защита прав потребителей" },
            "executive_summary": {
                "headline": "Nuqsonli tovar uchun pulni qaytarib olish yoki yangisiga almashtirish iste'molchining mutlaq huquqidir." if lang!='ru' else "Потребитель имеет право на замену некачественного товара или полный возврат уплаченных денег.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): O'zR «Iste'molchilarning huquqlarini himoya qilish to'g'risida»gi Qonunining 13 va 14-moddalariga binoan xaridor sifatsiz tovar aniqlanganda pulni to'liq qaytarib olishga haqli. Sotuvchiga 7 kunlik rasmiy Talabnoma yuborish kifoya. Sotuvchi bajarmasa, Raqobat qo'mitasi jarimaga tortadi.",
                "legal_articles": [
                    { "code_name": "O'zR «Iste'molchilarning huquqlarini himoya qilish to'g'risida»gi Qonuni", "article": "13-modda. Nuqsonli tovar sotilganda iste'molchining huquqlari", "relevance": "Tovarni sifatlisiga almashtirish yoki pulni to'liq qaytarib olish kafolatini beradi." },
                    { "code_name": "O'zR «Iste'molchilarning huquqlarini himoya qilish to'g'risida»gi Qonuni", "article": "19-modda. Xizmat ko'rsatish qoidalarini buzganlik uchun javobgarlik", "relevance": "Kechiktirilgan har bir kun uchun penya undirilishini belgilaydi." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True, "court_duration": "1 – 2 oy", "court_cost_estimate": "Iste'molchilar davlat bojisiz da'vo kiritadi",
                    "mediation_duration": "2 – 4 kun", "mediation_cost_estimate": "Mutlaqo bepul", "peaceful_resolution_chance": 90,
                    "why_mediation_better": "Sotuvchi Raqobat qo'mitasi tekshiruvidan qochib, pulni shu zahoti qaytarib beradi."
                },
                "warnings": ["Xarid cheki, tovar qutisi yoki to'lov kvitansiyasini saqlab qo'ying."]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Sotuvchiga talabnoma topshirish", "description": "Tayyor talabnomani do'kon ma'muriyatiga taqdim eting.", "deadline_hint": "Bugun" },
                { "step_number": 2, "title": "Raqobat qo'mitasiga shikoyat", "description": "Agar 7 kunda almashtirmasa, 1159 raqamiga xabar bering.", "deadline_hint": "7 kun ichida" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Iste'molchilar huquqlari bo'yicha notarius talab etilmaydi." },
            "official_document": {
                "is_generated": True, "document_type": "TALABNOMA",
                "recipient_title": "«[Do'kon / Sotuvchi tashkilot nomi]» ma'muriyatiga\nManzil: [Do'kon manzili]",
                "sender_placeholder": "Iste'molchi: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]",
                "document_title": "NUQSONLI TOVAR UCHUN TO'LANGAN MABLAG'NI QAYTARISH TO'G'RISIDA\nTALABNOMA (PRETENZIYA)",
                "document_body": f"Men Sizning savdo do'koningizdan xaridni amalga oshirgan edim. Biroq quyidagi holat aniqlandi:\n\n«{user_facts}»\n\nO'zbekiston Respublikasining «Iste'molchilarning huquqlarini himoya qilish to'g'risida»gi Qonunining 13, 14 va 18-moddalariga asosan, nuqsonli tovar sotilganda iste'molchi shartnomani bekor qilish va to'langan pul summasini to'liq qaytarib olishga haqli.\n\nShu munosabat bilan, qonunda belgilangan tartibda SIZDAN TALAB QILAMAN:\n\n1. Ushbu talabnoma olingan kundan boshlab 7 kun ichida to'langan mablag'ni to'liq qaytarishingizni;\n2. Talab qanoatlantirilmagan taqdirda, O'zR Raqobatni rivojlantirish qo'mitasiga va sudga murojaat qilinishini ma'lum qilaman.",
                "attachments_list": ["Xarid cheki yoki to'lov tasdig'i nusxasi", "Tovardagi nuqson fotosurati"]
            }
        }

    # ---------------- 4. YETKAZILGAN ZARAR VA YTH ----------------
    @classmethod
    def _generate_damages_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "yuqori" if lang!='ru' else "высокий", "case_category": "Zararni qoplash (YTH / Moddiy zarar)" if lang!='ru' else "Возмещение ущерба (ДТП / Имущественный вред)" },
            "executive_summary": {
                "headline": "O'zR Fuqarolik kodeksining 985-moddasi yetkazilgan zararni aybdor shaxs tomonidan to'liq qoplanishini shart qilib belgilaydi.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): Fuqarolik kodeksiga binoan boshqa shaxsning mol-mulkiga yetkazilgan zarar aybdor tomonidan to'liq hajmda qoplanishi shart. Sudda oylab sarson bo'lmaslik uchun, baholash dalolatnomasi asosida aybdor tarafga sudgacha Talabnoma yuborish va Mediasiya orqali to'lov muddatini kelishib olish tavsiya etiladi.",
                "legal_articles": [
                    { "code_name": "O'zR Fuqarolik kodeksi", "article": "985-modda. Zarar yetkazganlik uchun javobgarlikning umumiy asoslari", "relevance": "Yetkazilgan moddiy va ma'naviy zarar aybdor tomonidan to'liq hajmda to'lanishini ta'minlaydi." },
                    { "code_name": "O'zR Fuqarolik kodeksi", "article": "1005-modda. Yetkazilgan zararni to'lash usullari", "relevance": "Zararni natura shaklida yoki yetkazilgan zararni to'liq pulda qoplashni belgilaydi." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True, "court_duration": "2 – 4 oy", "court_cost_estimate": "Zarar summasining 4% davlat boji",
                    "mediation_duration": "3 – 7 kun", "mediation_cost_estimate": "Minimal", "peaceful_resolution_chance": 85,
                    "why_mediation_better": "Aybdor sud xarajatlari va jarimalardan qochib, zararni bo'lib to'lashga tezroq rozi bo'ladi."
                },
                "warnings": ["YTH ma'lumotnomasi, baholovchi tashkilot xulosasi va fotosuratlarni saqlang."]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Baholash xulosasini olish", "description": "Mustaqil baholovchi orqali yetkazilgan zarar miqdorini aniqlang.", "deadline_hint": "1-2 kunda" },
                { "step_number": 2, "title": "Talabnomani topshirish", "description": "Aybdor tarafga va sug'urta kompaniyasiga talabnoma yuboring.", "deadline_hint": "3 kun ichida" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Zararni qoplash bo'yicha kelishuvni notarial tasdiqlash mumkin." },
            "official_document": {
                "is_generated": True, "document_type": "TALABNOMA",
                "recipient_title": "Fuqaro [Aybdor shaxs F.I.SH.]ga\nManzil: [Yashash manzili]",
                "sender_placeholder": "Fuqaro: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]",
                "document_title": "YETKAZILGAN MODDIY ZARARNI IHTIYORIY QOPLASH TO'G'RISIDA\nTALABNOMA (PRETENZIYA)",
                "document_body": f"Sizning harakatlaringiz (harakatsizligingiz) oqibatida menga tegishli mulkka zarar yetkazildi:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi Fuqarolik kodeksining 985-moddasiga asosan, fuqaroning shaxsiga yoki mol-mulkiga yetkazilgan zarar uni yetkazgan shaxs tomonidan to'liq hajmda qoplanishi lozim.\n\nNizoni sudga olib chiqmasdan, ortiqcha davlat bojlari (zarar summasining 4 foizi) va ekspertiza xarajatlarining oldini olish maqsadida TALAB QILAMAN:\n\n1. Ushbu talabnoma olingan kundan boshlab 7 kun ichida yetkazilgan zararni ixtiyoriy to'lab berishingizni;\n2. Rad etilgan taqdirda fuqarolik ishlari bo'yicha sudga da'vo kiritilib, barcha advokat va davlat bojlari Sizdan undiriladi.",
                "attachments_list": ["Baholash tashkiloti xulosasi nusxasi", "Zarar yetkazilganligini tasdiqlovchi dalolatnoma / foto"]
            }
        }

    # ---------------- 5. IJARA NIZOSI ----------------
    @classmethod
    def _generate_rent_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "mediation", "secondary_intent": "appeal", "urgency_level": "o'rtacha", "case_category": "Uy-joy va Ijara nizosi" },
            "executive_summary": {
                "headline": "O'zR Fuqarolik kodeksining 553-moddasi ijara haqini o'z vaqtida to'lashni shart qilib qo'yadi.",
                "detailed_advice": f"Sizning masalangiz («{user_facts[:120]}...»): Ijara shartnomasi bo'yicha ijarachi o'z majburiyatlarini buzgan taqdirda, mulkdor shartnomani bekor qilish va qarzni undirishga haqli. Birinchi qadam sifatida 7 kunlik rasmiy Talabnoma yuborish va kelishuv tuzish shart.",
                "legal_articles": [
                    { "code_name": "O'zR Fuqarolik kodeksi", "article": "553-modda. Mulk ijarasi shartnomasi bo'yicha haq to'lash", "relevance": "Ijara haqini shartnomada ko'rsatilgan muddatda to'lash majburiyatini belgilaydi." },
                    { "code_name": "O'zR Fuqarolik kodeksi", "article": "600-modda. Turar joyni ijaraga berish shartnomasini bekor qilish", "relevance": "Haq to'lanmaganda ijarachini uydan chiqarish va shartnomani bekor qilish asoslarini beradi." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True, "court_duration": "2 – 3 oy", "court_cost_estimate": "Qarz summasining 4% davlat boji",
                    "mediation_duration": "3 – 5 kun", "mediation_cost_estimate": "Minimal", "peaceful_resolution_chance": 85,
                    "why_mediation_better": "Sud orqali uydan chiqarish oylab cho'ziladi, mediasiya orqali muddat va to'lov grafigi 3 kunda kelishiladi."
                },
                "warnings": ["Ijara shartnomasi soliq organlarida (ijara.soliq.uz) ro'yxatdan o'tganligini tekshiring."]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Ijara talabnomasi topshirish", "description": "Qarzni to'lash yoki uyni bo'shatish haqida talabnoma yuboring.", "deadline_hint": "Bugun" },
                { "step_number": 2, "title": "Mediasiya / Kelishuv", "description": "Qarzni to'lash grafigi bo'yicha tilxat yoki kelishuv oling.", "deadline_hint": "5 kun ichida" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Ijara shartnomasi notarial tasdiqlangan bo'lsa, sudsiz MIB orqali undiriladi." },
            "official_document": {
                "is_generated": True, "document_type": "TALABNOMA",
                "recipient_title": "Ijarachi: [F.I.SH.]ga\nManzil: [Ijara uy manzili]",
                "sender_placeholder": "Mulkdor: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]",
                "document_title": "IJARA HAQI QARZDORLIGINI TO'LASH VA TURAR JOYNI BO'SHATISH TO'G'RISIDA\nTALABNOMA",
                "document_body": f"O'rtamizda tuzilgan ijara munosabatlariga asosan:\n\n«{user_facts}»\n\nO'zR Fuqarolik kodeksining 553 va 600-moddalariga muvofiq, ijara haqini muntazam to'lamaslik shartnomani bekor qilish va yetkazilgan zararni undirish uchun asos hisoblanadi.\n\nShu munosabat bilan, SIZDAN TALAB QILAMAN:\n\n1. Ushbu talabnoma olingan kundan boshlab 5 (besh) bank kuni ichida ijara qarzdorligini to'liq so'ndirishingizni;\n2. Aks holda shartnoma bekor qilinib, majburiy tartibda ko'chirish va qarzdorlikni undirish bo'yicha sudga murojaat qilinadi.",
                "attachments_list": ["Ijara shartnomasi nusxasi", "Hisoblangan qarzdorlik ma'lumotnomasi"]
            }
        }

    # ---------------- 6. MA'MURIY JARIMALAR ----------------
    @classmethod
    def _generate_admin_fine_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "appeal", "secondary_intent": "mediation", "urgency_level": "yuqori", "case_category": "Ma'muriy jarima ustidan shikoyat" },
            "executive_summary": {
                "headline": "MJtKning 315-moddasiga ko'ra noqonuniy yoki asossiz qo'llanilgan jarima qarori ustidan 10 kun ichida shikoyat berish huquqiga egasiz.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): Yo'l harakati qoidabuzarligi yoki ma'muriy jarima asossiz qo'llanilgan bo'lsa, qaror chiqarilgan kundan boshlab 10 kun ichida yuqori turuvchi organga yoki ma'muriy sudga Shikoyat berilishi lozim. Shikoyat berilganda jarimani ijro etish vaqtincha to'xtatiladi.",
                "legal_articles": [
                    { "code_name": "O'zR Ma'muriy javobgarlik to'g'risidagi kodeksi", "article": "315-modda. Ma'muriy ish bo'yicha qaror ustidan shikoyat berish", "relevance": "Qaror nusxasi topshirilgan kundan 10 kun ichida bepul shikoyat berish huquqini beradi." },
                    { "code_name": "O'zR MJtK", "article": "318-modda. Shikoyat berilishi munosabati bilan qaror ijrosini to'xtatib turish", "relevance": "Shikoyat ko'rib chiqilguniga qadar jarimani to'lash to'xtatilishini kafolatlaydi." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": False, "court_duration": "10 – 15 kun", "court_cost_estimate": "Ma'muriy shikoyat bo'yicha bojsiz",
                    "mediation_duration": "3 – 5 kun", "mediation_cost_estimate": "Bepul", "peaceful_resolution_chance": 80,
                    "why_mediation_better": "Yuqori organga rasmiy shikoyat kiritish orqali xatolik sudgacha bekor qilinadi."
                },
                "warnings": ["10 kunlik shikoyat berish muddatini o'tkazib yubormang!"]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Qaror nusxasini olish", "description": "Jarima qarori raqami va sanasini aniqlang.", "deadline_hint": "Darhol" },
                { "step_number": 2, "title": "Shikoyat yuborish", "description": "YHQB / YHXBB yoki my.gov.uz orqali rasmiy shikoyat jo'nating.", "deadline_hint": "10 kun ichida" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Jarima ustidan shikoyat uchun boj undirilmaydi." },
            "official_document": {
                "is_generated": True, "document_type": "SHIKOYAT",
                "recipient_title": "[Tuman/Shahar] IIB YHXBB boshlig'iga\nNusxasi: Ma'muriy sudga",
                "sender_placeholder": "Fuqaro: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]",
                "document_title": "MA'MURIY HUQUQBUZARLIK TO'G'RISIDAGI QAROR USTIDAN\nSHIKOYAT",
                "document_body": f"Menga nisbatan ma'muriy jarima qo'llash to'g'risida qaror chiqarilgan. Biroq ushbu holat bo'yicha quyidagilarni ma'lum qilaman:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi MJtKning 315-moddasiga ko'ra, ma'muriy huquqbuzarlik to'g'risidagi qaror ustidan yuqori turuvchi organga shikoyat berish huquqiga egaman. MJtK 318-moddasiga muvofiq shikoyat berilishi qaror ijrosini to'xtatib turadi.\n\nYuqoridagilardan kelib chiqib, SIZDAN ILTIMOS QILAMAN:\n\n1. Chiqarilgan noqonuniy/asossiz jarima to'g'risidagi qarorni bekor qilishingizni;\n2. Ma'muriy ish yuritishni harakatdan to'xtatishingizni.",
                "attachments_list": ["Jarima qarori nusxasi", "Asossiz ekanligini tasdiqlovchi dalillar (video/foto/avto hujjat)"]
            }
        }

    # ---------------- 7. NOTARIAT VA MEROS ----------------
    @classmethod
    def _generate_notary_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": {
                "detected_language": lang, "primary_intent": "notary", "secondary_intent": "mediation",
                "urgency_level": "o'rtacha" if lang!='ru' else "средний",
                "case_category": "Notariat va Meros huquqi" if lang!='ru' else "Нотариат и Наследственное право"
            },
            "executive_summary": {
                "headline": "Meros va mulkni notarial rasmiylashtirishda 6 oylik muddat va zarur hujjatlar talab etiladi." if lang!='ru' else "Оформление наследства требует подачи заявления нотариусу в течение 6 месяцев.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): Meros ochilgan kundan boshlab 6 oy ichida notariusga ariza topshirilishi shart. Davlat xizmatlari markazi yoki my.gov.uz orqali hujjatlarni olish xususiy vositachilarga sarflanadigan mablag'ni tejaydi." if lang!='ru' else "Заявление о принятии наследства должно быть подано нотариусу в течение 6 месяцев со дня открытия.",
                "legal_articles": [
                    { "code_name": "O'zR Fuqarolik kodeksi" if lang!='ru' else "Гражданский кодекс РУз", "article": "1146-modda. Merosni qabul qilib olish" if lang!='ru' else "Статья 1146. Принятие наследства", "relevance": "6 oylik muddatni belgilaydi." if lang!='ru' else "Устанавливает 6-месячный срок." },
                    { "code_name": "O'zR «Davlat boji to'g'risida»gi Qonuni" if lang!='ru' else "Закон РУз «О государственной пошлине»", "article": "Notarial stavkalar" if lang!='ru' else "Ставки пошлины", "relevance": "BHM bo'yicha imtiyozli boj stavkalari." if lang!='ru' else "Льготные ставки пошлины." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": { "show_comparison": False, "court_duration": "1 – 3 oy", "court_cost_estimate": "Mulkning 4%", "mediation_duration": "1 – 3 kun", "mediation_cost_estimate": "Belgilangan davlat boji", "peaceful_resolution_chance": 95, "why_mediation_better": "Merosxo'rlar kelishuvi bo'lsa sudga hojat qolmaydi." },
                "warnings": ["6 oylik muddat o'tib ketsa faqat sud orqali tiklanadi!"]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Notariusga ariza topshirish", "description": "E-notarius.uz orqali navbat oling.", "deadline_hint": "6 oy ichida" }
            ],
            "notary_module": {
                "is_relevant": True, "service_title": "Meros guvohnomasi berish",
                "required_documents": [
                    {"doc_name": "O'lim haqidagi guvohnoma", "where_to_get": "FHDYO / my.gov.uz"},
                    {"doc_name": "Kadastr pasporti", "where_to_get": "Davlat kadastrlari palatasi / DXM"},
                    {"doc_name": "Qarindoshlik hujjati", "where_to_get": "FHDYO"}
                ],
                "estimated_state_duty": "BHMning 50% miqdorida (taxminan 187 500 so'm)",
                "pro_tip": "E-notarius.uz orqali navbatga yozilish vaqtni tejaydi."
            },
            "official_document": {
                "is_generated": True, "document_type": "ARIZA" if lang!='ru' else "ЗАЯВЛЕНИЕ",
                "recipient_title": "Davlat notarial idorasi notariusiga" if lang!='ru' else "Нотариусу государственной нотариальной конторы",
                "sender_placeholder": "Arizachi: [F.I.SH.]dan" if lang!='ru' else "От заявителя: [Ф.И.О.]",
                "document_title": "MEROSNI QABUL QILIB OLISH VA MEROS GUCHNOMASINI BERISH TO'G'RISIDA\nARIZA" if lang!='ru' else "ЗАЯВЛЕНИЕ\nО ПРИНЯТИИ НАСЛЕДСТВА",
                "document_body": f"Mening murojaatim mazmuni:\n«{user_facts}»\n\nO'zbekiston Respublikasi Fuqarolik kodeksining 1146-moddasiga muvofiq, marhumdan qolgan merosni qabul qilganimni ma'lum qilib, qonun bo'yicha meros guvohnomasi berishingizni so'rayman.",
                "attachments_list": ["O'lim haqidagi guvohnoma nusxasi", "Qarindoshlik guvohnomasi nusxasi"]
            }
        }

    # ---------------- 8. KOMMUNAL VA HOKIMIYAT ----------------
    @classmethod
    def _generate_communal_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": { "detected_language": lang, "primary_intent": "appeal", "secondary_intent": "mediation", "urgency_level": "yuqori" if lang!='ru' else "высокий", "case_category": "Davlat organlariga murojaat (Kommunal va Hokimiyat)" if lang!='ru' else "Обращение в государственные органы (ЖКХ)" },
            "executive_summary": {
                "headline": "O'RQ-445 qonuni asosida davlat idoralari murojaatni 15 kunda ko'rib chiqishi shart." if lang!='ru' else "По закону ЗРУ-445 госорганы обязаны рассмотреть обращение в течение 15 дней.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): Kommunal xizmatlarning asossiz uzilishi bo'yicha O'RQ-445 qonuniga mos rasmiy shikoyat tayyorlandi. Tuman hokimligi va Xalq qabulxonasi 15 kun ichida yozma javob berishga majbur." if lang!='ru' else "Сформирована официальная жалоба по вопросу работы коммунальных служб.",
                "legal_articles": [
                    { "code_name": "O'RQ-445" if lang!='ru' else "ЗРУ-445", "article": "19-modda" if lang!='ru' else "Статья 19", "relevance": "Murojaatni ko'rib chiqish muddatlari" if lang!='ru' else "Сроки рассмотрения обращений" },
                    { "code_name": "Iste'molchilar huquqlari to'g'risida", "article": "4-modda", "relevance": "Sifatli kommunal xizmat olish huquqi" }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": { "show_comparison": False, "court_duration": "1 – 2 oy", "court_cost_estimate": "Bojsiz", "mediation_duration": "3 – 5 kun", "mediation_cost_estimate": "Bepul", "peaceful_resolution_chance": 90, "why_mediation_better": "Xalq qabulxonasi va prokuratura nazorati orqali muammo tez hal bo'ladi." },
                "warnings": ["15 kunlik qonuniy javob muddatini nazorat qiling."]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Arizani topshirish", "description": "My.gov.uz yoki Xalq qabulxonasiga yuboring.", "deadline_hint": "Bugunoq" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Zaruriyat yo'q", "required_documents": [], "estimated_state_duty": "0 so'm", "pro_tip": "Davlat idoralariga murojaat uchun boj undirilmaydi." },
            "official_document": {
                "is_generated": True, "document_type": "SHIKOYAT" if lang!='ru' else "ЖАЛОБА",
                "recipient_title": "Xalq qabulxonasiga / Tuman hokimligiga" if lang!='ru' else "В Народную приемную / Хокимият",
                "sender_placeholder": "Fuqaro: [F.I.SH.]dan" if lang!='ru' else "От гражданина(ки): [Ф.И.О.]",
                "document_title": "KOMMUNAL TA'MINOTNING ASOSSIZ UZILISHI VA HUQUQLAR BUZILGANLIGI YUZASIDAN\nSHIKOYAT" if lang!='ru' else "ЖАЛОБА\nНА НЕУДОВЛЕТВОРИТЕЛЬНОЕ ОКАЗАНИЕ КОММУНАЛЬНЫХ УСЛУГ",
                "document_body": f"Mening murojaatim mazmuni:\n«{user_facts}»\n\nO'zbekiston Respublikasining «Jismoniy va yuridik shaxslarning murojaatlari to'g'risida»gi Qonuniga muvofiq, ko'rsatilgan holatni mutaxassislar ishtirokida to'liq tekshirib, 15 kun ichida yozma javob berishingizni va kamchiliklarni zudlik bilan bartaraf etishingizni so'rayman.",
                "attachments_list": ["To'lov kvitansiyalari nusxasi", "Uzilishlar holatini tasdiqlovchi dalolatnoma / foto"]
            }
        }

    # ---------------- 9. QARZ VA TILXAT ----------------
    @classmethod
    def _generate_debt_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        return {
            "meta": {
                "detected_language": lang, "primary_intent": "mediation", "secondary_intent": "appeal",
                "urgency_level": "o'rtacha" if lang!='ru' else "средний",
                "case_category": "Fuqarolik-huquqiy nizo (Qarzni undirish va Majburiyatlar)" if lang!='ru' else "Гражданско-правовой спор (Взыскание долга)"
            },
            "executive_summary": {
                "headline": "O'zR Fuqarolik kodeksining 732-moddasi qarzni o'z vaqtida qaytarishni qat'iy majburiyat qilib qo'yadi." if lang!='ru' else "Статья 732 Гражданского кодекса РУз обязывает должника своевременно вернуть сумму займа.",
                "detailed_advice": f"Sizning holatingiz («{user_facts[:120]}...»): Sudda da'vo summasidan 4% davlat boji to'lash va 2–4 oy kutish o'rniga, O'zR «Mediasiya to'g'risida»gi Qonuni asosida 7 kunlik rasmiy Talabnoma yuborish va mediasiya kelishuvini tuzish tavsiya etiladi. Bu qarzni tinch va tezroq qaytarish ehtimolini 85% ga oshiradi." if lang!='ru' else "Рекомендуется направить официальную претензию со сроком возврата 7 дней и предложить заключение медиативного соглашения.",
                "legal_articles": [
                    { "code_name": "O'zR Fuqarolik kodeksi" if lang!='ru' else "Гражданский кодекс РУз", "article": "732-modda. Qarz shartnomasi" if lang!='ru' else "Статья 732. Договор займа", "relevance": "Qarz oluvchining qarzni qaytarish majburiyatini belgilaydi." if lang!='ru' else "Обязывает заемщика вернуть полученные средства." },
                    { "code_name": "O'zR «Mediasiya to'g'risida»gi Qonuni" if lang!='ru' else "Закон РУз «О медиации»", "article": "15-modda. Mediasiya kelishuvi" if lang!='ru' else "Статья 15. Медиативное соглашение", "relevance": "Nizoni sudga bormasdan qonuniy kuchga ega kelishuv bilan hal qiladi." if lang!='ru' else "Придает медиативному соглашению обязательную силу." }
                ]
            },
            "risk_and_comparison": {
                "court_vs_mediation": {
                    "show_comparison": True, "court_duration": "2 – 4 oy" if lang!='ru' else "2 – 4 месяца", "court_cost_estimate": "Da'vo summasining 4% davlat boji" if lang!='ru' else "4% госпошлина + расходы",
                    "mediation_duration": "3 – 7 kun" if lang!='ru' else "3 – 7 дней", "mediation_cost_estimate": "Minimal" if lang!='ru' else "Минимально", "peaceful_resolution_chance": 85,
                    "why_mediation_better": "Mediasiya ortiqcha sud xarajatlari va asab buzilishisiz, o'zaro manfaatli to'lov grafigini beradi." if lang!='ru' else "Медиация исключает оплату пошлины и сохраняет отношения."
                },
                "warnings": ["Umumiy da'vo muddati 3 yil ekanligini unutmang. Tilxat asl nusxasini ehtiyot qiling." if lang!='ru' else "Срок исковой давности составляет 3 года. Сохраняйте оригинал расписки."]
            },
            "action_roadmap": [
                { "step_number": 1, "title": "Talabnomani topshirish", "description": "Qarzdorga tilxat yoki pochta orqali yuboring.", "deadline_hint": "Bugunoq" },
                { "step_number": 2, "title": "Mediasiya muzokaralari", "description": "Qarzni bo'lib to'lash grafigini imzolang.", "deadline_hint": "7 kun ichida" }
            ],
            "notary_module": { "is_relevant": False, "service_title": "Notarial ijro xati", "required_documents": [], "estimated_state_duty": "1%", "pro_tip": "Agar qarz notarial tasdiqlangan bo'lsa, sudsiz MIBga ijroga beriladi." },
            "official_document": {
                "is_generated": True, "document_type": "TALABNOMA" if lang!='ru' else "ПРЕТЕНЗИЯ",
                "recipient_title": "Qarzdor: [F.I.SH.]ga\nManzil: [Yashash manzili]\nTel: [Telefon raqami]" if lang!='ru' else "Должнику: [Ф.И.О.]\nАдрес: [Адрес]\nТел: [Телефон]",
                "sender_placeholder": "Qarz beruvchi: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]" if lang!='ru' else "Кредитор: [Ф.И.О.]\nАдрес: [Адрес]\nТел: [Телефон]",
                "document_title": "QARZ MAJBURIYATINI SUDGACHA HAL ETISH VA MEDIASIYA TAKLIFI TO'G'RISIDA\nTALABNOMA (PRETENZIYA)" if lang!='ru' else "ПРЕТЕНЗИЯ (ТРЕБОВАНИЕ)\nО ДОСУДЕБНОМ ВОЗВРАТЕ СУММЫ ДОЛГА И МЕДИАЦИИ",
                "document_body": f"O'rtamizda vujudga kelgan qarz munosabatlari yuzasidan:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi Fuqarolik kodeksining 732-moddasiga binoan, qarz oluvchi olingan pul summasini shartnomada belgilangan muddatda va tartibda qaytarishi shart.\n\nO'zR «Mediasiya to'g'risida»gi Qonuniga muvofiq, sudga chiqmasdan, qo'shimcha sud xarajatlarining (4 foiz davlat boji) oldini olgan holda hal qilish maqsadida SIZGA TALAB QILAMAN:\n\n1. Ushbu talabnoma olingan kundan boshlab 7 (yetti) bank kuni ichida qarz summasini to'liq qaytarishingizni;\n2. Mablag' yetishmasa, mediator ishtirokida qarzni bo'lib to'lash grafigi bo'yicha Mediasiya kelishuvini imzolashni.\n\nAks holda fuqarolik ishlari bo'yicha sudga da'vo arizasi kiritilib, qarz, bojlar va penya Sizdan undiriladi.",
                "attachments_list": ["Qarz tilxati yoki shartnoma nusxasi", "To'lov o'tkazmalari nusxasi"]
            }
        }

    # ---------------- 10. UNIVERSAL ERKIN MUAMMO (HAR QANDAY BOSHQALARI) ----------------
    @classmethod
    def _generate_universal_custom_case(cls, lang: str, user_facts: str) -> Dict[str, Any]:
        if lang == 'ru':
            return {
                "meta": {
                    "detected_language": "ru",
                    "primary_intent": "mediation",
                    "secondary_intent": "appeal",
                    "urgency_level": "средний",
                    "case_category": "Гражданско-правовые отношения и Досудебное урегулирование"
                },
                "executive_summary": {
                    "headline": "По вашему обращению сформирована правовая позиция и проект официального досудебного требования по законодательству РУз.",
                    "detailed_advice": f"По описанной вами ситуации («{user_facts[:120]}...»): на основании статьи 9 Гражданского кодекса РУз граждане имеют полное право на досудебную защиту своих законных интересов. Поскольку судебное разбирательство занимает от 2 до 4 месяцев и требует оплаты 4% госпошлины, в первую очередь рекомендуется направить официальную Претензию со сроком добровольного исполнения 7 дней. В случае отказа спор передается в суд с возложением всех расходов на виновную сторону.",
                    "legal_articles": [
                        { "code_name": "Гражданский кодекс Республики Узбекистан", "article": "Статья 9. Защита гражданских прав", "relevance": "Гарантирует защиту нарушенных гражданских прав законными способами." },
                        { "code_name": "Закон Республики Узбекистан «О медиации»", "article": "Статья 15. Медиативное соглашение", "relevance": "Обеспечивает обязательную юридическую силу мирного соглашения." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "2 – 4 месяца", "court_cost_estimate": "4% госпошлина + адвокат",
                        "mediation_duration": "3 – 7 дней", "mediation_cost_estimate": "Минимально", "peaceful_resolution_chance": 85,
                        "why_mediation_better": "Экономит расходы и нервы, обеспечивая быстрое мирное соглашение без публичных судебных тяжб."
                    },
                    "warnings": ["Сохраняйте все чеки, переписку, квитанции и документы как письменные доказательства."]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Вручение официальной претензии", "description": "Вручите подготовленное требование второй стороне под расписку или заказным письмом.", "deadline_hint": "Сегодня" },
                    { "step_number": 2, "title": "Переговоры и Медиация", "description": "Заключите соглашение на взаимовыгодных условиях с фиксацией сроков.", "deadline_hint": "В течение 7 дней" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Нотариальная консультация", "required_documents": [], "estimated_state_duty": "Не требуется", "pro_tip": "При необходимости достигнутое мировое соглашение можно удостоверить у нотариуса." },
                "official_document": {
                    "is_generated": True, "document_type": "ПРЕТЕНЗИЯ",
                    "recipient_title": "Второй стороне спора: [Ф.И.О. или Наименование организации]\nАдрес: [Адрес]",
                    "sender_placeholder": "От заявителя: [Ф.И.О.]\nАдрес: [Адрес проживания]\nТел: [Номер телефона]",
                    "document_title": "ПРЕТЕНЗИЯ (ДОСУДЕБНОЕ ТРЕБОВАНИЕ)\nО МИРНОМ УРЕГУЛИРОВАНИИ СПОРА И МЕДИАЦИИ",
                    "document_body": f"Настоящим обращаюсь к вам по следующему вопросу:\n\n«{user_facts}»\n\nВ соответствии со статьями Гражданского кодекса Республики Узбекистан стороны обязаны надлежащим образом исполнять взятые на себя обязательства.\n\nВ соответствии с Законом Республики Узбекистан «О медиации», в целях мирного урегулирования спора без обращения в суд и предотвращения судебных расходов (4% госпошлины), ТРЕБУЮ:\n\n1. В течение 7 (семи) банковских дней со дня получения настоящей претензии добровольно устранить допущенные нарушения и выполнить законные требования;\n2. В противном случае спор будет передан в суд с возложением всех убытков, судебных расходов и процентов на вас.",
                    "attachments_list": ["Копии подтверждающих документов", "Копия документа, удостоверяющего личность"]
                }
            }
        elif lang == 'uz_cyrl':
            return {
                "meta": {
                    "detected_language": "uz_cyrl",
                    "primary_intent": "mediation",
                    "secondary_intent": "appeal",
                    "urgency_level": "ўртача",
                    "case_category": "Фуқаролик-ҳуқуқий муносабатлар ва Муроса"
                },
                "executive_summary": {
                    "headline": "Киритилган муаммоли ҳолат бўйича ЎзР қонунчилиги асосида судгача расмий талабнома ва медиация стратегияси белгиланди.",
                    "detailed_advice": f"Сиз киритган масала («{user_facts[:120]}...»): Ўзбекистон Республикаси Фуқаролик кодексининг 9-моддасига кўра, фуқаролар ўз ҳуқуқларини судгача талабнома тақдим этиш ва қонуний воситалар орқали ҳимоя қилишга тўлиқ ҳақли. Судга чиқиш 2 дан 4 ойгача вақт ва даъво суммасининг 4% давлат божини талаб этганлиги сабабли, биринчи навбатда иккинчи тарафга 7 кунлик расмий Талабнома (Претензия) юбориш тавсия этилади. Ушбу ҳужжат сизнинг қонуний талабларингизни тўлиқ ифодалайди.",
                    "legal_articles": [
                        { "code_name": "Ўзбекистон Республикаси Фуқаролик кодекси", "article": "9-модда. Фуқаролик ҳуқуқларини ҳимоя қилиш", "relevance": "Фуқароларга ўз бузилган ҳуқуқларини қонуний воситалар орқали судгача ҳимоя қилиш кафолатини беради." },
                        { "code_name": "Ўзбекистон Республикасининг «Медиация тўғрисида»ги Қонуни", "article": "15-модда. Медиация келишуви", "relevance": "Низоларни судга чиқармасдан тарафларнинг ўзаро келишуви орқали ҳал қилиш кучини таъминлайди." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "2 – 4 ой", "court_cost_estimate": "Даъво суммасининг 4% давлат божи + адвокат",
                        "mediation_duration": "3 – 7 кун", "mediation_cost_estimate": "Минимал", "peaceful_resolution_chance": 85,
                        "why_mediation_better": "Суд харажатлари ва асаб бузилишисиз, ўзаро манфаатли график асосида тинч ечим беради."
                    },
                    "warnings": ["Барча мавжуд ёзишмалар, шартномалар, чеклар ва далилларни тартибга солиб сақланг."]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Расмий талабнома топшириш", "description": "Тайёрланган талабномани иккинчи тарафга тилхат ёки почта орқали топширинг.", "deadline_hint": "Бугуноқ" },
                    { "step_number": 2, "title": "Медиация ва музокара", "description": "Келишилган шартлар асосида тинч битим имзоланг.", "deadline_hint": "7 кун ичида" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Нотариал маслаҳат", "required_documents": [], "estimated_state_duty": "Белгиланмаган", "pro_tip": "Зарурат туғилса тузилган келишувни нотариусда тасдиқлаш мумкин." },
                "official_document": {
                    "is_generated": True, "document_type": "ТАЛАБНОМА",
                    "recipient_title": "Низо бўйича иккинчи тараф: [Ф.И.Ш. ёки Ташкилот номи]га\nМанзил: [Манзил]",
                    "sender_placeholder": "Мурожаатчи: [Ф.И.Ш.]дан\nМанзил: [Яшаш манзили]\nТел: [Телефон рақами]",
                    "document_title": "НИЗОЛИ ВАЗИЯТНИ СУДГАЧА ҲАЛ ЭТИШ ВА МЕДИАЦИЯ ТАКЛИФИ ТЎҒРИСИДА\nТАЛАБНОМА (ПРЕТЕНЗИЯ)",
                    "document_body": f"Мазкур талабнома орқали қуйидаги ҳолат юзасидан мурожаат қилмоқдаман:\n\n«{user_facts}»\n\nЎзбекистон Республикаси Фуқаролик кодексининг тегишли моддаларига асосан тарафлар ўз зиммаларидаги мажбуриятларни лозим даражада бажаришлари шарт.\n\nЎзбекистон Республикасининг «Медиация тўғрисида»ги Қонунига мувофиқ, юзага келган ушбу низоли ҳолатни судга олиб чиқмасдан, ортиқча давлат божлари (даъво суммасининг 4 фоизи) ва суд харажатларининг олдини олган ҳолда ҳал қилиш мақсадида СИЗГА ТАКЛИФ ВА ТАЛАБ ҚИЛАМАН:\n\n1. Ушбу талабнома олинган кундан бошлаб 7 (етти) банк куни ичида юзага келган масалани қонуний ва ўзаро манфаатли асосда ихтиёрий бартараф этишингизни;\n2. Акс ҳолда ваколатли давлат органлари ва фуқаролик судига мурожаат қилиниб, барча зарарлар ва суд харажатлари Сиздан ундирилади.",
                    "attachments_list": ["Ҳолатни тасдиқловчи бирламчи ҳужжатлар нусхаси", "Шахсни тасдиқловчи ҳужжат нусхаси"]
                }
            }
        else:
            return {
                "meta": {
                    "detected_language": "uz_latn",
                    "primary_intent": "mediation",
                    "secondary_intent": "appeal",
                    "urgency_level": "o'rtacha",
                    "case_category": "Fuqarolik-huquqiy munosabatlar va Murosa"
                },
                "executive_summary": {
                    "headline": "Kiritilgan muammoli holat bo'yicha O'zR qonunchiligi asosida sudgacha rasmiy talabnoma va mediasiya strategiyasi belgilandi.",
                    "detailed_advice": f"Siz kiritgan masala («{user_facts[:120]}...»): O'zbekiston Respublikasi Fuqarolik kodeksining 9-moddasiga ko'ra, fuqarolar o'z huquqlarini sudgacha talabnoma taqdim etish va qonuniy vositalar orqali himoya qilishga to'liq haqli. Sudga chiqish vaqt (2 dan 4 oygacha) va davlat boji (da'vo summasining 4%) talab etganligi sababli, birinchi navbatda ikkinchi tarafga 7 kunlik rasmiy Talabnoma (Pretenziya) yuborish tavsiya etiladi. Ushbu hujjat sizning qonuniy talablaringizni to'liq ifodalaydi.",
                    "legal_articles": [
                        { "code_name": "O'zbekiston Respublikasi Fuqarolik kodeksi", "article": "9-modda. Fuqarolik huquqlarini himoya qilish", "relevance": "Fuqarolarga o'z buzilgan huquqlarini qonuniy vositalar orqali sudgacha himoya qilish kafolatini beradi." },
                        { "code_name": "O'zbekiston Respublikasining «Mediasiya to'g'risida»gi Qonuni", "article": "15-modda. Mediasiya kelishuvi", "relevance": "Nizolarni sudga chiqarmasdan taraflarning o'zaro kelishuvi orqali hal qilish kuchini ta'minlaydi." }
                    ]
                },
                "risk_and_comparison": {
                    "court_vs_mediation": {
                        "show_comparison": True, "court_duration": "2 – 4 oy", "court_cost_estimate": "Da'vo summasining 4% davlat boji + advokat",
                        "mediation_duration": "3 – 7 kun", "mediation_cost_estimate": "Minimal", "peaceful_resolution_chance": 85,
                        "why_mediation_better": "Sud xarajatlari va asab buzilishisiz, o'zaro manfaatli grafik asosida tinch yechim beradi."
                    },
                    "warnings": ["Barcha mavjud yozishmalar, shartnomalar, cheklar va dalillarni tartibga solib saqlang."]
                },
                "action_roadmap": [
                    { "step_number": 1, "title": "Rasmiy talabnoma topshirish", "description": "Tayyorlangan talabnomani ikkinchi tarafga topshiring.", "deadline_hint": "Bugunoq" },
                    { "step_number": 2, "title": "Mediasiya va muzokara", "description": "Kelishilgan shartlar asosida tinch bitim imzolang.", "deadline_hint": "7 kun ichida" }
                ],
                "notary_module": { "is_relevant": False, "service_title": "Notarial maslahat", "required_documents": [], "estimated_state_duty": "Belgilanmagan", "pro_tip": "Zarurat tug'ilsa tuzilgan kelishuvni notariusda tasdiqlash mumkin." },
                "official_document": {
                    "is_generated": True, "document_type": "TALABNOMA",
                    "recipient_title": "Nizo bo'yicha ikkinchi taraf: [F.I.SH. yoki Tashkilot nomi]ga\nManzil: [Manzil]",
                    "sender_placeholder": "Murojaatchi: [F.I.SH.]dan\nManzil: [Yashash manzili]\nTel: [Telefon raqami]",
                    "document_title": "NIZOLI VAZIYATNI SUDGACHA HAL ETISH VA MEDIASIYA TAKLIFI TO'G'RISIDA\nTALABNOMA (PRETENZIYA)",
                    "document_body": f"Mazkur talabnoma orqali quyidagi holat yuzasidan murojaat qilmoqdaman:\n\n«{user_facts}»\n\nO'zbekiston Respublikasi Fuqarolik kodeksining tegishli moddalariga asosan taraflar o'z zimmalaridagi majburiyatlarni lozim darajada bajarishlari shart.\n\nO'zbekiston Respublikasining «Mediasiya to'g'risida»gi Qonuniga muvofiq, yuzaga kelgan ushbu nizoli holatni sudga olib chiqmasdan, ortiqcha davlat bojlari (da'vo summasining 4 foizi) va sud xarajatlarining oldini olgan holda hal qilish maqsadida SIZGA TAKLIF VA TALAB QILAMAN:\n\n1. Ushbu talabnoma olingan kundan boshlab 7 (yetti) bank kuni ichida yuzaga kelgan masalani qonuniy va o'zaro manfaatli asosda ixtiyoriy bartaraf etishingizni;\n2. Aks holda vakolatli davlat organlari va fuqarolik sudiga murojaat qilinib, barcha zararlar va sud xarajatlari Sizdan undiriladi.",
                    "attachments_list": ["Holatni tasdiqlovchi birlamchi hujjatlar nusxasi", "Shaxsni tasdiqlovchi hujjat nusxasi"]
                }
            }


# OfflineLegalDemoEngine DynamicLegalBrain bilan almashtirildi
OfflineLegalDemoEngine = DynamicLegalBrain


class WordDocumentGenerator:
    """
    O‘zbekiston Respublikasi davlat ish yuritish (делопроизводство) standartlariga
    100% mos keluvchi Word (.docx) rasmiy blanka generatori.
    """

    @staticmethod
    def generate_docx(doc_data: Dict[str, Any], output_path: str = None, sanctions_data: Dict[str, Any] = None) -> str:
        doc = Document()

        section = doc.sections[0]
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(13)
        font.color.rgb = RGBColor(0, 0, 0)

        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        col_widths = [Mm(75), Mm(90)]
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width

        right_cell = table.cell(0, 1)

        tcPr = right_cell._element.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m in ['top', 'left', 'bottom', 'right']:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), '0')
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

        recipient_text = doc_data.get('recipient_title', '[Kimga: Tashkilot yoki Shaxs]')
        sender_text = doc_data.get('sender_placeholder', '[Kimdan: Fuqaro F.I.SH.]')

        p_header = right_cell.paragraphs[0]
        p_header.paragraph_format.line_spacing = 1.15
        p_header.paragraph_format.space_after = Pt(4)
        run_rec = p_header.add_run(recipient_text.strip() + "\n\n")
        run_rec.font.name = 'Times New Roman'
        run_rec.font.size = Pt(12)
        run_rec.bold = True

        run_snd = p_header.add_run(sender_text.strip())
        run_snd.font.name = 'Times New Roman'
        run_snd.font.size = Pt(12)

        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(18)
        p_space.paragraph_format.space_after = Pt(6)

        title_text = doc_data.get('document_title', 'RASMIY HUJJAT').strip()
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(18)
        run_title = p_title.add_run(title_text)
        run_title.font.name = 'Times New Roman'
        run_title.font.size = Pt(14)
        run_title.bold = True

        body_text = doc_data.get('document_body', '').strip()
        paragraphs = [p.strip() for p in body_text.split('\n') if p.strip()]

        for p_str in paragraphs:
            p_body = doc.add_paragraph()
            p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_body.paragraph_format.first_line_indent = Mm(12.5)
            p_body.paragraph_format.line_spacing = 1.15
            p_body.paragraph_format.space_after = Pt(6)

            run_body = p_body.add_run(p_str)
            run_body.font.name = 'Times New Roman'
            run_body.font.size = Pt(13)

        # 5. Qonuniy ogohlantirish va javobgarlik choralari (Lex.uz standarti)
        sanctions_text = ""
        if sanctions_data and isinstance(sanctions_data, dict):
            s_app = sanctions_data.get('applicable_sanctions', '').strip()
            s_comp = sanctions_data.get('compensation_rights', '').strip()
            if s_app:
                sanctions_text = f"{s_app} {s_comp}".strip()

        # Tilni aniqlash
        is_ru = bool(re.search(r'[\u0400-\u04FF]', body_text) and not any(c in body_text.lower() for c in ['қ', 'ғ', 'ҳ', 'ў', 'бўлиб']))
        is_uz_cyrl = bool(any(c in body_text.lower() for c in ['қ', 'ғ', 'ҳ', 'ў', 'бўлиб', 'керак']))

        if sanctions_text:
            p_warn_head = doc.add_paragraph()
            p_warn_head.paragraph_format.first_line_indent = Mm(12.5)
            p_warn_head.paragraph_format.space_before = Pt(14)
            p_warn_head.paragraph_format.space_after = Pt(4)

            warn_title = "ПРЕДУПРЕЖДЕНИЕ И МЕРЫ ОТВЕТСТВЕННОСТИ:" if is_ru else ("ҚОНУНИЙ ОГОҲЛАНТИРИШ ВА ЖАВОБГАРЛИК ЧОРАЛАРИ:" if is_uz_cyrl else "QONUNIY OGOHLANTIRISH VA JAVOBGARLIK CHORALARI:")
            run_wh = p_warn_head.add_run(warn_title)
            run_wh.font.name = 'Times New Roman'
            run_wh.font.size = Pt(13)
            run_wh.bold = True

            p_warn = doc.add_paragraph()
            p_warn.paragraph_format.first_line_indent = Mm(12.5)
            p_warn.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_warn.paragraph_format.line_spacing = 1.15
            p_warn.paragraph_format.space_after = Pt(10)

            if is_ru:
                warn_msg = f"В случае неисполнения указанных в настоящем документе законных требований в добровольном порядке в установленный срок, в отношении виновной стороны будут применены следующие меры ответственности: {sanctions_text}. Также в судебном порядке будут принудительно взысканы все убытки, моральный вред, неустойка и судебные расходы."
            elif is_uz_cyrl:
                warn_msg = f"Мазкур расмий ҳужжатда кўрсатилган қонуний талаблар белгиланган муддатда ихтиёрий равишда бажарилмаган тақдирда, қонунбузар тарафга нисбатан қуйидаги жавобгарлик чоралари қўлланилиши ҳақида қатъий огоҳлантирилади: {sanctions_text}. Шунингдек, етказилган барча моддий ва маънавий зарарлар ҳамда суд харажатлари айбдор тараф ҳисобидан мажбурий ундириб олинади."
            else:
                warn_msg = f"Mazkur rasmiy talabnomada ko‘rsatilgan qonuniy majburiyatlar belgilangan muddatda ixtiyoriy ravishda bajarilmagan taqdirda, qonunbuzar tarafga nisbatan quyidagi javobgarlik choralari qo‘llanilishi haqida qat'iy ogohlantiriladi: {sanctions_text}. Shuningdek, yetkazilgan barcha moddiy va ma’naviy zararlar hamda sud xarajatlari to‘liq hajmda aybdor taraf hisobidan majburiy undirib olinadi."

            run_wt = p_warn.add_run(warn_msg)
            run_wt.font.name = 'Times New Roman'
            run_wt.font.size = Pt(12)
            run_wt.italic = True

        attachments = doc_data.get('attachments_list', [])
        if attachments:
            p_att_title = doc.add_paragraph()
            p_att_title.paragraph_format.first_line_indent = Mm(12.5)
            p_att_title.paragraph_format.space_before = Pt(14)
            p_att_title.paragraph_format.space_after = Pt(4)

            att_header = "Список прилагаемых документов:" if is_ru else ("Илова қилинаётган ҳужжатлар рўйхати:" if is_uz_cyrl else "Ilova qilinayotgan hujjatlar ro‘yxati:")
            run_att_h = p_att_title.add_run(att_header)
            run_att_h.font.name = 'Times New Roman'
            run_att_h.font.size = Pt(13)
            run_att_h.bold = True

            for idx, att in enumerate(attachments, 1):
                p_att = doc.add_paragraph()
                p_att.paragraph_format.first_line_indent = Mm(18)
                p_att.paragraph_format.line_spacing = 1.15
                p_att.paragraph_format.space_after = Pt(3)
                run_item = p_att.add_run(f"{idx}. {att}")
                run_item.font.name = 'Times New Roman'
                run_item.font.size = Pt(12)

        p_sig_space = doc.add_paragraph()
        p_sig_space.paragraph_format.space_before = Pt(24)

        sig_table = doc.add_table(rows=1, cols=2)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig_table.autofit = False

        for row in sig_table.rows:
            row.cells[0].width = Mm(80)
            row.cells[1].width = Mm(85)

        date_text = "«_____» ____________ 202___ г." if is_ru else ("«_____» ____________ 202___ йил" if is_uz_cyrl else "«_____» ____________ 202___ yil")
        sig_text = "Подпись: ____________ / [Ф.И.О.]" if is_ru else ("Имзо: ____________ / [Ф.И.Ш.]" if is_uz_cyrl else "Imzo: ____________ / [F.I.SH.]")

        p_date = sig_table.cell(0, 0).paragraphs[0]
        run_date = p_date.add_run(date_text)
        run_date.font.name = 'Times New Roman'
        run_date.font.size = Pt(13)

        p_sig = sig_table.cell(0, 1).paragraphs[0]
        p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_sig = p_sig.add_run(sig_text)
        run_sig.font.name = 'Times New Roman'
        run_sig.font.size = Pt(13)

        if not output_path:
            filename = f"hujjat_{int(time.time())}.docx"
            media_docs_dir = Path(settings.MEDIA_ROOT) / 'generated_docs'
            media_docs_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(media_docs_dir / filename)

        doc.save(output_path)
        return output_path


class UniversalDocumentReader:
    """
    PDF, Rasm (JPG, JPEG, PNG, WEBP), Word (DOCX), Oddiy matn (TXT) va
    Ovozli xabarlardan (WAV, MP3, M4A, OGG) yuridik matnlarni o'qib oluvchi Universal Tizim.
    """
    _ocr_reader = None
    _whisper_model = None

    @classmethod
    def get_ocr_reader(cls):
        if cls._ocr_reader is None:
            try:
                import easyocr
                cls._ocr_reader = easyocr.Reader(['ru', 'en'], gpu=False)
            except Exception as e:
                cls._ocr_reader = False
        return cls._ocr_reader if cls._ocr_reader is not False else None

    @classmethod
    def get_whisper_model(cls):
        if cls._whisper_model is None:
            try:
                import whisper
                cls._whisper_model = whisper.load_model('base')
            except Exception as e:
                cls._whisper_model = False
        return cls._whisper_model if cls._whisper_model is not False else None

    @classmethod
    def extract_text_from_file(cls, file_bytes: bytes, filename: str) -> Dict[str, Any]:
        ext = os.path.splitext(filename)[1].lower()
        size_kb = round(len(file_bytes) / 1024, 1)

        result = {
            "success": False,
            "filename": filename,
            "ext": ext,
            "size_kb": size_kb,
            "file_type": "unknown",
            "extracted_text": "",
            "char_count": 0,
            "message": ""
        }

        try:
            # 1. PDF HUJJATLAR (PyMuPDF + Skaner OCR zaxirasi)
            if ext == '.pdf':
                result["file_type"] = "pdf"
                import fitz
                pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
                pages_text = []
                for p_idx, page in enumerate(pdf_doc):
                    t = page.get_text().strip()
                    if t:
                        pages_text.append(f"[{p_idx + 1}-sahifa]:\n{t}")
                extracted = "\n\n".join(pages_text).strip()

                # Agar skanerlangan PDF bo'lsa (matn kam yoki yo'q bo'lsa), EasyOCR orqali sahifalarni o'qiymiz
                if len(extracted) < 40 and len(pdf_doc) > 0:
                    ocr = cls.get_ocr_reader()
                    if ocr:
                        ocr_pages = []
                        for p_idx, page in enumerate(pdf_doc[:3]):
                            pix = page.get_pixmap(dpi=150)
                            img_bytes = pix.tobytes("png")
                            lines = ocr.readtext(img_bytes, detail=0)
                            if lines:
                                ocr_pages.append(f"[{p_idx + 1}-sahifa (Skaner OCR)]:\n" + "\n".join(lines))
                        if ocr_pages:
                            extracted = "\n\n".join(ocr_pages).strip()

                result["extracted_text"] = extracted
                result["char_count"] = len(extracted)
                result["success"] = bool(extracted)
                result["message"] = f"PDF hujjatdan {len(extracted)} belgi o'qildi" if extracted else "PDF faylda matn aniqlanmadi"

            # 2. RASMLAR (JPG, JPEG, PNG, WEBP, BMP - VoiceLab Vision + EasyOCR zaxirasi)
            elif ext in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
                result["file_type"] = "image"
                extracted = ""

                # VoiceLab Vision (Aisha-Halo) orqali yuridik hujjatni o'qish
                api_key = getattr(settings, 'VOICELAB_API_KEY', '')
                if api_key:
                    try:
                        import base64
                        b64_img = base64.b64encode(file_bytes).decode('utf-8')
                        mime = "image/jpeg" if ext in ['.jpg', '.jpeg'] else "image/png"
                        headers = {
                            "Authorization": f"Bearer {api_key}",
                            "Content-Type": "application/json"
                        }
                        payload = {
                            "model": "aisha-halo",
                            "messages": [
                                {
                                    "role": "user",
                                    "content": [
                                        {
                                            "type": "text",
                                            "text": "Ushbu yuridik hujjat (tilxat, sud qarori, shartnoma, chek yoki ariza) rasmida yozilgan barcha matnni, sanalarni, taraflar ism-shariflarini va summalarni aynan o'qib, tartibli matn ko'rinishida chiqarib ber. Faqat hujjatdagi haqiqiy matnni yoz."
                                        },
                                        {
                                            "type": "image_url",
                                            "image_url": {"url": f"data:{mime};base64,{b64_img}"}
                                        }
                                    ]
                                }
                            ],
                            "max_tokens": 4096,
                            "temperature": 0.1
                        }
                        r = requests.post("https://api.voicelab.uz/v1/chat/completions", headers=headers, json=payload, timeout=25)
                        if r.status_code == 200:
                            j = r.json()
                            extracted = j['choices'][0]['message'].get('content', '').strip()
                    except Exception as vision_err:
                        pass

                # Zaxira: EasyOCR
                if not extracted:
                    ocr = cls.get_ocr_reader()
                    if ocr:
                        lines = ocr.readtext(file_bytes, detail=0)
                        extracted = "\n".join(lines).strip()

                result["extracted_text"] = extracted
                result["char_count"] = len(extracted)
                result["success"] = bool(extracted)
                result["message"] = f"Rasm / Tilxatdan {len(extracted)} belgi o'qildi" if extracted else "Rasmda matn aniqlanmadi"

            # 3. WORD (.DOCX) HUJJATLAR
            elif ext == '.docx':
                result["file_type"] = "docx"
                doc = Document(io.BytesIO(file_bytes))
                lines = [p.text for p in doc.paragraphs if p.text.strip()]
                for t in doc.tables:
                    for row in t.rows:
                        row_vals = [c.text.strip() for c in row.cells if c.text.strip()]
                        if row_vals:
                            lines.append(" | ".join(row_vals))
                extracted = "\n".join(lines).strip()
                result["extracted_text"] = extracted
                result["char_count"] = len(extracted)
                result["success"] = bool(extracted)
                result["message"] = f"Word hujjatdan {len(extracted)} belgi o'qildi"

            # 4. OVOZ VA AUDIO (.WAV, .MP3, .OGG, .M4A)
            elif ext in ['.wav', '.mp3', '.ogg', '.m4a']:
                result["file_type"] = "audio"
                whisper_model = cls.get_whisper_model()
                if whisper_model:
                    import tempfile
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
                        tf.write(file_bytes)
                        tf_path = tf.name
                    try:
                        transcription = whisper_model.transcribe(tf_path)
                        extracted = transcription.get('text', '').strip()
                    finally:
                        if os.path.exists(tf_path):
                            os.remove(tf_path)
                result["extracted_text"] = extracted
                result["char_count"] = len(extracted)
                result["success"] = bool(extracted)
                result["message"] = f"Ovozli murojaat matnga aylantirildi ({len(extracted)} belgi)"

            # 5. ODDIY MATN (.TXT)
            elif ext in ['.txt', '.rtf']:
                result["file_type"] = "text"
                for enc in ['utf-8', 'cp1251', 'latin-1']:
                    try:
                        extracted = file_bytes.decode(enc).strip()
                        break
                    except UnicodeDecodeError:
                        continue
                result["extracted_text"] = extracted
                result["char_count"] = len(extracted)
                result["success"] = bool(extracted)
                result["message"] = f"Matn faylidan {len(extracted)} belgi o'qildi"

            else:
                result["message"] = f"Qo'llab-quvvatlanmaydigan format: {ext}"

        except Exception as e:
            result["message"] = f"Hujjatni o'qishda xatolik: {str(e)}"

        return result


class LegalAIService:
    """
    ADOLAT.AI Tizimining Bosh Xizmat Koordinatori:
    VoiceLab Aisha Orbit / Qwen LLM / Dynamic Legal Brain / python-docx integratsiyasi.
    Har qanday yangi savolga 100% professional huquqiy yechim beradi!
    """

    @classmethod
    def process_case(cls, query_text: str, forced_lang: str = None, document_text: str = None) -> Tuple[Dict[str, Any], str, float]:
        start_time = time.time()
        parsed_data = None

        if document_text and document_text.strip():
            combined_query = f"""[FUQARONING HUQUQIY SAVOLI / MUAMMOSI]:
{query_text.strip() if query_text else "Iltimos, ilova qilingan rasmiy hujjatni to'liq tahlil qilib, huquqiy xulosa va rasmiy davlat hujjati loyihasini tayyorlang."}

[ILOVADAGI RASMIY HUJJAT (PDF / RASM / TILXAT / SHARTNOMA / OVOZ) MATNI]:
{document_text.strip()}"""
        else:
            combined_query = query_text

        messages = build_llm_messages(combined_query, forced_lang)

        # 1. VoiceLab Aisha Orbit (Pulli va yuqori darajali O'zbekiston AI modeli)
        voicelab_client = VoiceLabClient()
        if voicelab_client.is_available():
            try:
                raw_response = voicelab_client.generate(messages)
                parsed_data = JSONSanitizer.clean_and_parse(raw_response)
            except Exception as vl_err:
                # print(f"[VoiceLab Error / Fallback]: {vl_err}")
                parsed_data = None

        # 2. Agar VoiceLab ishlamasa, ikkinchi darajali Qwen/OpenRouter mijozi
        if not parsed_data or not isinstance(parsed_data, dict) or 'executive_summary' not in parsed_data:
            qwen_client = QwenRESTClient()
            if qwen_client.is_available() and qwen_client.api_url != voicelab_client.api_url:
                try:
                    raw_response = qwen_client.generate(messages)
                    parsed_data = JSONSanitizer.clean_and_parse(raw_response)
                except Exception:
                    parsed_data = None

        # 3. KUCHLI VA DINAMIK LEGAL BRAIN (Offline kafolatlangan dvijok)
        if not parsed_data or not isinstance(parsed_data, dict) or 'executive_summary' not in parsed_data:
            parsed_data = DynamicLegalBrain.analyze_custom_query(query_text, forced_lang)
        else:
            # Agar LLM javob bergan bo'lsa-yu, lekin sanctions_and_liability yetarli bo'lmasa:
            if 'legal_sanctions_and_liability' not in parsed_data or not parsed_data['legal_sanctions_and_liability']:
                lang = parsed_data.get('meta', {}).get('detected_language', 'uz')
                parsed_data['legal_sanctions_and_liability'] = DynamicLegalBrain.get_sanctions_and_liability('universal', lang)

        # 4. Rasmiy Word (.docx) hujjatini yaratish
        official_doc = parsed_data.get('official_document', {})
        docx_path = ""
        sanctions_data = parsed_data.get('legal_sanctions_and_liability', {})
        if official_doc.get('is_generated', True):
            try:
                docx_path = WordDocumentGenerator.generate_docx(official_doc, sanctions_data=sanctions_data)
            except Exception as docx_err:
                print(f"[Docx Gen Error]: {docx_err}")

        execution_time = round(time.time() - start_time, 2)
        return parsed_data, docx_path, execution_time

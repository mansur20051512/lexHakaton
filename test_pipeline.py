import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
django.setup()

from legal.services import LegalAIService
from django.test import Client

def main():
    print(">>> 1. Testing LegalAIService (Mediation)...")
    res, docx_path, exec_time = LegalAIService.process_case("Qo'shnim 15 mln so'm qarz olib qaytarmayapti", "uz")
    assert res['meta']['primary_intent'] == 'mediation'
    assert os.path.exists(docx_path), "Docx file was not created"
    print(f"    Language: {res['meta']['detected_language']}, Intent: {res['meta']['primary_intent']}, Docx: {docx_path}")

    print(">>> 2. Testing API Analyze View (/api/analyze/)...")
    c = Client()
    payload = {"query": "Qarz undirish bo'yicha yordam bering", "lang": "uz"}
    r = c.post('/api/analyze/', payload, content_type='application/json')
    assert r.status_code == 200, f"Status code {r.status_code}"
    res_json = r.json()
    case_id = res_json['case_id']
    print(f"    Case created with ID: {case_id}, Exec time: {res_json['execution_time']}s")

    print(">>> 3. Testing Docx Download View (/download-docx/<id>/)...")
    r_doc = c.get(f'/download-docx/{case_id}/')
    assert r_doc.status_code == 200, f"Status code {r_doc.status_code}"
    file_bytes = b"".join(r_doc.streaming_content) if hasattr(r_doc, 'streaming_content') else r_doc.getvalue()
    assert len(file_bytes) > 1000, f"File size too small: {len(file_bytes)}"
    print(f"    Docx downloaded successfully! Byte length: {len(file_bytes)}")

    print(">>> 4. Testing Russian Language Input...")
    res_ru, docx_ru, _ = LegalAIService.process_case("Сосед не возвращает долг 15 миллионов сумов", "ru")
    assert res_ru['meta']['detected_language'] == 'ru'
    assert "ПРЕТЕНЗИЯ" in res_ru['official_document']['document_title']
    print("    Russian input correctly returned Russian response!")

    print("\n✅ ALL BACKEND & SERVICE TESTS PASSED 100% SUCCESSFULLY!")

if __name__ == '__main__':
    main()

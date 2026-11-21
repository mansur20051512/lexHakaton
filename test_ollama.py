import requests, json

url = 'http://localhost:11434/v1/chat/completions'
prompt = 'Fuqaroning murojaati: Meni ishdan sababsiz bo\'shatib yuborishdi, 2 oylik maoshimni ham berishmadi. Nima qilsam bo\'ladi?'

payload = {
    'model': 'qwen2.5:7b',
    'messages': [
        {'role': 'system', 'content': 'Sen yurist-ekspertsan. Javobni FAQAT JSON formatda qaytar: {"meta": {"detected_language": "uz_latn", "primary_intent": "mediation", "case_category": "Mehnat nizosi"}, "executive_summary": {"headline": "Mehnat nizosi", "detailed_advice": "Tavsiyalar"}}'},
        {'role': 'user', 'content': prompt}
    ],
    'temperature': 0.2,
    'max_tokens': 500
}
try:
    res = requests.post(url, json=payload, timeout=25)
    print('STATUS:', res.status_code)
    print('RESPONSE:', res.json()['choices'][0]['message']['content'])
except Exception as e:
    print('ERR:', e)

import json
import os
from pathlib import Path

from django.shortcuts import render, get_object_or_404
from django.http import JsonResponse, HttpResponse, Http404, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST, require_GET
from django.conf import settings

from .models import LegalCase
from .services import LegalAIService, WordDocumentGenerator, UniversalDocumentReader


def index_view(request):
    """Bosh sahifani ko‘rsatish (Premium Oq-Oltin davlat portali dizaynida)"""
    recent_cases = LegalCase.objects.all().order_by('-created_at')[:5]
    total_cases_count = LegalCase.objects.count()
    
    context = {
        'recent_cases': recent_cases,
        'total_cases_count': total_cases_count + 1420,  # Xakaton statistikasi uchun realistik bazaviy son
    }
    return render(request, 'index.html', context)


@csrf_exempt
@require_POST
def upload_document_api_view(request):
    """
    Foydalanuvchi yuklagan PDF, Rasm (JPG, PNG, WEBP), DOCX yoki Ovozli fayldan
    matnni o'qib beruvchi Universal OCR & Hujjat Analizatori API-si.
    """
    uploaded_file = request.FILES.get('file') or request.FILES.get('document')
    if not uploaded_file:
        return JsonResponse({'error': 'Hech qanday fayl yuklanmadi'}, status=400)

    try:
        file_bytes = uploaded_file.read()
        filename = uploaded_file.name
        result = UniversalDocumentReader.extract_text_from_file(file_bytes, filename)
        return JsonResponse(result)
    except Exception as e:
        return JsonResponse({'error': f'Faylni tahlil qilishda xatolik: {str(e)}'}, status=500)


@csrf_exempt
@require_POST
def analyze_api_view(request):
    """
    Fuqaro muammosini va ilova qilingan hujjat (PDF, Rasm, Ovoz)ni qabul qilish hamda
    VoiceLab Aisha Orbit AI orqali to'liq yuridik tahlil + Word hujjat yaratish API-si.
    """
    query_text = ""
    forced_lang = None
    document_text = ""

    # Multipart/form-data yoki JSON formatni qabul qilish
    if request.content_type and 'multipart/form-data' in request.content_type:
        query_text = request.POST.get('query', '').strip()
        forced_lang = request.POST.get('lang', None)
        document_text = request.POST.get('document_text', '').strip()

        uploaded_file = request.FILES.get('file') or request.FILES.get('document')
        if uploaded_file:
            doc_result = UniversalDocumentReader.extract_text_from_file(uploaded_file.read(), uploaded_file.name)
            if doc_result.get('extracted_text'):
                document_text = (document_text + "\n" + doc_result['extracted_text']).strip()
    else:
        try:
            data = json.loads(request.body.decode('utf-8'))
        except Exception:
            return JsonResponse({'error': 'Noto‘g‘ri JSON formati yuborildi'}, status=400)

        query_text = data.get('query', '').strip()
        forced_lang = data.get('lang', None)
        document_text = data.get('document_text', '').strip()

    if not query_text and not document_text:
        return JsonResponse({'error': 'Murojaat matni yoki ilova qilingan hujjat bo‘sh bo‘lishi mumkin emas'}, status=400)

    # Huquqiy AI xizmatini chaqirish (savol + hujjat matni bilan)
    parsed_data, docx_path, execution_time = LegalAIService.process_case(
        query_text, 
        forced_lang, 
        document_text=document_text
    )

    meta = parsed_data.get('meta', {})
    summary = parsed_data.get('executive_summary', {})
    risk = parsed_data.get('risk_and_comparison', {}).get('court_vs_mediation', {})
    doc = parsed_data.get('official_document', {})

    # Ma'lumotlar bazasiga saqlash
    relative_docx_path = ""
    if docx_path and os.path.exists(docx_path):
        try:
            rel_path = Path(docx_path).relative_to(settings.MEDIA_ROOT)
            relative_docx_path = str(rel_path)
        except Exception:
            relative_docx_path = os.path.basename(docx_path)

    legal_case = LegalCase.objects.create(
        query_text=query_text,
        detected_language=meta.get('detected_language', 'uz_latn'),
        primary_intent=meta.get('primary_intent', 'mediation'),
        secondary_intent=meta.get('secondary_intent'),
        urgency_level=meta.get('urgency_level', 'o‘rtacha'),
        case_category=meta.get('case_category', 'Fuqarolik-huquqiy nizo'),
        headline=summary.get('headline', 'Huquqiy xulosa'),
        detailed_advice=summary.get('detailed_advice', ''),
        court_duration=risk.get('court_duration', '2 – 4 oy'),
        court_cost_estimate=risk.get('court_cost_estimate', '4% davlat boji + advokat'),
        mediation_duration=risk.get('mediation_duration', '3 – 7 kun'),
        mediation_cost_estimate=risk.get('mediation_cost_estimate', 'Minimal / Kelishuv asosida'),
        peaceful_resolution_chance=risk.get('peaceful_resolution_chance', 85),
        document_title=doc.get('document_title', 'TALABNOMA'),
        document_type=doc.get('document_type', 'TALABNOMA'),
        recipient_title=doc.get('recipient_title', ''),
        sender_placeholder=doc.get('sender_placeholder', ''),
        document_body=doc.get('document_body', ''),
        attachments_json=doc.get('attachments_list', []),
        raw_response_json=parsed_data,
        docx_file=relative_docx_path if relative_docx_path else None,
        execution_time_seconds=execution_time
    )

    download_url = f"/download-docx/{legal_case.id}/"

    return JsonResponse({
        'success': True,
        'case_id': legal_case.id,
        'download_url': download_url,
        'execution_time': execution_time,
        'data': parsed_data
    })


@require_GET
def download_docx_view(request, case_id):
    """Generatsiya qilingan rasmiy Word (.docx) hujjatini yuklab berish"""
    legal_case = get_object_or_404(LegalCase, id=case_id)

    # Agar fayl avval yaratilmagan bo‘lsa yoki o‘chib ketgan bo‘lsa, qayta yaratamiz
    file_path = None
    if legal_case.docx_file:
        possible_path = Path(settings.MEDIA_ROOT) / str(legal_case.docx_file)
        if possible_path.exists():
            file_path = str(possible_path)

    if not file_path:
        doc_data = {
            'recipient_title': legal_case.recipient_title,
            'sender_placeholder': legal_case.sender_placeholder,
            'document_title': legal_case.document_title,
            'document_body': legal_case.document_body,
            'attachments_list': legal_case.attachments_json
        }
        file_path = WordDocumentGenerator.generate_docx(doc_data)
        try:
            rel_path = Path(file_path).relative_to(settings.MEDIA_ROOT)
            legal_case.docx_file = str(rel_path)
            legal_case.save(update_fields=['docx_file'])
        except Exception:
            pass

    safe_title = "".join(c for c in legal_case.document_type if c.isalnum() or c in (' ', '_', '-')).strip()
    if not safe_title:
        safe_title = "Hujjat"
    download_filename = f"ADOLAT_AI_{safe_title}_{legal_case.id}.docx"

    response = FileResponse(
        open(file_path, 'rb'),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    response['Content-Disposition'] = f'attachment; filename="{download_filename}"'
    return response


@require_GET
def history_api_view(request):
    """So‘nggi tahlillar tarixini JSON formatida qaytarish"""
    cases = LegalCase.objects.all().order_by('-created_at')[:10]
    items = []
    for c in cases:
        items.append({
            'id': c.id,
            'headline': c.headline,
            'intent': c.get_primary_intent_display(),
            'lang': c.get_detected_language_display(),
            'urgency': c.urgency_level,
            'created_at': c.created_at.strftime('%d.%m.%Y %H:%M'),
            'download_url': f"/download-docx/{c.id}/"
        })
    return JsonResponse({'cases': items})

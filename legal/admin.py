from django.contrib import admin
from .models import LegalCase

@admin.register(LegalCase)
class LegalCaseAdmin(admin.ModelAdmin):
    list_display = ('id', 'headline', 'primary_intent', 'detected_language', 'urgency_level', 'execution_time_seconds', 'created_at')
    list_filter = ('primary_intent', 'detected_language', 'urgency_level', 'created_at')
    search_fields = ('query_text', 'headline', 'document_title', 'document_body')
    readonly_fields = ('created_at', 'execution_time_seconds')

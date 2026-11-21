from django.urls import path
from . import views

urlpatterns = [
    path('', views.index_view, name='index'),
    path('api/analyze/', views.analyze_api_view, name='api_analyze'),
    path('api/upload-document/', views.upload_document_api_view, name='api_upload_document'),
    path('download-docx/<int:case_id>/', views.download_docx_view, name='download_docx'),
    path('api/history/', views.history_api_view, name='api_history'),
]

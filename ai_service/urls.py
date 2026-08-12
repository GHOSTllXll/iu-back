# backend/ai_service/urls.py
from django.urls import path
from .views import (
    TestAIView,
    DocumentUploadView,
    UnderwritePropertyView,
    UnderwritePropertyDownloadView,
    DocumentHistoryView,
    ProcessedDocumentDeleteView,
    AnalysisReportListView,
    AnalysisReportDetailView,
    DashboardStatsView,
)

urlpatterns = [
    path('test/', TestAIView.as_view(), name='test_ai'),
    path('upload/', DocumentUploadView.as_view(), name='upload_documents'),
    path('underwrite/', UnderwritePropertyView.as_view(), name='underwrite_property'),
    path('underwrite/download/', UnderwritePropertyDownloadView.as_view(), name='underwrite_download'),
    path('documents/', DocumentHistoryView.as_view(), name='document_history'),
    path('documents/<int:document_id>/', ProcessedDocumentDeleteView.as_view(), name='document_delete'),  # NEW
    path('reports/', AnalysisReportListView.as_view(), name='report_list'),
    path('reports/<int:report_id>/', AnalysisReportDetailView.as_view(), name='report_detail'),  # now handles GET + DELETE
    path('dashboard-stats/', DashboardStatsView.as_view(), name='dashboard_stats'),
]
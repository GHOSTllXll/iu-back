# backend/ai_service/urls.py
from django.urls import path
from .views import (
    TestAIView,
    DocumentUploadView,
    UnderwritePropertyView,
    UnderwritePropertyDownloadView,
    DocumentHistoryView,
    AnalysisReportListView,
    AnalysisReportDetailView,
    DashboardStatsView,
)

urlpatterns = [
    path('test/', TestAIView.as_view(), name='test_ai'),
    path('upload/', DocumentUploadView.as_view(), name='upload_documents'),
    path('underwrite/', UnderwritePropertyView.as_view(), name='underwrite_property'),  # Returns JSON
    path('underwrite/download/', UnderwritePropertyDownloadView.as_view(), name='underwrite_download'),  # Returns Excel
    path('documents/', DocumentHistoryView.as_view(), name='document_history'),  # Documents page: stats + list
    path('reports/', AnalysisReportListView.as_view(), name='report_list'),  # Outputs page: card list
    path('reports/<int:report_id>/', AnalysisReportDetailView.as_view(), name='report_detail'),  # Outputs page: Preview
    path('dashboard-stats/', DashboardStatsView.as_view(), name='dashboard_stats'),  # Dashboard: 4 live stat cards + quota
]
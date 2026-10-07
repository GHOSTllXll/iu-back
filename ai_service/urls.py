# backend/ai_service/urls.py
from django.urls import path
from .views import (
    TestAIView,
    DocumentUploadView,
    ClassifyDocumentView,
    UnderwritePropertyView,
    UnderwritePropertyDownloadView,
    DocumentHistoryView,
    ProcessedDocumentDeleteView,
    AnalysisReportListView,
    AnalysisReportDetailView,
    DashboardStatsView,
    PPMUnderwriteView,
    ExportPPMMasterGridView,
    TaskStatusView,
)

urlpatterns = [
    path('test/', TestAIView.as_view(), name='test_ai'),
    path('upload/', DocumentUploadView.as_view(), name='upload_documents'),

    # Intelligent Document Classification Layer — the frontend's single
    # mixed-file drop zone calls this once per dropped file to decide which
    # of the 3 slots below (om_file/t12_file/rent_roll_file) it belongs in.
    # See ClassifyDocumentView's docstring.
    path('classify-document/', ClassifyDocumentView.as_view(), name='classify_document'),

    path('underwrite/', UnderwritePropertyView.as_view(), name='underwrite_property'),
    path('underwrite/download/', UnderwritePropertyDownloadView.as_view(), name='underwrite_download'),
    path('documents/', DocumentHistoryView.as_view(), name='document_history'),
    path('documents/<int:document_id>/', ProcessedDocumentDeleteView.as_view(), name='document_delete'),  # NEW
    path('reports/', AnalysisReportListView.as_view(), name='report_list'),
    path('reports/<int:report_id>/', AnalysisReportDetailView.as_view(), name='report_detail'),  # now handles GET + DELETE
    path('dashboard-stats/', DashboardStatsView.as_view(), name='dashboard_stats'),

    # Institutional PPM pipeline — single-document-per-request; see
    # PPMUnderwriteView docstring for why the batch loop lives client-side.
    path('ppm/underwrite/', PPMUnderwriteView.as_view(), name='ppm_underwrite'),
    path('ppm/export-master-grid/', ExportPPMMasterGridView.as_view(), name='ppm_export_master_grid'),

    # Shared polling endpoint for the Celery-backed CRE (underwrite/) and
    # PPM (ppm/underwrite/) tasks above - see TaskStatusView's docstring.
    path('task-status/<str:task_id>/', TaskStatusView.as_view(), name='task_status'),
]
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
    PropertyProjectListView,
    PropertyProjectDetailView,
    TeamNoteListCreateView,
    PropertyProjectDownloadView,
    ExcelOverrideCheckView,
    ExcelOverrideHistoryView,
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

    # Feature 3 (Enterprise Pipeline & Workflow Dashboard): the Deal
    # Pipeline. See permissions.py for the visibility rule shared by all
    # four of these, and views.py's own comment block above
    # PropertyProjectListView for the full writeup.
    path('pipeline/projects/', PropertyProjectListView.as_view(), name='pipeline_project_list'),
    path('pipeline/projects/<int:project_id>/', PropertyProjectDetailView.as_view(), name='pipeline_project_detail'),
    path('pipeline/projects/<int:project_id>/notes/', TeamNoteListCreateView.as_view(), name='pipeline_project_notes'),
    path('pipeline/projects/<int:project_id>/download/', PropertyProjectDownloadView.as_view(), name='pipeline_project_download'),

    # Feature 4 (Cell Override Audit Log): "Re-upload Edited Underwriting
    # File" — see views.py's comment block above ExcelOverrideCheckView.
    path('pipeline/projects/<int:project_id>/override-check/', ExcelOverrideCheckView.as_view(), name='pipeline_project_override_check'),
    path('pipeline/projects/<int:project_id>/overrides/', ExcelOverrideHistoryView.as_view(), name='pipeline_project_overrides'),
]
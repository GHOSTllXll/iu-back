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
    ApproveLockBaselineView,
    VPSignOffView,
    ICSummaryPDFView,
    PresentDealDataView,
    TeamNoteListCreateView,
    PropertyProjectDownloadView,
    ExcelOverrideCheckView,
    ExcelOverrideHistoryView,
    OverrideAlertFeedView,
    LeadershipDigestView,
    LeadershipDigestPDFView,
    DealChecklistListCreateView,
    DealChecklistItemResolveView,
    DealVaultListView,
    DealVaultFileUploadView,
    DealVaultFileDownloadView,
    DealAssignableAnalystsView,
    LenderShareLinkListCreateView,
    LenderShareLinkRevokeView,
    LenderSharePortalView,
    LenderShareExcelDownloadView,
    LenderShareVaultFileDownloadView,
    GridFieldValuesView,
    GridEditSaveView,
    WaterfallScenarioView,
    WaterfallScenarioCalculateView,
    MarketCompsView,
    WorkspaceBuyBoxView,
    NotificationListView,
    NotificationMarkReadView,
    TeamScorecardView,
    EmailConnectionListView,
    EmailConnectStartView,
    EmailOAuthCallbackView,
    EmailConnectionDisconnectView,
    EmailThreadSearchView,
    ProjectEmailThreadsView,
    ProjectEmailThreadUnlinkView,
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

    # Feature 2 (Immutable Snapshots): "Approve & Lock Baseline" — see
    # views.py's comment block above ApproveLockBaselineView.
    path('pipeline/projects/<int:project_id>/lock/', ApproveLockBaselineView.as_view(), name='pipeline_project_lock'),

    # Pipeline Modules Phase 2 (VP Sign-Off Gate & Leadership Digest): the
    # combined "Sign Off & Advance to Offer Submitted" action — see views.py's
    # comment block above VPSignOffView.
    path('pipeline/projects/<int:project_id>/vp-sign-off/', VPSignOffView.as_view(), name='pipeline_project_vp_sign_off'),

    # Pipeline Modules Phase 2 follow-up: per-deal "Download IC Summary" —
    # see views.py's comment block above ICSummaryPDFView.
    path('pipeline/projects/<int:project_id>/ic-summary/pdf/', ICSummaryPDFView.as_view(), name='pipeline_project_ic_summary_pdf'),

    # V2 Module 4c (Boardroom IC Presentation Mode) — see views.py's
    # comment block above PresentDealDataView. A sibling of
    # ic-summary/pdf/ above: same locked-baseline data, JSON instead of
    # a PDF, for the frontend's full-screen slide view.
    path(
        'pipeline/projects/<int:project_id>/present-deal/',
        PresentDealDataView.as_view(), name='pipeline_project_present_deal',
    ),

    path('pipeline/projects/<int:project_id>/notes/', TeamNoteListCreateView.as_view(), name='pipeline_project_notes'),
    path('pipeline/projects/<int:project_id>/download/', PropertyProjectDownloadView.as_view(), name='pipeline_project_download'),

    # Feature 4 (Cell Override Audit Log): "Re-upload Edited Underwriting
    # File" — see views.py's comment block above ExcelOverrideCheckView.
    path('pipeline/projects/<int:project_id>/override-check/', ExcelOverrideCheckView.as_view(), name='pipeline_project_override_check'),
    path('pipeline/projects/<int:project_id>/overrides/', ExcelOverrideHistoryView.as_view(), name='pipeline_project_overrides'),

    # Pipeline Modules Phase 8 (Live In-Browser Spreadsheet Editor) — see
    # views.py's Phase 8 comment block above GridFieldValuesView/
    # GridEditSaveView. Deliberately siblings of override-check/overrides
    # above, not nested under them — this is a different entry point into
    # the SAME underlying diff/vault pipeline, not a sub-resource of it.
    path('pipeline/projects/<int:project_id>/grid-values/', GridFieldValuesView.as_view(), name='pipeline_project_grid_values'),
    path('pipeline/projects/<int:project_id>/grid-save/', GridEditSaveView.as_view(), name='pipeline_project_grid_save'),

    # V2 Module 3 (Live Debt & Equity Waterfall Scenario Builder, Phase
    # V2-1) — see views.py's comment block above WaterfallScenarioView/
    # WaterfallScenarioCalculateView.
    path(
        'pipeline/projects/<int:project_id>/waterfall-scenario/',
        WaterfallScenarioView.as_view(), name='pipeline_project_waterfall_scenario',
    ),
    path(
        'pipeline/projects/<int:project_id>/waterfall-calculate/',
        WaterfallScenarioCalculateView.as_view(), name='pipeline_project_waterfall_calculate',
    ),

    # V2 Module 2 (Tax & Rent Comps Panel via RentCast) — see views.py's
    # comment block above MarketCompsView. Deliberately a sibling of
    # waterfall-scenario/ above, not nested under override-check/ — this
    # is an entirely separate data source (an external vendor), not part
    # of the Excel override/audit pipeline.
    path(
        'pipeline/projects/<int:project_id>/market-comps/',
        MarketCompsView.as_view(), name='pipeline_project_market_comps',
    ),

    # V2 Module 4a (Buy-Box Auto-Alerts) — see views.py's comment block
    # above WorkspaceBuyBoxView/Notification*View. Workspace-scoped (not
    # project-scoped), deliberately spelled out under its own
    # pipeline/workspaces/ namespace rather than nested under
    # pipeline/projects/<id>/ — a buy-box is a property of the WORKSPACE,
    # evaluated against deals, not a property of any one deal.
    path(
        'pipeline/workspaces/<int:workspace_id>/buy-box/',
        WorkspaceBuyBoxView.as_view(), name='pipeline_workspace_buy_box',
    ),
    path('notifications/', NotificationListView.as_view(), name='notification_list'),
    path('notifications/mark-read/', NotificationMarkReadView.as_view(), name='notification_mark_read'),

    # Feature 1 (Variance & Change Ledger) extension: firm-wide proactive
    # alert feed (GPR/OpEx/NOI, >2.0% variance) — see views.py's comment
    # block above OverrideAlertFeedView. Not project-scoped — intentionally
    # above the override-check/overrides list so it's clear this one spans
    # every visible deal at once.
    path('pipeline/override-alerts/', OverrideAlertFeedView.as_view(), name='pipeline_override_alerts'),

    # Pipeline Modules Phase 2: the Leadership Morning Digest — see views.py's
    # comment block above LeadershipDigestView/LeadershipDigestPDFView.
    path('pipeline/leadership-digest/', LeadershipDigestView.as_view(), name='pipeline_leadership_digest'),
    path('pipeline/leadership-digest/pdf/', LeadershipDigestPDFView.as_view(), name='pipeline_leadership_digest_pdf'),

    # V2 Module 4b (Analyst Speed & Accuracy Scorecard) — see views.py's
    # comment block above TeamScorecardView. A sibling of
    # leadership-digest/ above, not nested under it — the frontend's new
    # "Team Performance" tab calls this directly rather than folding it
    # into LeadershipDigestView's own response shape.
    path('pipeline/team-scorecard/', TeamScorecardView.as_view(), name='pipeline_team_scorecard'),

    # Pipeline Modules Phase 4 (Persistent Missing-Doc Checklist) — see
    # views.py's comment block above DealChecklistListCreateView/
    # DealChecklistItemResolveView. (The unified Activity Feed itself has
    # no endpoint of its own — it's inlined into
    # pipeline_project_detail's GET response; see _get_deal_activity_feed
    # in views.py.)
    path('pipeline/projects/<int:project_id>/checklist/', DealChecklistListCreateView.as_view(), name='pipeline_project_checklist'),
    path(
        'pipeline/projects/<int:project_id>/checklist/<int:item_id>/resolve/',
        DealChecklistItemResolveView.as_view(), name='pipeline_project_checklist_resolve',
    ),

    # Pipeline Modules Phase 5 (Excel Versioning & Supporting Document
    # Vault) — see views.py's comment block above DealVaultListView/
    # DealVaultFileUploadView/DealVaultFileDownloadView. Auto-captured
    # versions (lock, re-upload) have no endpoint of their own — they're
    # written by _create_baseline_lock/ExcelOverrideCheckView directly —
    # only the manual upload and the two read paths are routed here.
    path('pipeline/projects/<int:project_id>/vault/', DealVaultListView.as_view(), name='pipeline_project_vault'),
    path('pipeline/projects/<int:project_id>/vault/upload/', DealVaultFileUploadView.as_view(), name='pipeline_project_vault_upload'),
    path(
        'pipeline/projects/<int:project_id>/vault/<int:file_id>/download/',
        DealVaultFileDownloadView.as_view(), name='pipeline_project_vault_download',
    ),

    # Pipeline Modules Phase 6 add-on: the deal drawer's assigned_analyst
    # picker — see views.py's comment block above DealAssignableAnalystsView
    # for why this is its own narrowly-scoped endpoint rather than reusing
    # MyOrganizationMembersView (admin-only).
    path(
        'pipeline/projects/<int:project_id>/assignable-analysts/',
        DealAssignableAnalystsView.as_view(), name='pipeline_project_assignable_analysts',
    ),

    # Pipeline Modules Phase 6 (One-Click Lender Submission Packet & View
    # Tracking) — see views.py's comment block above
    # LenderShareLinkListCreateView/LenderShareLinkRevokeView for the
    # authenticated, team-side half of this feature.
    path(
        'pipeline/projects/<int:project_id>/lender-links/',
        LenderShareLinkListCreateView.as_view(), name='pipeline_project_lender_links',
    ),
    path(
        'pipeline/projects/<int:project_id>/lender-links/<int:link_id>/',
        LenderShareLinkRevokeView.as_view(), name='pipeline_project_lender_link_revoke',
    ),

    # Pipeline Modules Phase 6: the PUBLIC, unauthenticated lender portal —
    # deliberately OUTSIDE the pipeline/projects/<project_id>/ namespace
    # above (a lender is never handed a project_id, only an opaque
    # access_token) — see views.py's comment block above
    # LenderSharePortalView for the 401-avoidance discipline every view
    # here follows.
    path('share/deal/<str:access_token>/', LenderSharePortalView.as_view(), name='lender_share_portal'),
    path(
        'share/deal/<str:access_token>/excel/',
        LenderShareExcelDownloadView.as_view(), name='lender_share_excel_download',
    ),
    path(
        'share/deal/<str:access_token>/vault/<int:file_id>/',
        LenderShareVaultFileDownloadView.as_view(), name='lender_share_vault_download',
    ),

    # V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3).
    # Connection-level (CustomUser's own mailbox — see EmailOAuthConnection's
    # own docstring on why these are never project- or org-scoped):
    path('email-sync/connections/', EmailConnectionListView.as_view(), name='email_sync_connections'),
    path('email-sync/connect/', EmailConnectStartView.as_view(), name='email_sync_connect_start'),
    # Must match email_oauth._redirect_uri() character-for-character —
    # this exact path is what gets registered as the "Authorized redirect
    # URI" in Google Cloud Console / the Microsoft Entra App Registration.
    path(
        'email-sync/oauth/callback/<str:provider>/',
        EmailOAuthCallbackView.as_view(), name='email_sync_oauth_callback',
    ),
    path(
        'email-sync/connections/<str:provider>/disconnect/',
        EmailConnectionDisconnectView.as_view(), name='email_sync_disconnect',
    ),

    # Thread-level (explicit per-deal "tag this thread" action — gated
    # through visible_projects_for exactly like every other pipeline
    # drawer panel above).
    path(
        'email-sync/projects/<int:project_id>/thread-search/',
        EmailThreadSearchView.as_view(), name='email_sync_thread_search',
    ),
    path(
        'email-sync/projects/<int:project_id>/threads/',
        ProjectEmailThreadsView.as_view(), name='email_sync_project_threads',
    ),
    path(
        'email-sync/projects/<int:project_id>/threads/<int:thread_link_id>/',
        ProjectEmailThreadUnlinkView.as_view(), name='email_sync_project_thread_unlink',
    ),
]
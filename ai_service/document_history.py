# backend/ai_service/document_history.py
"""
Helpers for recording processed documents (History feature) and mapping raw
uploaded files to human-readable format labels.
"""
import os

FORMAT_LABELS = {
    '.pdf': 'PDF',
    '.doc': 'Word',
    '.docx': 'Word',
    '.xls': 'Excel',
    '.xlsx': 'Excel',
    '.csv': 'CSV',
}


def get_file_format_label(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    return FORMAT_LABELS.get(ext, ext.lstrip('.').upper() or 'Unknown')


def record_processed_documents(organization, uploaded_by, om_file, t12_file, rent_roll_file, status: str):
    """
    Writes one ProcessedDocument row per source file (OM, T12, Rent Roll).

    NOTE: this records ALL THREE files with the SAME status. If a failure
    happened partway through processing, we don't currently know which
    specific file caused it — so on failure, all three get marked 'failed'
    even though only one may actually be at fault. This is an approximation,
    not per-file failure attribution; worth revisiting if that granularity
    ever matters.

    organization may be None (user with no org attached) — in that case,
    nothing is recorded, since ProcessedDocument requires an organization.
    """
    from .models import ProcessedDocument  # local import avoids circulars at module load

    if organization is None:
        return

    files_and_types = [
        (om_file, 'OM'),
        (t12_file, 'T12'),
        (rent_roll_file, 'RENT_ROLL'),
    ]

    records = []
    for file_obj, doc_type in files_and_types:
        if file_obj is None:
            continue
        records.append(ProcessedDocument(
            organization=organization,
            uploaded_by=uploaded_by,
            file_name=file_obj.name,
            document_type=doc_type,
            file_format=get_file_format_label(file_obj.name),
            file_size_bytes=getattr(file_obj, 'size', 0) or 0,
            status=status,
        ))

    if records:
        ProcessedDocument.objects.bulk_create(records)


def record_processed_ppm_document(organization, uploaded_by, ppm_file, status: str):
    """
    PPM equivalent of record_processed_documents. The PPM pipeline processes
    exactly ONE file per request (see PPMUnderwriteView in views.py — the
    frontend's batch console loops this call once per file rather than
    submitting several at once), so this writes a single ProcessedDocument
    row rather than the three-files-at-once shape used by the CRE pipeline.

    organization may be None (user with no org attached) — in that case,
    nothing is recorded, matching record_processed_documents' behavior.
    """
    from .models import ProcessedDocument

    if organization is None or ppm_file is None:
        return

    ProcessedDocument.objects.create(
        organization=organization,
        uploaded_by=uploaded_by,
        file_name=ppm_file.name,
        document_type='PPM',
        file_format=get_file_format_label(ppm_file.name),
        file_size_bytes=getattr(ppm_file, 'size', 0) or 0,
        status=status,
    )


def record_analysis_report(organization, uploaded_by, metrics: dict, tier: str,
                            processing_seconds: float = None, document_type: str = 'CRE',
                            project_id=None, rent_roll_data=None, t12_monthly_data=None):
    """
    Saves a completed analysis result for later review on the Outputs page.
    Also the source of truth for upload-quota counting (see views.py) — one
    row per successful analysis, regardless of document_type (CRE and PPM
    analyses deliberately draw from the same rolling 30-day quota pool).
    organization may be None (user with no org attached) — in that case,
    nothing is recorded.

    document_type defaults to 'CRE' so every existing call site (the
    OM/T12/Rent Roll pipeline) keeps working unchanged without passing it.

    Feature 3 (Enterprise Pipeline & Workflow Dashboard): AnalysisReport.project
    is a REQUIRED field (see models.py / migration 0009) — every row created
    from here on must have one, so this function resolves one unconditionally:

    - project_id, when given, links this run to an EXISTING PropertyProject —
      the "Re-upload & Re-analyze Files" flow, clicked from inside a specific
      project's view. It's looked up scoped to THIS organization, same
      reasoning as the workspace/organization match enforced elsewhere — a
      project_id ultimately comes from client input and must never be
      trusted to belong to the caller's own org just because it was sent.
      An id that doesn't resolve (wrong org, deleted, or simply made up)
      falls through to creating a new project below exactly as if none had
      been given, rather than erroring the whole analysis out over it.
    - Otherwise (the default — e.g. a fresh upload from the main dashboard,
      or any PPM analysis, which has no pipeline UI of its own) a brand-new
      single-analysis PropertyProject is created automatically, per the
      approved UX: uploads from the main dashboard always start a new deal.

    A newly-created project inherits uploaded_by's workspace (None if they
    don't have one yet) — without this, a member's own new deal would be
    invisible to their own teammates in the same workspace, defeating the
    point of workspaces. An EXISTING project (project_id path) keeps
    whatever workspace it already has; re-analyzing a deal never moves it.

    rent_roll_data/t12_monthly_data are the raw, JSON-serializable inputs
    this run's Excel output was built from — see AnalysisReport's own field
    docstrings for exactly what each holds. Optional (None when not
    applicable/available — e.g. PPM analyses, or a T12 format Module 7 can't
    parse monthly columns from); a row missing one simply can't support
    historical regeneration of that part later.
    """
    from .models import AnalysisReport, PropertyProject

    if organization is None:
        return

    if document_type == 'PPM':
        # PPM metrics are a flat dict with target_asset_name at the top
        # level — there's no nested property_metadata section like the CRE
        # shape has.
        property_name = metrics.get('target_asset_name') or ''
    else:
        property_name = (metrics.get('property_metadata') or {}).get('property_name') or ''

    project = None
    if project_id is not None:
        project = PropertyProject.objects.filter(pk=project_id, organization=organization).first()

    if project is None:
        project = PropertyProject.objects.create(
            organization=organization,
            workspace=getattr(uploaded_by, 'workspace', None),
            property_name=property_name,
            created_by=uploaded_by,
        )

    AnalysisReport.objects.create(
        organization=organization,
        uploaded_by=uploaded_by,
        project=project,
        property_name=property_name,
        tier=tier,
        document_type=document_type,
        metrics=metrics,
        processing_seconds=processing_seconds,
        status='ready',
        rent_roll_data=rent_roll_data,
        t12_monthly_data=t12_monthly_data,
    )
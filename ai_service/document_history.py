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
                            processing_seconds: float = None, document_type: str = 'CRE'):
    """
    Saves a completed analysis result for later review on the Outputs page.
    Also the source of truth for upload-quota counting (see views.py) — one
    row per successful analysis, regardless of document_type (CRE and PPM
    analyses deliberately draw from the same rolling 30-day quota pool).
    organization may be None (user with no org attached) — in that case,
    nothing is recorded.

    document_type defaults to 'CRE' so every existing call site (the
    OM/T12/Rent Roll pipeline) keeps working unchanged without passing it.
    """
    from .models import AnalysisReport

    if organization is None:
        return

    if document_type == 'PPM':
        # PPM metrics are a flat dict with target_asset_name at the top
        # level — there's no nested property_metadata section like the CRE
        # shape has.
        property_name = metrics.get('target_asset_name') or ''
    else:
        property_name = (metrics.get('property_metadata') or {}).get('property_name') or ''

    AnalysisReport.objects.create(
        organization=organization,
        uploaded_by=uploaded_by,
        property_name=property_name,
        tier=tier,
        document_type=document_type,
        metrics=metrics,
        processing_seconds=processing_seconds,
        status='ready',
    )
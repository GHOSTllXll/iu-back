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


def record_analysis_report(organization, uploaded_by, metrics: dict, tier: str):
    """
    Saves a completed analysis result for later review on the Outputs page.
    organization may be None (user with no org attached) — in that case,
    nothing is recorded.
    """
    from .models import AnalysisReport

    if organization is None:
        return

    property_name = (metrics.get('property_metadata') or {}).get('property_name') or ''

    AnalysisReport.objects.create(
        organization=organization,
        uploaded_by=uploaded_by,
        property_name=property_name,
        tier=tier,
        metrics=metrics,
        status='ready',
    )
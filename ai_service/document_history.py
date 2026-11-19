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
                            project_id=None, rent_roll_data=None, t12_monthly_data=None,
                            asking_price=None, documents_provided=None):
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

    Sprint "Workflow Upgrades" Feature 1:

    - asking_price: the analyst's own entered/confirmed price for THIS deal
      (from the upload screen's "Target Purchase Price" field, or the
      download screen's). None means "the analyst didn't enter one for this
      run" — never written over an existing PropertyProject.asking_price in
      that case, so a later run with no price typed in doesn't blank out a
      price an earlier run already recorded. Only applied when truthy.
    - documents_provided: {"om": bool, "t12": bool, "rent_roll": bool} for
      THIS run specifically — see AnalysisReport.documents_provided's own
      docstring in models.py. Defaults to {} (not None) on the row so
      reader code's dict.get(...) calls never hit a None.
    """
    from .models import AnalysisReport, PropertyProject, DealActivityLog

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

    # V2 Module 4a (Buy-Box Auto-Alerts): recorded BEFORE project gets
    # reassigned to the newly-created row below — see the buy-box check
    # at the bottom of this function for why this must only ever fire
    # for a brand-new deal, never a re-analysis of an existing one.
    is_new_project = project is None

    if project is None:
        project = PropertyProject.objects.create(
            organization=organization,
            workspace=getattr(uploaded_by, 'workspace', None),
            property_name=property_name,
            created_by=uploaded_by,
            asking_price=asking_price,
        )
        # Pipeline Modules Phase 1: seed the Activity feed with a "deal
        # created" row (from_stage=None) so it always has a starting
        # point — see DealActivityLog's own docstring.
        DealActivityLog.objects.create(
            project=project,
            user=uploaded_by,
            from_stage=None,
            to_stage=project.pipeline_stage,
        )
    elif asking_price:
        # Existing project, re-analyzed: only overwrite when THIS run
        # actually supplied a price — never null out a previously-recorded
        # one just because this particular re-upload left the field blank.
        project.asking_price = asking_price
        project.save(update_fields=['asking_price'])

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
        documents_provided=documents_provided or {},
    )

    # Pipeline Modules Phase 4 (Persistent Missing-Doc Checklist): seed one
    # DealChecklistItem per core document THIS run reports missing — only
    # for CRE analyses, since documents_provided's om/t12/rent_roll shape
    # doesn't apply to a PPM run (see AnalysisReport.document_type's own
    # docstring on the two pipelines sharing this table).
    if document_type == 'CRE':
        _seed_missing_doc_checklist_items(project, documents_provided, uploaded_by)

        # V2 Module 4a (Buy-Box Auto-Alerts): only on a brand-new deal's
        # FIRST analysis — see is_new_project above. A re-analyzed
        # existing deal never re-checks/re-alerts, so a deal already
        # sitting in the pipeline doesn't re-notify the team every time
        # it's reprocessed (e.g. after a document correction).
        if is_new_project:
            _check_buy_box_match_and_notify(project, uploaded_by, metrics)


def _seed_missing_doc_checklist_items(project, documents_provided, uploaded_by):
    """
    "Only seed if no item with that doc_key exists yet for THIS project,
    regardless of resolved state" — the whole point of this persistence
    rule is that a resolved item (e.g. "got the T12 from the broker over
    email, outside this platform") must never be silently re-created just
    because a later re-analysis still reports that same document missing.
    So this is a one-way ratchet: a doc_key gets exactly one
    DealChecklistItem ever created for a given project, on whichever
    analysis run first reports it missing — never re-evaluated or
    re-seeded after that, resolved or not. A doc_key that was never
    reported missing on any run simply never gets a row, same as before
    this feature existed.
    """
    from .models import DealChecklistItem

    documents_provided = documents_provided or {}
    missing_doc_keys = [
        doc_key for doc_key, _label in DealChecklistItem.DOC_KEY_CHOICES
        # Mirrors _compute_deal_scorecard's own read convention (views.py):
        # an ABSENT key means "was provided" (.get(key, True)), so only an
        # EXPLICIT False here counts as missing.
        if documents_provided.get(doc_key, True) is False
    ]
    if not missing_doc_keys:
        return

    already_seeded = set(
        DealChecklistItem.objects
        .filter(project=project, doc_key__in=missing_doc_keys)
        .values_list('doc_key', flat=True)
    )
    doc_key_labels = dict(DealChecklistItem.DOC_KEY_CHOICES)
    new_items = [
        DealChecklistItem(
            project=project,
            doc_key=doc_key,
            label=f"Missing: {doc_key_labels[doc_key]}",
            created_by=uploaded_by,
        )
        for doc_key in missing_doc_keys
        if doc_key not in already_seeded
    ]
    if new_items:
        DealChecklistItem.objects.bulk_create(new_items)


def _check_buy_box_match_and_notify(project, uploaded_by, metrics):
    """
    V2 Module 4a (Buy-Box Auto-Alerts): if this brand-new deal matches
    its workspace's configured buy-box criteria (see
    buy_box.check_buy_box_match and TeamWorkspace.get_buy_box_criteria
    in users/models.py), notifies that workspace's leadership — every
    Workspace Head of THIS workspace, plus every Org Admin of this
    organization (the same "need-to-know" leadership tier
    OverrideAlertFeedView/_get_leadership_digest_data treat as "the
    people who'd want a firm-wide heads-up", not every analyst in the
    org) — via BOTH an in-app Notification row (see models.py) and a
    transactional email (reusing send_mail, the exact same function
    PasswordResetView already uses — see users/views.py — no new email
    provider wired up for this).

    No-op entirely when the deal has no workspace (buy-box criteria is a
    per-workspace setting; a workspace-less deal has nothing to check
    against) or when check_buy_box_match returns None (alerts disabled,
    or the deal simply doesn't qualify).

    Best-effort for both side effects: a Notification write or an email
    send failing here must never fail the underlying analysis request
    that triggered it — same "log and swallow" posture as
    _sync_grid_edits_to_working_state in views.py.
    """
    if project.workspace_id is None:
        return

    from django.db.models import Q
    from django.core.mail import send_mail
    from django.conf import settings
    from .buy_box import check_buy_box_match
    from .models import Notification
    from users.models import CustomUser

    try:
        criteria = project.workspace.get_buy_box_criteria()
        match = check_buy_box_match(criteria, metrics, project.asking_price)
        if match is None:
            return

        recipients = list(
            CustomUser.objects.filter(organization_id=project.organization_id)
            .filter(Q(role='admin') | Q(workspace_id=project.workspace_id, workspace_role='head'))
            .exclude(email=getattr(uploaded_by, 'email', None))
            .distinct()
        )
        if not recipients:
            return

        address = (metrics.get('property_metadata') or {}).get('address') or 'address not extracted'
        criteria_text = '; '.join(match['matched_criteria']) or 'this workspace\'s buy-box criteria'
        message = (
            f"New deal \"{project.property_name or 'Untitled'}\" ({address}) matches "
            f"{project.workspace.name}'s buy-box: {criteria_text}."
        )

        Notification.objects.bulk_create([
            Notification(
                recipient=recipient, notification_type='buy_box_match', project=project, message=message,
            )
            for recipient in recipients
        ])

        deal_link = f"{getattr(settings, 'FRONTEND_BASE_URL', '')}/dashboard/pipeline?project={project.id}"
        for recipient in recipients:
            try:
                send_mail(
                    subject=f"Buy-Box Match: {project.property_name or 'New Deal'}",
                    message=f"{message}\n\nView deal: {deal_link}",
                    from_email=getattr(settings, 'DEFAULT_FROM_EMAIL', 'noreply@inboundunderwriting.com'),
                    recipient_list=[recipient.email],
                    fail_silently=True,
                )
            except Exception:
                # Best-effort — an SMTP/provider hiccup on one recipient's
                # email must never block the others or the deal creation
                # that triggered this whole check.
                pass
    except Exception:
        # Buy-box matching/notifying is a value-add, never a hard
        # requirement for the underlying analysis to succeed — same
        # defensive posture as Module 1 reconciliation above.
        pass
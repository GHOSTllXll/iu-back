# backend/ai_service/team_scorecard.py
"""
V2 Module 4b (Analyst Speed & Accuracy Scorecard): per-analyst
leadership metrics for the "Team Performance" tab of the Leadership
Morning Digest — see views.py's comment block above TeamScorecardView.

Takes the VIEWER's already-scoped PropertyProject queryset
(visible_projects_for(user) — exactly the same function every other
leadership view in this codebase uses) rather than re-deriving
visibility here, so "per analyst" naturally means "per analyst this
viewer can see": an Org Admin's scorecard is firm-wide, a Workspace
Head's is their own workspace only — no special case.

Four metrics, per the client's explicit scope decision:
  - deals_underwritten: count of deals (any pipeline stage) assigned to
    this analyst, within the viewer's visible scope.
  - avg_turnaround_hours: average hours from a deal's creation to the
    FIRST time it ever reached 'vp_review', across that analyst's deals
    that have reached vp_review at least once. A deal still short of
    vp_review has no turnaround to measure yet and isn't counted.
  - post_upload_variance_rate_pct: of this analyst's "Re-upload Edited
    Underwriting File" events (ExcelOverrideUpload), the percentage
    that produced at least one GPR/OpEx/NOI change exceeding
    OVERRIDE_ALERT_THRESHOLD_PCT — reusing OVERRIDE_ALERT_TRACKED_
    FIELDS/_THRESHOLD_PCT from views.py verbatim (passed in as
    arguments rather than re-imported, to avoid a views<->this-module
    circular import), so this can never silently disagree with the
    firm-wide proactive alert feed about what counts.
  - forensic_flags_triggered: total count of Module 1 reconciliation
    flags (reconciliation.py's "Bullshit Detector") across this
    analyst's deals' CURRENT (latest) analysis each — "how many
    forensic flags are live right now across this analyst's book",
    not a lifetime/cumulative event count across every past re-analysis.
"""
from collections import defaultdict


def compute_analyst_scorecards(projects_qs, override_alert_tracked_fields: dict, override_alert_threshold_pct: float) -> list:
    from .models import ExcelOverrideUpload, ExcelOverrideDelta

    projects = list(
        projects_qs.filter(assigned_analyst__isnull=False).select_related('assigned_analyst')
    )

    by_analyst = defaultdict(list)
    for project in projects:
        by_analyst[project.assigned_analyst].append(project)

    rows = []
    for analyst, analyst_projects in by_analyst.items():
        project_ids = [p.id for p in analyst_projects]

        # --- Average Turnaround Time (new_incoming -> vp_review) ---
        turnaround_hours = []
        for project in analyst_projects:
            first_vp_review = project.activity_log.filter(to_stage='vp_review').order_by('created_at').first()
            if first_vp_review is not None:
                delta = first_vp_review.created_at - project.created_at
                turnaround_hours.append(delta.total_seconds() / 3600.0)
        avg_turnaround_hours = (
            round(sum(turnaround_hours) / len(turnaround_hours), 1) if turnaround_hours else None
        )

        # --- >2.0% Post-Upload Variance Rate ---
        uploads = ExcelOverrideUpload.objects.filter(project_id__in=project_ids)
        total_uploads = uploads.count()
        qualifying_upload_ids = set(
            ExcelOverrideDelta.objects.filter(
                upload__in=uploads,
                field_name__in=override_alert_tracked_fields,
                variance_pct__isnull=False,
                variance_pct__gt=override_alert_threshold_pct,
            ).values_list('upload_id', flat=True)
        )
        variance_rate_pct = (
            round(len(qualifying_upload_ids) / total_uploads * 100, 1) if total_uploads else None
        )

        # --- Forensic Seller-Trap Flags Triggered (Module 1 reconciliation) ---
        forensic_flags_triggered = 0
        for project in analyst_projects:
            latest = project.latest_analysis
            if latest is None:
                continue
            flags = ((latest.metrics or {}).get('reconciliation') or {}).get('flags') or []
            forensic_flags_triggered += len(flags)

        display_name = f"{analyst.first_name} {analyst.last_name}".strip() or analyst.email
        rows.append({
            'analyst_email': analyst.email,
            'analyst_name': display_name,
            'deals_underwritten': len(analyst_projects),
            'avg_turnaround_hours': avg_turnaround_hours,
            'post_upload_variance_rate_pct': variance_rate_pct,
            'forensic_flags_triggered': forensic_flags_triggered,
        })

    rows.sort(key=lambda r: r['analyst_name'].lower())
    return rows

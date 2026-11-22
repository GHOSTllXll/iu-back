# backend/ai_service/models.py
import math
import secrets
import uuid
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.db import models
from django.db.models.signals import post_delete
from django.dispatch import receiver
from django.utils import timezone


def _sanitize_json_floats(value):
    """
    Recursively replaces NaN/Infinity/-Infinity floats with None.

    MySQL's JSON column type (unlike Postgres' native json/jsonb) rejects
    these outright -- they aren't valid per the JSON spec, Python's json
    module just doesn't enforce that by default -- so a divide-by-zero or
    0/0 anywhere upstream (a metrics ratio, a rent-roll calc over an empty
    set, etc.) silently produces a NaN/inf that sails through every
    Python-side check and only then blows up the actual INSERT/UPDATE with
    MySQL error 3140 ("Invalid JSON text"), right at the very end of an
    otherwise-successful analysis run. numpy.float64 (what pandas/financial
    math actually hand back) is itself a float subclass, so this catches
    those too, not just plain Python floats.
    """
    if isinstance(value, float):
        return None if (math.isnan(value) or math.isinf(value)) else value
    if isinstance(value, dict):
        return {k: _sanitize_json_floats(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_sanitize_json_floats(v) for v in value]
    return value


class PrivateFileSystemStorage(FileSystemStorage):
    """
    A FileSystemStorage whose .url() unconditionally raises, so these files
    can never be served through a public URL -- NOT achieved by passing
    base_url=None to the base class. VERIFIED directly (not assumed): doing
    that silently falls back to settings.MEDIA_URL instead of disabling
    .url() at all, which would have been a real information-disclosure
    gap the moment anyone (now or in some future refactor) ever called
    .file.url on one of these. Overriding the method closes that off for
    good, regardless of what MEDIA_URL is set to.
    """
    def url(self, name):
        raise ValueError(
            "Files in this private storage have no public URL by design -- "
            "read them server-side via storage.path()/.open(), from an "
            "authenticated, visibility-checked view, never through a "
            "public URL. (Shared by ClientExcelTemplate and DealVaultFile "
            "below -- both are private, org/deal-scoped file stores.)"
        )


# Private storage for client-provided Excel templates — deliberately NOT
# Django's default file storage, which is typically wired to a public
# MEDIA_URL. These are a client's proprietary valuation models — nothing
# here should ever be reachable by a guessable URL the way STATIC_URL-served
# files are; the only way to read one back is server-side, via this
# storage's own .path()/.open(), from an authenticated admin view (see
# users/views.py's AdminGetExcelTemplateView).
#
# A CALLABLE, not a Storage instance, passed to FileField's storage= below.
# Django's migration writer calls .deconstruct() on a bare Storage instance
# and bakes the resolved location path into the migration file literally —
# wrong the moment that migration runs on a different machine/path. A
# callable is instead serialized as a reference to this function itself,
# and gets called (and settings.PRIVATE_MEDIA_ROOT re-resolved) fresh every
# time Django actually needs the storage.
def private_template_storage():
    return PrivateFileSystemStorage(
        location=str(settings.PRIVATE_MEDIA_ROOT / 'excel_templates'),
    )


def client_template_upload_path(instance, filename):
    """
    Keyed by organization id, not the client's own filename — two clients
    with similarly-named files can never collide, and a client can't
    influence where their file lands on disk by naming it something odd.
    One fixed path per org (not a timestamped/unique one) is deliberate:
    re-uploading a replacement template should overwrite the old file
    rather than silently accumulating orphaned old versions on disk — see
    AdminUploadExcelTemplateView, which deletes the previous file first.
    """
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'xlsx'
    return f"org_{instance.organization_id}/template.{ext}"


# Pipeline Modules Phase 5 (Excel Versioning & Supporting Document Vault):
# a SEPARATE private store from the one above, for the exact same reason
# (see PrivateFileSystemStorage's own docstring) but scoped by DEAL
# (PropertyProject) rather than by organization, and holding a growing
# version history rather than one reusable asset — see DealVaultFile's own
# docstring below for the full design. A distinct subfolder
# ('deal_vault', not 'excel_templates') keeps the two stores from ever
# colliding on disk even though both resolve under the same
# PRIVATE_MEDIA_ROOT. Also a callable, for the identical migration-
# portability reason given above private_template_storage.
def deal_vault_storage():
    return PrivateFileSystemStorage(
        location=str(settings.PRIVATE_MEDIA_ROOT / 'deal_vault'),
    )


def deal_vault_upload_path(instance, filename):
    """
    Keyed by project id, then a random uuid4 — the OPPOSITE convention
    from client_template_upload_path above. That function deliberately
    overwrites (ClientExcelTemplate is one reusable asset, not a version
    history); DealVaultFile exists specifically to keep every version, so
    each upload needs its own never-colliding path on disk instead.
    """
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'bin'
    return f"project_{instance.project_id}/{uuid.uuid4().hex}.{ext}"


class ProcessedDocument(models.Model):
    """
    A record of a single source document (OM, T12, or Rent Roll) that has been
    run through analysis. The file itself is NEVER stored — only this metadata
    — per the platform's "files are destroyed after analysis" requirement.
    Exists so users can see what's already been analyzed (avoid re-uploading
    the same document) and so the org has a lightweight audit trail.

    NOTE on 'processing' status: this system currently processes everything
    synchronously within a single request, so a record is only ever written
    once a request has already finished — meaning today, status will only
    ever land as COMPLETED or FAILED. 'processing' is kept as a valid choice
    for when background task processing (Celery) is introduced.
    """
    STATUS_CHOICES = [
        ('completed', 'Completed'),
        ('processing', 'Processing'),
        ('failed', 'Failed'),
    ]

    DOCUMENT_TYPE_CHOICES = [
        ('OM', 'Offering Memorandum'),
        ('T12', 'T12 Statement'),
        ('RENT_ROLL', 'Rent Roll'),
        ('PPM', 'Private Placement Memorandum'),
    ]

    organization = models.ForeignKey(
        'users.Organization',
        on_delete=models.CASCADE,
        related_name='processed_documents'
    )
    uploaded_by = models.ForeignKey(
        'users.CustomUser',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='processed_documents'
    )

    is_deleted = models.BooleanField(default=False)
    file_name = models.CharField(max_length=255)
    document_type = models.CharField(max_length=20, choices=DOCUMENT_TYPE_CHOICES)
    file_format = models.CharField(max_length=10)  # e.g. 'PDF', 'Excel', 'Word', 'CSV'
    file_size_bytes = models.BigIntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='completed')

    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-uploaded_at']
        indexes = [
            models.Index(fields=['organization', '-uploaded_at']),
        ]

    def __str__(self):
        return f"{self.file_name} ({self.status}) — {self.organization.name}"


class PropertyProject(models.Model):
    """
    Feature 3 (Enterprise Pipeline & Workflow Dashboard): a stable wrapper
    around one or more AnalysisReport rows for the SAME deal. Exists because
    an AnalysisReport is a single analysis RUN — real-estate deals are
    iterative, re-analyzed as negotiations progress, and pipeline_stage/Team
    Notes need to survive that re-analysis rather than resetting with every
    new run. See AnalysisReport.project below for how the two connect.

    Every PropertyProject existed in some form before this field did, too:
    the data migration that introduces this table also auto-wraps every
    pre-existing AnalysisReport into its own single-analysis
    PropertyProject, by deliberate decision — no attempt is made to guess
    which historical analyses were actually re-runs of the same deal.
    """
    # Pipeline Modules Phase 1 (Daily Deal Pipeline & Team Command Center):
    # replaces the old 3-state STATUS_CHOICES/deal_status (reviewing/
    # approved/archived) outright, per explicit client decision — that was
    # a business OUTCOME, not a workflow funnel position, and didn't map
    # cleanly onto this 5-stage pipeline (e.g. "archived" and "dead/passed"
    # aren't the same axis). See migration 0015's data migration for the
    # one-time mapping applied to every existing row:
    # reviewing -> underwriting, approved -> offer_submitted,
    # archived -> dead_passed. Values are lowercase snake_case, matching
    # this field's own predecessor and every other choices field in this
    # codebase — not the spec doc's ALL_CAPS enum names.
    PIPELINE_STAGE_CHOICES = [
        ('new_incoming', 'New Incoming'),
        ('underwriting', 'Underwriting'),
        ('vp_review', 'VP Review'),
        ('offer_submitted', 'Offer Submitted'),
        ('dead_passed', 'Dead / Passed'),
    ]

    organization = models.ForeignKey(
        'users.Organization', on_delete=models.CASCADE, related_name='property_projects'
    )
    workspace = models.ForeignKey(
        'users.TeamWorkspace', on_delete=models.SET_NULL, null=True, blank=True, related_name='property_projects'
    )

    property_name = models.CharField(max_length=255, blank=True, default='')
    pipeline_stage = models.CharField(max_length=20, choices=PIPELINE_STAGE_CHOICES, default='new_incoming')

    created_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True, related_name='created_projects'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    is_deleted = models.BooleanField(default=False)

    # Sprint "Workflow Upgrades" Feature 1 (10-Second Scorecard): the
    # analyst's own entered/confirmed asking price for this deal — NOT the
    # AI-extracted deal_valuation.target_purchase_price (whatever figure
    # the OM itself stated), which is a separate, unconfirmed value that
    # already lives inside AnalysisReport.metrics. Captured optionally on
    # the upload screen (or at Excel-download time, the pre-existing
    # "Target Purchase Price" field) and persisted here so the Deal
    # Scorecard's cap-rate/price-per-unit math and Feature 5's DSCR/max-
    # loan sizing both have it automatically on every later view of this
    # deal, without re-asking. Null until an analyst actually provides one
    # — never defaulted or backfilled from the AI's own OM-derived guess.
    asking_price = models.FloatField(null=True, blank=True)

    # Pipeline Modules Phase 1: who's actively working this deal day to
    # day — distinct from created_by (who ran the first analysis) and from
    # workspace membership (who CAN see it per permissions.py). Nullable:
    # assignment is an explicit later action, never inferred from
    # created_by or auto-assigned on creation.
    assigned_analyst = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='assigned_deals'
    )
    # Pipeline Modules Phase 1: the analyst's own target date for closing
    # or submitting an offer — a planning input the team sets by hand,
    # never derived or defaulted by the system.
    target_closing_or_offer_date = models.DateField(null=True, blank=True)

    class Meta:
        ordering = ['-updated_at']
        indexes = [
            models.Index(fields=['organization', '-updated_at']),
            models.Index(fields=['workspace', '-updated_at']),
        ]

    def __str__(self):
        return f"{self.property_name or 'Untitled'} [{self.pipeline_stage}] — {self.organization.name}"

    @property
    def latest_analysis(self):
        return self.analyses.order_by('-created_at').first()


class DealActivityLog(models.Model):
    """
    Pipeline Modules Phase 1 (Daily Deal Pipeline & Team Command Center):
    one row per pipeline_stage transition on a PropertyProject — the
    system-event half of Module 2's "Deal Activity & Notes" drawer
    (TeamNote below is the free-text half). Deliberately scoped to STAGE
    TRANSITIONS ONLY for v1, not a general-purpose audit log for every
    field edit on the project:
      - NOT the same thing as users.ActivityLog, which is a platform-level
        admin audit trail (login, billing, team-member management) keyed
        by a free-text `target` string, not a project FK — this table is
        per-deal and queryable by project for exactly that deal's own
        timeline.
      - NOT the same thing as ExcelOverrideUpload/-Delta below, which
        audits re-uploaded Excel files and their cell-level diffs, not
        stage movement.

    from_stage is null only for the very first row on a project — logged
    at PropertyProject creation time (see record_analysis_report in
    document_history.py) so the Activity feed always has a starting point
    rather than silently beginning mid-history on whatever stage change
    happens to come first.
    """
    project = models.ForeignKey(
        PropertyProject, on_delete=models.CASCADE, related_name='activity_log'
    )
    user = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='deal_stage_transitions'
    )
    from_stage = models.CharField(
        max_length=20, choices=PropertyProject.PIPELINE_STAGE_CHOICES, null=True, blank=True
    )
    to_stage = models.CharField(max_length=20, choices=PropertyProject.PIPELINE_STAGE_CHOICES)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['project', '-created_at'], name='deal_activity_project_idx'),
        ]

    def __str__(self):
        who = self.user.email if self.user else 'Unknown'
        return f"{self.project}: {self.from_stage or '—'} -> {self.to_stage} by {who}"


class TeamNote(models.Model):
    """
    Feature 3: a free-text note on a PropertyProject, posted by a team
    member (e.g. "Waiting on updated rent roll from broker"). Deliberately
    flat and unthreaded — v1 scope is a running log, not a comment thread;
    see IssueReport's docstring for the same reasoning applied elsewhere in
    this codebase.
    """
    project = models.ForeignKey(PropertyProject, on_delete=models.CASCADE, related_name='notes')
    author = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True, related_name='team_notes'
    )
    body = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)

    # Pipeline Modules Phase 4 (@Mentions, in-app only per explicit client
    # decision -- no email notification infrastructure exists in this
    # codebase yet): every org member whose full email address was found
    # inside this note's body text at post time -- see _parse_mentions()
    # in views.py for the exact regex/matching rule (same org-scoped
    # CustomUser.objects.filter(email=..., organization_id=...) pattern
    # already used for PropertyProject.assigned_analyst). Populated once,
    # at creation (TeamNoteListCreateView.post()) -- editing a note's body
    # isn't supported anywhere in this codebase, so there's no "re-parse
    # on edit" path to worry about. Powers the drawer's "mentions you"
    # badge/filter purely in-app.
    mentioned_users = models.ManyToManyField(
        'users.CustomUser', blank=True, related_name='deal_mentions'
    )

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"Note on {self.project} by {self.author.email if self.author else 'Unknown'}"


class AnalysisReport(models.Model):
    """
    Stores the RESULT of a completed analysis (the metrics JSON) so a user can
    come back and review it later — this is the "Outputs" page. Also doubles
    as the source of truth for upload-quota counting (see views.py) — one row
    per successful analysis, so counting rows in a time window == counting
    "uploads used", with no separate counter field needed.

    Historically did NOT store the generated Excel file — only Preview was
    supported on the Outputs page (no Download) — just the derived metrics
    JSON. Feature 3 (Enterprise Pipeline & Workflow Dashboard) adds
    rent_roll_data/t12_monthly_data (see below) so NEW analyses can regenerate
    their Excel file on demand from a PropertyProject's history. Rows created
    before that change have neither field populated, so Download stays
    unavailable for them specifically — see each field's own docstring for
    exactly what is (and deliberately is not) persisted going forward.

    NOTE on 'status': like ProcessedDocument, this system currently processes
    synchronously, so a report row is only ever created after a request has
    already finished successfully. 'generating' is kept as a valid choice for
    when background processing exists, but in practice today every row lands
    as 'ready'.

    NOTE on 'document_type': added for the Institutional PPM pipeline, which
    shares this exact model/table rather than getting its own — same quota
    counting (get_period_usage doesn't filter by document_type, so PPM and
    CRE analyses draw from the same rolling 30-day pool, by design), same
    Outputs page, same dashboard stats, for free. Defaults to 'CRE' so every
    existing row (and every existing call site) keeps working unchanged.
    """
    STATUS_CHOICES = [
        ('ready', 'Ready'),
        ('generating', 'Generating'),
    ]

    DOCUMENT_TYPE_CHOICES = [
        ('CRE', 'CRE Underwriting (OM/T12/Rent Roll)'),
        ('PPM', 'Private Placement Memorandum'),
    ]

    organization = models.ForeignKey(
        'users.Organization',
        on_delete=models.CASCADE,
        related_name='analysis_reports'
    )
    uploaded_by = models.ForeignKey(
        'users.CustomUser',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='analysis_reports'
    )

    # Feature 3: links this run to its deal wrapper. Was nullable for the
    # brief migration window between this field being added and the data
    # migration backfilling it (every existing row auto-wrapped into its own
    # single-analysis PropertyProject) — now required, since that backfill
    # is confirmed complete (0 rows with project_id NULL) and every row
    # created going forward always has one from the moment it's created.
    project = models.ForeignKey('PropertyProject', on_delete=models.CASCADE, related_name='analyses')

    is_deleted = models.BooleanField(default=False)
    property_name = models.CharField(max_length=255, blank=True, default='')
    tier = models.CharField(max_length=20)
    document_type = models.CharField(max_length=10, choices=DOCUMENT_TYPE_CHOICES, default='CRE')
    metrics = models.JSONField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='ready')

    # Feature 3: the raw rent roll rows and monthly T12 series behind this
    # analysis's `metrics`, persisted so a later "Re-upload & Re-analyze" or
    # a Download Excel click can regenerate a full workbook without needing
    # the source files again (which are never stored — see ProcessedDocument).
    # Deliberately does NOT include debt assumptions (purchase price, loan
    # terms, etc.) — those are re-prompted for at download time rather than
    # persisted, since they're scenario inputs the user may want to vary
    # between downloads, not a fixed property of the analysis itself. Null
    # on every row created before this field existed, which is exactly the
    # set of rows Download Excel stays unavailable for ("Detailed download
    # unavailable for legacy analyses processed prior to system upgrade.").
    rent_roll_data = models.JSONField(null=True, blank=True)
    t12_monthly_data = models.JSONField(null=True, blank=True)

    # Feature 4 (Cell Override Audit Log): exactly which value this system
    # injected into each IU_* field the LAST time an Excel was actually
    # generated for this analysis — {"scalars": {name: value, ...}, "rows":
    # [{name: value, ...}, one dict per tenant]}. Computed by
    # custom_template_generator.compute_injected_values_snapshot() against
    # this row's own metrics/rent_roll_data/debt assumptions, and (re)written
    # every time a download actually happens — NOT once at analysis time —
    # because 3 of the 18 scalar fields (IU_DebtLTVPct/InterestRatePct/
    # AmortizationYears) are debt assumptions typed in at download time, not
    # known until then, and can legitimately differ between repeat downloads
    # of the same analysis. ExcelOverrideUpload/-Delta below diff a
    # re-uploaded file's CURRENT named-range values against whatever this
    # field held at the time of that re-upload. Null for every analysis that
    # predates this feature (and for any analysis never actually downloaded
    # yet) — override tracking for those simply has nothing to compare
    # against, same "unavailable for legacy analyses" precedent as
    # rent_roll_data/t12_monthly_data above.
    injected_values_snapshot = models.JSONField(null=True, blank=True)

    # Sprint "Workflow Upgrades" Feature 1 (10-Second Scorecard & Missing
    # Doc Check): exactly which of the 3 core documents were actually
    # submitted for THIS run — {"om": bool, "t12": bool, "rent_roll": bool}.
    # Deliberately an explicit record of what was SUBMITTED, never inferred
    # from whether metrics happen to look empty — a document can be
    # provided and still yield little data (e.g. a T12 the AI couldn't
    # parse well), and that must never be confused with "wasn't sent at
    # all" for the Scorecard's missing-document check. Defaults to {}
    # (Django's required mutable-default pattern for a JSONField) for
    # every row that predates partial uploads — reader code
    # (_compute_deal_scorecard in views.py) treats an ABSENT key as "was
    # provided" (.get(key, True)), which is correct for every legacy row
    # too: all three were hard-required before this feature existed.
    documents_provided = models.JSONField(
        default=dict, blank=True,
    )

    # How long the actual analysis pipeline took (parsing + AI call + Enterprise
    # modules), in seconds. Real, measured data — used to power the "Avg.
    # Processing" dashboard stat honestly instead of a fabricated number.
    processing_seconds = models.FloatField(null=True, blank=True)

    # V2 Module 3 (Live Debt & Equity Waterfall Scenario Builder): the
    # deal's working financing/equity scenario — {"inputs": {...the last
    # submitted debt_structure/LTV/promote_tiers/etc...}, "results":
    # {...compute_waterfall()'s return value...}, "computed_at": iso
    # string}. Deliberately DOES persist here, unlike debt_assumptions on
    # PropertyProjectDownloadView (see that field's own comment above) —
    # that field is a one-off download-time input re-prompted every time
    # on purpose; this one is an explicit, named "Scenario Builder" the
    # analyst expects to leave and come back to, per the V2 spec's own
    # "persist scenario inputs in the Deal's working JSON state" line.
    # Null until the analyst runs "Calculate" at least once. See
    # waterfall_calculator.py for the actual math and its documented
    # modeling conventions, and ExcelBaselineLock.waterfall_summary below
    # for how this gets frozen at lock time.
    waterfall_scenario = models.JSONField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['organization', '-created_at']),
        ]

    # Bug fix (Nov 2026): every JSONField on this model is the output of
    # pandas/financial-math computations (metrics ratios, rent-roll/T12
    # aggregation, the waterfall calculator) rather than hand-typed data,
    # so any of them can carry a NaN/Infinity from some edge case (a zero
    # denominator, an empty rent roll) upstream. Listed explicitly rather
    # than discovered via introspection so adding a new JSONField to this
    # model is a deliberate, visible decision about whether it needs the
    # same protection -- not something that silently opts in or out.
    _SANITIZED_JSON_FIELDS = (
        'metrics', 'rent_roll_data', 't12_monthly_data',
        'injected_values_snapshot', 'documents_provided', 'waterfall_scenario',
    )

    def save(self, *args, **kwargs):
        """
        Sanitizes NaN/Infinity out of every JSONField above before handing
        off to the DB. Doing it here, once, at the model layer, means it's
        enforced no matter which call site saves an AnalysisReport (a brand
        new analysis via document_history.create_analysis_report, or the
        Live Grid Editor's partial `update_fields=[...]` save in views.py)
        and for any future call site too -- see _sanitize_json_floats'
        own docstring above for why this is needed at all (MySQL's JSON
        column type rejects NaN/Infinity outright, unlike Postgres).
        """
        for field_name in self._SANITIZED_JSON_FIELDS:
            value = getattr(self, field_name, None)
            if value is not None:
                setattr(self, field_name, _sanitize_json_floats(value))
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.property_name or 'Untitled'} ({self.status}) — {self.organization.name}"


class ExcelOverrideUpload(models.Model):
    """
    Feature 4 (Cell Override Audit Log): one row per "Re-upload Edited
    Underwriting File" event on a PropertyProject — catches an analyst
    editing a number locally in Excel AFTER downloading our generated
    workbook, before it goes any further (to a lender, to IC, etc.).

    `project` is denormalized alongside `analysis` (rather than only going
    through analysis.project) purely so every pipeline/history query can
    filter directly on project without an extra join — same reasoning as
    AnalysisReport.organization existing alongside AnalysisReport.project.

    The individual field-level changes detected on this specific re-upload
    live in ExcelOverrideDelta below, as children of this row — a grouping
    the UI needs to show "Alex's upload on Tuesday: 3 changes" as one card,
    not a flat list.
    """
    project = models.ForeignKey(
        PropertyProject, on_delete=models.CASCADE, related_name='override_uploads'
    )
    # The analysis this re-upload was diffed against (project.latest_analysis
    # at the moment of upload) — kept explicit rather than re-derived later,
    # since "latest" for the project can move on after this record is made.
    analysis = models.ForeignKey(
        AnalysisReport, on_delete=models.CASCADE, related_name='override_uploads'
    )
    uploaded_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='excel_override_uploads'
    )
    original_filename = models.CharField(max_length=255, blank=True, default='')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    # Feature 2 (Immutable Snapshots) extension: true when this project
    # already had an ExcelBaselineLock AT THE MOMENT this upload happened.
    # Set once, at creation, never recomputed later — whether a deal gets
    # locked AFTER this upload already happened shouldn't retroactively
    # relabel history. Purely a display hint ("this is a child version
    # branch on top of the locked baseline, not an edit to the baseline
    # itself") — it changes nothing about how the diff itself is computed;
    # see ExcelBaselineLock's own docstring for why the diff target
    # (analysis.injected_values_snapshot) is already frozen by that point
    # regardless of this flag.
    is_post_lock_revision = models.BooleanField(default=False)
    # Sprint "Workflow Upgrades" Feature 2 (Variance Alert Context Notes):
    # required leadership-facing 1-sentence reason, collected ONLY when this
    # upload's deltas include a qualifying >OVERRIDE_ALERT_THRESHOLD_PCT
    # shift on one of OVERRIDE_ALERT_TRACKED_FIELDS (GPR/OpEx/NOI) — see
    # ExcelOverrideCheckView.post()'s two-step flow in views.py. Blank for
    # every upload that never crossed that threshold, never backfilled or
    # required retroactively.
    variance_note = models.TextField(blank=True, default='')

    class Meta:
        ordering = ['-uploaded_at']
        indexes = [
            # Explicit name so this matches the migration's own AddIndex
            # exactly (0010_excel_override_audit.py hand-named it
            # 'override_upload_project_idx' for readability instead of
            # taking Django's auto-generated hash name) - otherwise
            # `makemigrations` sees the model state's auto-generated name
            # disagree with the migration's explicit one and proposes a
            # no-op RenameIndex migration every time it's run.
            models.Index(fields=['project', '-uploaded_at'], name='override_upload_project_idx'),
        ]

    def __str__(self):
        who = self.uploaded_by.email if self.uploaded_by else 'Unknown'
        return f"Override re-upload on {self.project} by {who}"


class ExcelOverrideDelta(models.Model):
    """
    Feature 4: ONE changed field detected on a single ExcelOverrideUpload —
    e.g. field_name='IU_PhysicalOccupancyPct', original_value='5.00%',
    new_value='8.00%', powering the UI's "Alex modified Projected Vacancy
    from 5% to 8%" line. A field whose re-uploaded value matches the
    snapshot exactly produces no row at all — this table only ever holds
    actual changes, never a full before/after of every tracked field.

    original_value/new_value are TextField, not JSONField — intentionally
    just the stringified display value (whatever
    compute_injected_values_snapshot / extract_defined_names produced),
    never recomputed from these rows. The authoritative source of "what
    SHOULD be here" stays AnalysisReport.injected_values_snapshot; this
    table is a read-oriented audit trail, not a second copy of that data.

    row_index is null for a scalar field (IU_PropertyName, etc.) and set to
    the tenant's position for a row field (IU_Row_BaseRent, etc.) — row
    matching is by POSITION only (same convention used when these values
    are originally written), not by any identity like tenant name, so an
    inserted/deleted row surfaces as a row-count mismatch rather than a
    smart re-match. Confirmed acceptable for v1.
    """
    upload = models.ForeignKey(
        ExcelOverrideUpload, on_delete=models.CASCADE, related_name='deltas'
    )
    field_name = models.CharField(max_length=100)
    row_index = models.IntegerField(null=True, blank=True)
    original_value = models.TextField(blank=True, default='')
    new_value = models.TextField(blank=True, default='')
    # Feature 1 (Variance & Change Ledger) extension: percentage change
    # between original_value and new_value, computed ONLY when both sides
    # were a real non-boolean number with a non-zero original — see
    # _variance_pct() in views.py. Null for every other change (a text/date
    # field like IU_PropertyName or IU_Row_LeaseStart, or a zero baseline a
    # percentage can't meaningfully describe) rather than a fabricated
    # number — same "never guess" discipline as everything else in this
    # audit log. Powers both the per-deal delta display and the firm-wide
    # override-alerts feed's >2.0% threshold check.
    variance_pct = models.FloatField(null=True, blank=True)

    class Meta:
        ordering = ['row_index', 'field_name']

    def __str__(self):
        target = f"{self.field_name}[{self.row_index}]" if self.row_index is not None else self.field_name
        return f"{target}: {self.original_value!r} -> {self.new_value!r}"


class ExcelBaselineLock(models.Model):
    """
    Feature 2 (Immutable Snapshots): "Approve & Lock Baseline" — once this
    row exists for a PropertyProject, that deal's current analysis
    snapshot is permanently frozen as its official underwriting baseline.

    Exactly what locking means, per the client's own explicit scope
    decision (narrow by design, not an accident):
      - AnalysisReport.injected_values_snapshot stops being overwritten on
        every subsequent Excel download — see the guard in
        _save_injected_values_snapshot() in views.py. A post-lock
        re-upload of an edited file is still diffed and recorded as usual
        (ExcelOverrideUpload/-Delta), it just diffs against the now-frozen
        snapshot forever, instead of a snapshot that could otherwise still
        legitimately move (see AnalysisReport.injected_values_snapshot's
        own docstring on why it's normally re-written every download).
      - The locked AnalysisReport can no longer be soft-deleted — see the
        guard in AnalysisReportDetailView's DELETE handler in views.py.
        (PropertyProject itself has no delete endpoint at all today — see
        that view's own comment — so this is the one real deletion path
        that could otherwise destroy the data a lock exists to protect.)
      - pipeline_stage and Team Notes are DELIBERATELY left fully editable —
        locking freezes the financial baseline, not the deal's workflow.
        The team can still move a locked deal through the pipeline funnel
        (New Incoming -> Underwriting -> VP Review -> Offer Submitted /
        Dead-Passed) and keep adding notes.
      - A post-lock Excel RE-UPLOAD that would otherwise need to replace
        the locked baseline instead becomes a separate child version
        branch (see ExcelOverrideUpload.is_post_lock_revision below) — the
        locked JSON snapshot itself is never touched, so the original
        baseline stays intact for audit history no matter what gets
        re-uploaded afterward.

    One lock per PropertyProject (OneToOneField) — a deliberately one-way,
    one-time action for v1: there is no "unlock" endpoint. Baseline
    integrity is the entire point of this feature, so undoing a lock is a
    real product decision to make deliberately later, not something to
    half-support now by guessing at a safe unlock story.

    `analysis` is on_delete=PROTECT, not CASCADE — the specific
    AnalysisReport whose snapshot got frozen must never be deletable out
    from under an existing lock; attempting to would raise ProtectedError
    rather than silently orphaning (or worse, cascading away) the audit
    record this entire feature exists to preserve. The project's own
    "latest" analysis can still move on past this one via re-analysis —
    this FK deliberately stays pointed at the ORIGINAL locked analysis,
    never implicitly re-pointed to a newer one.

    snapshot_sha256: per the client's explicit instruction, hashes ONLY
    analysis.injected_values_snapshot (the JSON field) — never the .xlsx
    file bytes and never the full AnalysisReport.metrics blob. Consistent
    with this platform's zero-retention policy for uploaded/generated
    Excel (see ClientExcelTemplate's docstring below for the one
    deliberate, unrelated exception to that policy) — a locked deal's
    Excel file itself is still never stored, only a hash of the JSON
    values it was generated from.
    """
    project = models.OneToOneField(
        PropertyProject, on_delete=models.CASCADE, related_name='baseline_lock'
    )
    analysis = models.ForeignKey(
        AnalysisReport, on_delete=models.PROTECT, related_name='baseline_locks'
    )
    locked_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='locked_baselines'
    )
    locked_at = models.DateTimeField(auto_now_add=True)
    snapshot_sha256 = models.CharField(max_length=64)

    # V2 Module 3: a FROZEN copy of compute_waterfall()'s result dict,
    # captured at the moment of lock (best-effort — see
    # _vault_locked_baseline_excel's own docstring for the identical
    # pattern this mirrors: a failure here must never block the lock
    # itself). Null if the analyst never ran the Scenario Builder before
    # locking. Deliberately a SEPARATE frozen copy, not a live pointer to
    # AnalysisReport.waterfall_scenario — that field keeps moving after
    # lock (the analyst can keep exploring new scenarios), exactly the
    # same "working state keeps moving, the lock is a snapshot in time"
    # split already established by injected_values_snapshot vs. this
    # model's own snapshot_sha256 above.
    waterfall_summary = models.JSONField(null=True, blank=True)

    def __str__(self):
        return f"Locked baseline for {self.project} @ {self.locked_at:%Y-%m-%d}"


class ClientExcelTemplate(models.Model):
    """
    A client-provided, proprietary Excel workbook (e.g. an institutional
    investor's own master valuation template) that generate_underwriting_excel()'s
    output is replaced with for this organization — see
    generate_excel_from_template() in custom_template_generator.py.

    Deliberately a DIFFERENT kind of stored artifact from every other
    upload on this platform: an OM/T12/Rent Roll (see ProcessedDocument's
    docstring) is explicitly NEVER stored, because it's analyzed once and
    discarded. A client's own master template is the opposite — a
    long-lived, reusable asset the client wants reused on every future
    analysis for their org — so it's deliberately kept on disk (see
    private_template_storage above) until replaced or removed. This is a
    named exception to the "we destroy your files" policy and should be
    reflected as such in the privacy policy / data retention documentation.

    One active template per Organization (OneToOneField) — if a client
    ever needs more than one layout (e.g. office vs. multifamily), that's
    a real product decision to make deliberately later, not something to
    half-support now by guessing which one applies.
    """
    organization = models.OneToOneField(
        'users.Organization',
        on_delete=models.CASCADE,
        related_name='excel_template',
    )
    file = models.FileField(upload_to=client_template_upload_path, storage=private_template_storage)
    original_filename = models.CharField(max_length=255)

    # Populated at upload time by custom_template_generator.extract_named_ranges()
    # — exactly which of our known IU_* field names this client's template
    # actually defines, and which are missing. Stored so the admin can see
    # this (and so the upload endpoint can report it back) WITHOUT
    # re-opening the workbook just to check. See also missing_required_fields
    # below for the subset that blocks the per-row table entirely.
    detected_named_ranges = models.JSONField(default=list, blank=True)

    uploaded_by = models.ForeignKey(
        'users.CustomUser',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='uploaded_excel_templates',
    )
    uploaded_at = models.DateTimeField(auto_now_add=True)

    # Lets an admin disable a template that turned out to be broken/stale
    # (e.g. the client's layout changed and a re-upload hasn't happened
    # yet) WITHOUT deleting the record and losing detected_named_ranges /
    # upload history. An inactive template is treated as "no template" by
    # the download view — it falls back to the standard generator.
    is_active = models.BooleanField(default=True)

    class Meta:
        indexes = [
            models.Index(fields=['organization']),
        ]

    def __str__(self):
        return f"{self.organization.name} custom template ({self.original_filename})"


@receiver(post_delete, sender=ClientExcelTemplate)
def _delete_clientexceltemplate_file_from_storage(sender, instance, **kwargs):
    """
    Belt-and-suspenders file cleanup, independent of
    AdminDeleteExcelTemplateView (which already deletes the file itself
    before calling .delete()). Django never deletes a FileField's actual
    file when a row is deleted -- only the DB row -- so WITHOUT this
    signal, a ClientExcelTemplate removed any other way (an organization
    deleted via AdminDeleteOrganizationView, which CASCADEs through this
    OneToOneField; a queryset .delete(); the Django admin; a shell
    session) would silently leave the client's proprietary template file
    sitting on the VPS forever. Caught by testing the admin endpoints'
    replace-then-check-for-orphans path directly, not assumed safe.

    `if instance.file` is False when AdminDeleteExcelTemplateView already
    cleared it first, so this never double-deletes.
    """
    if instance.file:
        instance.file.delete(save=False)


class DealChecklistItem(models.Model):
    """
    Pipeline Modules Phase 4 (Persistent Missing-Doc Checklist): one row
    per document (or custom item) the team is tracking as outstanding on
    a deal. The 3 core documents are seeded automatically whenever an
    analysis run reports one missing — see record_analysis_report() in
    document_history.py — and the team can add anything else by hand
    (e.g. "Updated insurance quote", "Signed LOI") on top of that.

    doc_key is set ONLY for the 3 auto-seeded items ('om'/'t12'/
    'rent_roll') — it's what record_analysis_report() checks against to
    decide whether this project has already been seeded for that
    document, so a resolved item is never silently re-created just
    because a later re-analysis still reports that same document
    missing. A custom, hand-added item has doc_key=None and is never
    auto-seeded or auto-resolved — the team manages its lifecycle
    entirely by hand.

    is_resolved deliberately has no "auto-unresolve" path: once the team
    marks an item resolved (e.g. "got the T12 from the broker over
    email, outside this platform"), it stays resolved even if a later
    re-analysis still reports that document missing — per the client's
    explicit decision, this checklist tracks the team's own ground
    truth, not a live mirror of AnalysisReport.documents_provided.
    """
    DOC_KEY_CHOICES = [
        ('om', 'Offering Memorandum (OM)'),
        ('t12', 'T12 Financial Statement'),
        ('rent_roll', 'Rent Roll'),
    ]

    project = models.ForeignKey(
        PropertyProject, on_delete=models.CASCADE, related_name='checklist_items'
    )
    # Null for a custom, hand-added item — see this model's own docstring.
    doc_key = models.CharField(max_length=20, choices=DOC_KEY_CHOICES, null=True, blank=True)
    label = models.CharField(max_length=255)
    is_resolved = models.BooleanField(default=False)
    created_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='created_checklist_items'
    )
    resolved_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='resolved_checklist_items'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    resolved_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['is_resolved', 'created_at']
        indexes = [
            models.Index(fields=['project', 'is_resolved'], name='checklist_project_resolved_idx'),
        ]

    def __str__(self):
        return f"[{'x' if self.is_resolved else ' '}] {self.label} ({self.project})"


class DealVaultFile(models.Model):
    """
    Pipeline Modules Phase 5 (Excel Versioning & Supporting Document
    Vault): one row per versioned file kept against a PropertyProject —
    either an Excel snapshot this system vaulted automatically, or a
    supporting document a team member chose to attach by hand.

    Deliberately NOT a general-purpose "attach anything, anytime"
    uploader — the client's explicit v1 scope is auto-versioning ONLY at
    two moments (see file_type below), plus optional manual supporting
    documents. There is no "save version now" button anywhere in the UI.

    file_type:
      - 'locked_baseline_excel': captured automatically the instant a
        deal's baseline is locked (see _create_baseline_lock() in
        views.py). The workbook is REGENERATED at that moment from the
        locked analysis's metrics/rent_roll_data plus this org's CURRENT
        debt assumptions (Organization.get_debt_assumptions()) — never
        pulled from any specific earlier download. Debt inputs like
        purchase price/LTV/rate are never persisted on an analysis (see
        AnalysisReport's own docstring on why), so this vaulted copy may
        not be byte-identical to any one historical download — it's the
        best-effort "what the locked numbers look like as a workbook"
        snapshot, not a replay of a specific past request. Deliberately
        SKIPPED for organizations using a ClientExcelTemplate (custom
        Excel template) in this v1, to avoid the more complex
        custom_template_generator.py branching — those orgs simply get
        no auto-vaulted Excel on lock (see _create_baseline_lock()'s own
        try/except around this, which never blocks the lock itself on
        any failure, custom-template or otherwise).
      - 're_upload': captured automatically, byte-exact, every time an
        analyst successfully re-uploads an edited workbook through "Re-
        upload Edited Underwriting File" (see ExcelOverrideCheckView) —
        no regeneration needed here, the request already has the exact
        bytes in hand.
      - 'supporting_document': added by hand through
        DealVaultFileUploadView — anything else the team wants kept
        against this deal (an appraisal, an insurance quote, a signed
        LOI). This platform's zero-retention policy (see
        ProcessedDocument's docstring) does NOT apply to this file type —
        a document a team member deliberately chooses to vault here is a
        long-lived record by intent, the same named exception already
        established for ClientExcelTemplate above.

    Uses the SAME private-storage discipline as ClientExcelTemplate —
    never a public URL, read back only server-side through an
    authenticated, visibility-checked download view.
    """
    FILE_TYPE_CHOICES = [
        ('locked_baseline_excel', 'Locked Baseline Excel'),
        ('re_upload', 'Re-Uploaded Excel'),
        ('supporting_document', 'Supporting Document'),
        # Pipeline Modules Phase 8 (Live In-Browser Spreadsheet Editor):
        # captured automatically every time an analyst saves an edit
        # through the Live Grid Editor (see GridEditSaveView in
        # views.py). Deliberately its OWN file_type, not lumped in with
        # 're_upload' — the vaulted bytes here are a MINIMAL synthetic
        # workbook built only to carry the grid's edited values through
        # the shared override-diff machinery (_build_grid_edit_workbook),
        # not a real formatted re-upload a human edited in Excel. Labeling
        # it accurately keeps the Vault's download list from implying
        # it's the deliverable workbook when it never was.
        ('live_grid_edit', 'Live Grid Edit'),
    ]

    project = models.ForeignKey(
        PropertyProject, on_delete=models.CASCADE, related_name='vault_files'
    )
    file = models.FileField(upload_to=deal_vault_upload_path, storage=deal_vault_storage)
    file_type = models.CharField(max_length=24, choices=FILE_TYPE_CHOICES)
    original_filename = models.CharField(max_length=255)
    # A short human label distinguishing versions in the UI — e.g.
    # "Locked 2026-11-16", "Re-upload — Nov 16", or whatever filename-ish
    # text fits a supporting document. Never used for anything but
    # display.
    version_label = models.CharField(max_length=255, blank=True, default='')
    uploaded_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='vaulted_deal_files'
    )
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-uploaded_at']
        indexes = [
            models.Index(fields=['project', '-uploaded_at'], name='deal_vault_project_idx'),
        ]

    def __str__(self):
        return f"{self.get_file_type_display()} on {self.project} ({self.original_filename})"


@receiver(post_delete, sender=DealVaultFile)
def _delete_dealvaultfile_file_from_storage(sender, instance, **kwargs):
    """
    Same belt-and-suspenders file cleanup as
    _delete_clientexceltemplate_file_from_storage above, for the exact
    same reason — Django never deletes a FileField's actual file when a
    row is deleted, only the DB row.
    """
    if instance.file:
        instance.file.delete(save=False)


def _generate_lender_share_token():
    """
    Pipeline Modules Phase 6 (One-Click Lender Submission Packet & View
    Tracking): an opaque, unguessable token identifying one
    LenderShareLink's public portal URL (/share/deal/<token> on the
    frontend) — deliberately NOT sequential, since this is the one and
    only thing standing between "read-only lender packet" and "anyone on
    the internet can browse every locked deal this firm has ever done."
    A module-level function (not a bound default), for the identical
    migration-portability reason given above deal_vault_upload_path.
    """
    return secrets.token_urlsafe(32)


class LenderShareLink(models.Model):
    """
    Pipeline Modules Phase 6 (One-Click Lender Submission Packet & View
    Tracking): a secure, read-only, UNAUTHENTICATED public portal link for
    one locked deal — see LenderSharePortalView in views.py for exactly
    what it bundles (the Deal Scorecard/Executive Summary, the Locked
    Underwritten Excel Model, and manually-attached Document Vault
    'supporting_document' files — never raw OM/T12/Rent Roll inputs,
    which this platform's zero-retention policy never persists anywhere
    to begin with).

    Deliberately does NOT write view events into DealActivityLog — that
    model is explicitly scoped to pipeline_stage transitions only (see its
    own docstring). A lender opening this link has no stage. Instead,
    view_count/last_viewed_at below ARE the view-tracking mechanism, and
    PropertyProjectDetailView's activity feed (_get_deal_activity_feed)
    surfaces them back to the team as a 'lender_link_viewed' event —
    one feed entry per link that's been viewed at least once, not a
    per-visit log (this model keeps a running count, not a timestamped
    history of every individual view).

    access_token: see _generate_lender_share_token above. Never logged,
    never exposed anywhere except to the analyst who generated it (to
    copy/send themselves) and the lender who was sent it.

    is_active is a manual revoke/reactivate switch, independent of
    expires_at (a time-based auto-expiry set at generation — see
    LenderShareLinkListCreateView's default-30-days/override/no-expiry
    options). The public portal view (and is_usable below) treats EITHER
    is_active=False OR an expired expires_at identically — "this link no
    longer works" — and returns 404, deliberately never 401: this
    platform's global $fetch interceptor
    (frontend/app/plugins/auth_interceptor.client.ts) force-logs-out and
    redirects to /login on ANY 401 response anywhere in the app, which
    would strand an external lender who has no account to log into.

    Generating a link requires the deal to already be locked — there's
    nothing frozen to show a lender otherwise, same "nothing to show
    before lock" principle as ICSummaryPDFView/_vault_locked_baseline_excel.

    workspace is a denormalized copy of project.workspace AT GENERATION
    TIME, for firm-side filtering/display only — never re-synced if the
    deal's workspace changes later, and never consulted by the public
    portal view itself (which only ever needs `project`).
    """
    project = models.ForeignKey(
        PropertyProject, on_delete=models.CASCADE, related_name='lender_share_links'
    )
    workspace = models.ForeignKey(
        'users.TeamWorkspace', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='lender_share_links'
    )
    access_token = models.CharField(
        max_length=64, unique=True, default=_generate_lender_share_token, editable=False
    )
    recipient_label = models.CharField(max_length=255, blank=True, default='')
    is_active = models.BooleanField(default=True)
    view_count = models.PositiveIntegerField(default=0)
    last_viewed_at = models.DateTimeField(null=True, blank=True)
    expires_at = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='created_lender_share_links'
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['project', '-created_at'], name='lender_share_project_idx'),
        ]

    def __str__(self):
        return f"Lender link for {self.project} ({'active' if self.is_active else 'revoked'})"

    @property
    def is_expired(self):
        return self.expires_at is not None and self.expires_at <= timezone.now()

    @property
    def is_usable(self):
        """
        The single source of truth both LenderShareLinkListCreateView (for
        display) and the public portal views (for the actual gate) check —
        so "is this link still good" is never independently re-derived in
        two places that could silently drift apart.
        """
        return self.is_active and not self.is_expired


class InboundEmailEvent(models.Model):
    """
    Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): one row
    per inbound email webhook delivery — accepted or rejected. Serves TWO
    purposes at once:

      1. Idempotency — provider_message_id is the ENTIRE dedup mechanism.
         Postmark (like every inbound-email provider) retries a webhook
         delivery that times out or returns a non-2xx response; without
         this unique constraint, a single forwarded email could create
         duplicate PropertyProjects on retry. See
         InboundEmailWebhookView's own docstring for exactly how this is
         checked BEFORE any processing happens.
      2. Audit trail — the client's spec explicitly calls for rejected/
         unauthorized senders to be logged, not just silently dropped.
         Every delivery gets a row here, so "what came in and what
         happened to it" is always answerable without grepping server
         logs — including deliveries that never created a deal at all
         (unknown workspace slug, a sender who isn't a member of that
         workspace, or no usable OM/T12/Rent Roll attachment found).
    """
    STATUS_CHOICES = [
        ('accepted', 'Accepted'),
        ('rejected_unknown_slug', 'Rejected — Unknown Workspace Slug'),
        ('rejected_unauthorized_sender', 'Rejected — Unauthorized Sender'),
        ('rejected_no_usable_attachments', 'Rejected — No Usable Attachments'),
        ('rejected_malformed', 'Rejected — Malformed Payload'),
    ]

    provider_message_id = models.CharField(max_length=255, unique=True)
    workspace = models.ForeignKey(
        'users.TeamWorkspace', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='inbound_email_events'
    )
    sender_email = models.EmailField(blank=True, default='')
    subject = models.CharField(max_length=500, blank=True, default='')
    status = models.CharField(max_length=32, choices=STATUS_CHOICES)
    # Null for every rejected status above — only an 'accepted' delivery
    # ever creates (and links to) a PropertyProject.
    project = models.ForeignKey(
        PropertyProject, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='inbound_email_events'
    )
    received_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-received_at']
        indexes = [
            models.Index(fields=['workspace', '-received_at'], name='inbound_email_workspace_idx'),
        ]

    def __str__(self):
        return f"{self.status}: {self.sender_email or 'unknown sender'} @ {self.received_at:%Y-%m-%d %H:%M}"


class PropertyMarketDataSnapshot(models.Model):
    """
    V2 Module 2 (Tax & Rent Comps Panel via RentCast): the CACHE that
    makes this feature safe to run on RentCast's free 50-request/month
    tier, per the client's explicit instruction — "cache the RentCast
    response on the deal record the first time an address is checked so
    repeated opens never burn extra lookups." OneToOne with
    PropertyProject (not AnalysisReport): the address being checked is a
    property of the DEAL, not of any one analysis run, and re-analyzing
    a deal must never trigger a second, redundant RentCast lookup for
    the same address.

    Stores ONLY the raw, expensive-to-fetch RentCast responses (see
    rentcast_client.py) plus a few directly-extracted scalars for quick
    display — never the derived variance flags themselves, which are
    always recomputed live from these raw fields plus the project's
    CURRENT metrics (see market_comps.build_market_comps_payload) so a
    later Excel re-upload or grid edit is compared against up-to-date
    seller-stated figures without ever needing to re-fetch RentCast.

    Created once, on the first successful (or attempted) lookup for a
    project (see MarketCompsView in views.py) — there is no "expires
    after N days" auto-refresh logic; see address_checked's own
    docstring for the one deliberate, explicitly-gated exception.
    """
    project = models.OneToOneField(
        'PropertyProject', on_delete=models.CASCADE, related_name='market_data_snapshot'
    )

    # The exact address string sent to RentCast — kept even though it's
    # presumably also sitting in metrics.property_metadata.address at
    # fetch time, because that metrics value can keep changing after
    # this snapshot is cached (a grid edit, a re-upload) while this
    # field records what was ACTUALLY looked up, for an honest "last
    # checked against this address" display if the two ever drift apart.
    address_checked = models.CharField(max_length=500)

    fetched_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='market_data_lookups'
    )
    # Deliberately NOT auto_now_add — a gated "refresh" (see MarketCompsView)
    # updates this same row rather than creating a new one, and that re-fetch
    # must move this timestamp forward too. Set explicitly by the view on
    # every real RentCast fetch (both first-create and refresh).
    fetched_at = models.DateTimeField(default=timezone.now)

    # Raw RentCast response bodies (see rentcast_client.fetch_all_market_data)
    # — never transformed here; see market_comps.py for all derived reads.
    raw_property_record = models.JSONField(null=True, blank=True)
    raw_avm_value = models.JSONField(null=True, blank=True)
    raw_rental_comps = models.JSONField(null=True, blank=True)

    # Set when fetch_all_market_data() raised (e.g. RentCast misconfigured,
    # network failure, or every one of its 3 sub-calls failed) — a row
    # with this set still counts as "checked" for caching purposes
    # (see RE-FETCH note below): a RentCast outage on first check must
    # never silently retry on every subsequent drawer open and burn the
    # monthly quota on a vendor that's currently down anyway. Analysts
    # see this as a "Market comps lookup failed" state with the message
    # below, and a Workspace Head/Org Admin can explicitly force a
    # re-check (see MarketCompsView's refresh= handling) once ready.
    fetch_error = models.TextField(blank=True, default='')

    class Meta:
        indexes = [
            models.Index(fields=['project'], name='market_snapshot_project_idx'),
        ]

    def __str__(self):
        return f"Market data for {self.project} @ {self.address_checked!r}"


class Notification(models.Model):
    """
    V2 Module 4a (Buy-Box Auto-Alerts): the FIRST in-app notification
    infrastructure in this codebase. Pipeline Modules Phase 4's
    TeamNote.mentioned_users (above) explicitly stayed in-app-badge-only
    with NO backing table of its own, by deliberate client decision, for
    exactly this reason — no notification infrastructure existed yet.
    This table is that infrastructure's generic home: NOTIFICATION_TYPE_
    CHOICES has exactly one value today (buy_box_match — see
    buy_box.check_buy_box_match / document_history._check_buy_box_match_
    and_notify), but the shape (recipient, type, optional project link,
    message, is_read, created_at) is deliberately generic so a future
    notification source can reuse this table rather than inventing its
    own — @mentions were explicitly called out as a candidate the last
    time this gap was discussed.
    """
    NOTIFICATION_TYPE_CHOICES = [
        ('buy_box_match', 'Buy-Box Match'),
        # V2 Module 1 (Phase V2-3): a connected mailbox's refresh token
        # stopped working (provider-side revocation, password change,
        # admin-revoked app access, or this app's own encryption key
        # rotating — see EmailOAuthConnection.status in this same file)
        # — surfaced here exactly like a buy-box match, rather than
        # silently leaving the Settings page as the only place a user
        # would ever notice their sync quietly stopped.
        ('email_sync_error', 'Email Sync Needs Reconnect'),
    ]

    recipient = models.ForeignKey(
        'users.CustomUser', on_delete=models.CASCADE, related_name='notifications'
    )
    notification_type = models.CharField(max_length=30, choices=NOTIFICATION_TYPE_CHOICES)
    # Nullable: generic infra, and not every future notification type will
    # necessarily be about one specific deal. Every notification today
    # (buy_box_match) always sets this.
    project = models.ForeignKey(
        'PropertyProject', on_delete=models.CASCADE, null=True, blank=True, related_name='notifications'
    )
    message = models.TextField()
    is_read = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['recipient', 'is_read', '-created_at'], name='notification_recipient_idx'),
        ]

    def __str__(self):
        return f"[{self.notification_type}] to {self.recipient_id}: {self.message[:60]}"


# ---------------------------------------------------------------------------
# V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3)
# ---------------------------------------------------------------------------
# Three models, one per layer of this feature:
#   EmailOAuthConnection  - one CustomUser's connected Gmail or Microsoft 365
#                           mailbox (at most one of EACH provider at a time).
#   LinkedEmailThread     - one existing thread in that mailbox, explicitly
#                           "tagged" by an analyst to a specific deal (per
#                           the client's own wording: "deal-tagged threads" -
#                           this is deliberately NEVER auto-matched by
#                           counterparty address; see email_sync.py).
#   SyncedEmailMessage    - one message the 5-minute poll (see the
#                           sync_email_threads management command) pulled
#                           from a linked thread into this deal's Unified
#                           Activity Feed (_get_deal_activity_feed in
#                           views.py gets a new 'email' event type sourced
#                           from this table).
#
# Zero-retention boundary, stated precisely because it's easy to blur: the
# classify-then-discard policy applies to ATTACHMENTS classified as OM/T12/
# Rent Roll ONLY (see email_sync.py's own docstring) - raw bytes for those
# three types are never written anywhere, not even transiently to disk.
# SyncedEmailMessage.body_text, by contrast, is an ordinary persisted
# record of the correspondence itself (the entire point of "sync into the
# Activity Feed") - no different in kind from TeamNote.body, just sourced
# from an external mailbox instead of typed in-app.


class EmailOAuthConnection(models.Model):
    """
    One CustomUser's OAuth grant to their own Gmail or Microsoft 365
    mailbox, read-only (see email_oauth.py for the exact scopes: Gmail
    `gmail.readonly`, Microsoft Graph `Mail.Read` - this feature only
    ever READS a connected mailbox; it never sends mail through it).
    `unique_together` with provider - a user can connect at most ONE
    Gmail AND at most ONE Microsoft 365 account at the same time, never
    two of the same provider (reconnecting the same provider replaces
    the existing row rather than creating a second one - see
    EmailConnectView in views.py).

    access_token/refresh_token are stored ENCRYPTED (see token_crypto.py
    and set_access_token()/get_access_token() etc. below) - never as
    plaintext columns, per the client's explicit "encrypted token
    storage" requirement. refresh_token in particular is long-lived and
    is what makes the 5-minute poll possible without the user re-
    authenticating; losing it (provider-side revocation, or this app's
    own encryption key rotating - see token_crypto.py) moves `status`
    to 'error', which the Settings page and sync_email_threads both
    treat as "needs reconnect," never a crash.
    """
    PROVIDER_CHOICES = [
        ('google', 'Google (Gmail)'),
        ('microsoft', 'Microsoft 365'),
    ]
    STATUS_CHOICES = [
        ('active', 'Active'),
        ('error', 'Needs Reconnect'),
        ('disconnected', 'Disconnected'),
    ]

    user = models.ForeignKey(
        'users.CustomUser', on_delete=models.CASCADE, related_name='email_oauth_connections'
    )
    provider = models.CharField(max_length=10, choices=PROVIDER_CHOICES)
    # The connected mailbox's own address (from the provider's userinfo/
    # profile endpoint at connect time) - shown in Settings so an analyst
    # can tell at a glance which inbox is linked, and used nowhere in any
    # security check (that's entirely `user`/`status` below).
    email_address = models.EmailField()
    encrypted_access_token = models.TextField(blank=True, default='')
    encrypted_refresh_token = models.TextField(blank=True, default='')
    access_token_expires_at = models.DateTimeField(null=True, blank=True)
    # Space-separated scope string exactly as the provider granted it -
    # kept for audit/debugging only, never parsed to gate anything.
    scopes = models.TextField(blank=True, default='')
    status = models.CharField(max_length=15, choices=STATUS_CHOICES, default='active')
    last_error = models.TextField(blank=True, default='')
    connected_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['user', 'provider'], name='unique_oauth_connection_per_user_provider'),
        ]
        ordering = ['-connected_at']

    def __str__(self):
        return f"{self.get_provider_display()} ({self.email_address}) for user {self.user_id} [{self.status}]"

    def set_access_token(self, plaintext: str, expires_at) -> None:
        from .token_crypto import encrypt_token
        self.encrypted_access_token = encrypt_token(plaintext) if plaintext else ''
        self.access_token_expires_at = expires_at

    def set_refresh_token(self, plaintext: str) -> None:
        # Google only resends a refresh_token on the VERY FIRST consent
        # (or when prompt=consent forces a new one); a bare token refresh
        # response omits it entirely. Microsoft's refresh response MAY
        # include a new one it wants the old one discarded in favor of.
        # Either way: never overwrite a real stored refresh_token with a
        # blank one just because this particular response didn't repeat
        # it - only write when the caller actually has a new value.
        if plaintext:
            from .token_crypto import encrypt_token
            self.encrypted_refresh_token = encrypt_token(plaintext)

    def get_access_token(self) -> str:
        from .token_crypto import decrypt_token
        return decrypt_token(self.encrypted_access_token)

    def get_refresh_token(self) -> str:
        from .token_crypto import decrypt_token
        return decrypt_token(self.encrypted_refresh_token)


class LinkedEmailThread(models.Model):
    """
    One existing thread in a connected mailbox, explicitly linked
    ("tagged") to one deal by an analyst — see EmailThreadSearchView/
    ProjectEmailThreadsView in views.py for the search-then-link UI this
    backs. provider_thread_id is Gmail's `threadId` for a 'google'
    connection or Microsoft Graph's `conversationId` for a 'microsoft'
    one — opaque past that; never parsed or assumed to mean anything.

    Deliberately NEVER auto-matched by counterparty email address or
    subject heuristics — this app has no reliable, structured "broker
    email" / "lender email" field on PropertyProject to match against,
    and silently pulling mailbox content into a shared company record
    without an explicit, visible, per-thread action is exactly the kind
    of surprising behavior this codebase avoids elsewhere (see e.g.
    MarketCompsView never auto-refreshing from RentCast). `label` is a
    short free-text tag (e.g. "Broker — Jane @ CBRE", "Lender — Chase")
    for the Activity Feed to show next to synced messages — deliberately
    NOT a hard enum: this app already prefers free text over rigid
    categories wherever the exact label doesn't drive any logic (see
    e.g. PropertyProject's own free-text asset_class).

    is_active=False SOFT-unlinks a thread (see
    ProjectEmailThreadsView.delete()) — every SyncedEmailMessage already
    pulled in stays in the Activity Feed as history; only the five-
    minute poll stops touching it. There's no hard delete endpoint.
    """
    BACKFILL_STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('complete', 'Complete'),
        ('failed', 'Failed'),
    ]

    project = models.ForeignKey(
        'PropertyProject', on_delete=models.CASCADE, related_name='linked_email_threads'
    )
    connection = models.ForeignKey(
        EmailOAuthConnection, on_delete=models.CASCADE, related_name='linked_threads'
    )
    provider_thread_id = models.CharField(max_length=255)
    subject = models.CharField(max_length=500, blank=True, default='')
    label = models.CharField(max_length=120, blank=True, default='')
    linked_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='linked_email_threads'
    )
    linked_at = models.DateTimeField(auto_now_add=True)
    is_active = models.BooleanField(default=True)
    last_synced_at = models.DateTimeField(null=True, blank=True)
    # Per the client's explicit "backfill full history on link" decision
    # (see email_sync.backfill_linked_thread), set once the initial pull
    # of every pre-existing message in the thread finishes — the 5-minute
    # poll skips any thread still 'pending' (its backfill task hasn't
    # run yet) rather than racing it, and surfaces 'failed' plainly in
    # the deal drawer rather than silently retrying forever.
    backfill_status = models.CharField(max_length=15, choices=BACKFILL_STATUS_CHOICES, default='pending')
    backfill_error = models.TextField(blank=True, default='')

    class Meta:
        ordering = ['-linked_at']
        indexes = [
            models.Index(fields=['project', 'is_active'], name='linked_thread_project_idx'),
            models.Index(fields=['connection', 'is_active'], name='linked_thread_connection_idx'),
        ]

    def __str__(self):
        return f"{self.label or self.subject or self.provider_thread_id} on project {self.project_id}"


class SyncedEmailMessage(models.Model):
    """
    One message the 5-minute poll (sync_email_threads) pulled from a
    LinkedEmailThread — the 'email' event type _get_deal_activity_feed()
    in views.py renders. `provider_message_id` + `thread_link` is unique
    so a retried/overlapping poll run can never double-insert the same
    message (get_or_create in email_sync.py, not a pre-check — see that
    module's own docstring on why).

    body_text is the message's PLAIN-TEXT body (never raw HTML — kept
    simple and never rendered as markup), capped at email_sync.py's
    MAX_SYNCED_BODY_CHARS. This is an ordinary persisted record of the
    correspondence, NOT covered by this platform's attachment zero-
    retention policy (see this section's own module-level docstring
    above) — attachment_summary below is where that policy shows up.

    attachment_summary is a list of {filename, classified_as, action}
    dicts, one per attachment this message had:
      - classified_as: 'om' | 't12' | 'rent_roll' | 'other' | 'unclassified'
      - action: 'discarded' (om/t12/rent_roll — raw bytes were never
        written anywhere, classified in memory and dropped) or
        'stored_in_vault' (anything else — saved as a normal
        DealVaultFile('supporting_document'), the same long-lived-by-
        intent exception that type already carries — see its own
        docstring) or 'skipped_too_large'/'skipped_error' for the rare
        attachment this couldn't even classify.
    """
    DIRECTION_CHOICES = [
        ('inbound', 'Inbound'),
        ('outbound', 'Outbound'),
    ]

    thread_link = models.ForeignKey(LinkedEmailThread, on_delete=models.CASCADE, related_name='messages')
    provider_message_id = models.CharField(max_length=255)
    sent_at = models.DateTimeField()
    sender_email = models.EmailField(blank=True, default='')
    sender_name = models.CharField(max_length=255, blank=True, default='')
    direction = models.CharField(max_length=10, choices=DIRECTION_CHOICES)
    body_text = models.TextField(blank=True, default='')
    attachment_summary = models.JSONField(default=list, blank=True)
    # The DealVaultFile rows (if any) a 'stored_in_vault' attachment on
    # this message created — lets the Activity Feed entry deep-link to
    # the Vault rather than just naming the file in attachment_summary.
    vault_file_ids = models.JSONField(default=list, blank=True)
    synced_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['thread_link', 'provider_message_id'], name='unique_synced_message_per_thread'),
        ]
        ordering = ['sent_at']
        indexes = [
            models.Index(fields=['thread_link', '-sent_at'], name='synced_msg_thread_idx'),
        ]

    def __str__(self):
        return f"{self.direction} message on thread {self.thread_link_id} at {self.sent_at}"
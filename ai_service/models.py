# backend/ai_service/models.py
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.db import models
from django.db.models.signals import post_delete
from django.dispatch import receiver


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
            "ClientExcelTemplate files have no public URL by design -- "
            "read them server-side via storage.path()/.open(), from an "
            "authenticated admin view, never through a public URL."
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
    iterative, re-analyzed as negotiations progress, and deal_status/Team
    Notes need to survive that re-analysis rather than resetting with every
    new run. See AnalysisReport.project below for how the two connect.

    Every PropertyProject existed in some form before this field did, too:
    the data migration that introduces this table also auto-wraps every
    pre-existing AnalysisReport into its own single-analysis
    PropertyProject, by deliberate decision — no attempt is made to guess
    which historical analyses were actually re-runs of the same deal.
    """
    STATUS_CHOICES = [
        ('reviewing', 'Reviewing'),
        ('approved', 'Approved'),
        ('archived', 'Archived'),
    ]

    organization = models.ForeignKey(
        'users.Organization', on_delete=models.CASCADE, related_name='property_projects'
    )
    workspace = models.ForeignKey(
        'users.TeamWorkspace', on_delete=models.SET_NULL, null=True, blank=True, related_name='property_projects'
    )

    property_name = models.CharField(max_length=255, blank=True, default='')
    deal_status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='reviewing')

    created_by = models.ForeignKey(
        'users.CustomUser', on_delete=models.SET_NULL, null=True, blank=True, related_name='created_projects'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    is_deleted = models.BooleanField(default=False)

    class Meta:
        ordering = ['-updated_at']
        indexes = [
            models.Index(fields=['organization', '-updated_at']),
            models.Index(fields=['workspace', '-updated_at']),
        ]

    def __str__(self):
        return f"{self.property_name or 'Untitled'} [{self.deal_status}] — {self.organization.name}"

    @property
    def latest_analysis(self):
        return self.analyses.order_by('-created_at').first()


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

    # How long the actual analysis pipeline took (parsing + AI call + Enterprise
    # modules), in seconds. Real, measured data — used to power the "Avg.
    # Processing" dashboard stat honestly instead of a fabricated number.
    processing_seconds = models.FloatField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['organization', '-created_at']),
        ]

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
      - deal_status and Team Notes are DELIBERATELY left fully editable —
        locking freezes the financial baseline, not the deal's workflow.
        The team can still move a locked deal through Reviewing ->
        Approved -> Archived and keep adding notes.
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
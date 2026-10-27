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


class AnalysisReport(models.Model):
    """
    Stores the RESULT of a completed analysis (the metrics JSON) so a user can
    come back and review it later — this is the "Outputs" page. Also doubles
    as the source of truth for upload-quota counting (see views.py) — one row
    per successful analysis, so counting rows in a time window == counting
    "uploads used", with no separate counter field needed.

    Deliberately does NOT store the generated Excel file — only Preview is
    supported on the Outputs page (no Download), so there's no need for file
    storage infrastructure here, just the JSON result itself.

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

    is_deleted = models.BooleanField(default=False)
    property_name = models.CharField(max_length=255, blank=True, default='')
    tier = models.CharField(max_length=20)
    document_type = models.CharField(max_length=10, choices=DOCUMENT_TYPE_CHOICES, default='CRE')
    metrics = models.JSONField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='ready')

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
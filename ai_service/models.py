# backend/ai_service/models.py
from django.db import models

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
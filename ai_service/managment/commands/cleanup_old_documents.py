# backend/ai_service/management/commands/cleanup_old_documents.py
"""
Deletes ProcessedDocument history records older than the retention period.

Run manually or via cron for now:
    python manage.py cleanup_old_documents

No Celery Beat scheduling yet (background task infrastructure hasn't been
built — see the scaling conversation). Once it exists, this becomes a
scheduled periodic task instead of a manually-triggered command.

Suggested cron entry (daily at 3am):
    0 3 * * * cd /path/to/backend && /path/to/venv/bin/python manage.py cleanup_old_documents
"""
from datetime import timedelta
from django.core.management.base import BaseCommand
from django.utils import timezone

RETENTION_DAYS = 30


class Command(BaseCommand):
    help = f"Deletes document history records older than {RETENTION_DAYS} days."

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help="Show what would be deleted without actually deleting anything.",
        )

    def handle(self, *args, **options):
        from ai_service.models import ProcessedDocument

        cutoff = timezone.now() - timedelta(days=RETENTION_DAYS)
        queryset = ProcessedDocument.objects.filter(uploaded_at__lt=cutoff)
        count = queryset.count()

        if options['dry_run']:
            self.stdout.write(f"[DRY RUN] Would delete {count} record(s) older than {RETENTION_DAYS} days.")
            return

        queryset.delete()
        self.stdout.write(self.style.SUCCESS(f"Deleted {count} record(s) older than {RETENTION_DAYS} days."))
# backend/users/management/commands/cleanup_expired_trials.py
"""
Hard-deletes trial Organizations (and their user, via cascade) once
trial_expires_at has passed — regardless of whether the trial's one
analysis was ever used.

Run manually or via cron:
    python manage.py cleanup_expired_trials

Given the short (24h) lifetime, this needs to run much more frequently than
the 30-day document retention job — suggested cron entry (every hour):
    0 * * * * cd /path/to/backend && /path/to/venv/bin/python manage.py cleanup_expired_trials

No Celery Beat scheduling yet — same caveat as cleanup_old_documents.
"""
from django.core.management.base import BaseCommand
from django.utils import timezone


class Command(BaseCommand):
    help = "Deletes trial organizations (and their user) whose trial_expires_at has passed."

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help="Show what would be deleted without actually deleting anything.",
        )

    def handle(self, *args, **options):
        from users.models import Organization

        now = timezone.now()
        queryset = Organization.objects.filter(is_trial=True, trial_expires_at__lte=now)
        count = queryset.count()

        if options['dry_run']:
            for org in queryset:
                self.stdout.write(f"[DRY RUN] Would delete: {org.name} (expired {org.trial_expires_at})")
            self.stdout.write(f"[DRY RUN] Would delete {count} expired trial account(s).")
            return

        # Deleting the Organization cascades to delete its user(s) too
        # (CustomUser.organization has on_delete=CASCADE).
        queryset.delete()
        self.stdout.write(self.style.SUCCESS(f"Deleted {count} expired trial account(s)."))
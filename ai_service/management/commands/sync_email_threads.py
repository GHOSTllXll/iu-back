# backend/ai_service/management/commands/sync_email_threads.py
"""
V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3): the
client's "5-minute polling task" — run as a Django management command
fired by CPANEL CRON (`*/5 * * * *`), NOT Celery beat. See email_sync.py's
own module docstring and config/celery.py's deploy-environment notes for
why: this host's cPanel Application Manager doesn't supervise a Celery
beat process at all, only a cron-based watchdog that keeps the Celery
WORKER (not beat) alive for the existing AI-pipeline tasks. A plain `python
manage.py sync_email_threads` invoked directly by cron needs no daemon,
no watchdog, and no new infrastructure — it runs, does one full pass over
every thread due for a poll, and exits; cron fires the next one 5 minutes
later regardless of how long (or whether) the previous run finished. The
client's own instruction said "Celery beat / cron" as an either/or — cron
is what this host can actually run unattended.

See backend’s own CRON_SETUP.md for the exact crontab line and a one-time
`crontab -e` walkthrough.

SCOPE PER RUN: every LinkedEmailThread with is_active=True, backfill_
status='complete' (a thread still 'pending' hasn't finished its one-time
backfill yet — see LinkedEmailThread's own docstring on why this command
skips rather than races it; a 'failed' backfill needs the user to notice
and relink rather than being silently retried forever), and a connection
whose status is still 'active' (a connection already known to be dead —
see EmailOAuthConnection.status — is never retried here; it was already
reported via Notification the first time it broke, by whichever call
site discovered that — see email_sync.mark_connection_needs_reauth).

FAILURE ISOLATION: one broken thread or connection NEVER stops the rest
of this run — every thread is wrapped in its own try/except, and a
connection that fails once (ReauthRequired) is skipped for every
remaining thread on that SAME connection within this run (no point
retrying a dead refresh token 40 more times in the next 40 threads),
but every OTHER connection's threads still get their turn.
"""
import logging

from django.core.management.base import BaseCommand
from django.utils import timezone


class Command(BaseCommand):
    help = (
        "Polls every active, backfilled LinkedEmailThread for new messages "
        "(the 5-minute broker/lender email sync). Intended to be fired by "
        "cron every 5 minutes — see backend’s own CRON_SETUP.md."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--thread-link-id',
            type=int,
            default=None,
            help=(
                "Sync exactly ONE LinkedEmailThread by id, ignoring its "
                "is_active/backfill_status (an ops/debugging escape hatch — "
                "the real 5-minute cron run never passes this)."
            ),
        )

    def handle(self, *args, **options):
        from ai_service.models import LinkedEmailThread
        from ai_service import email_sync

        logger = logging.getLogger('ai_service.sync_email_threads')
        started_at = timezone.now()

        thread_link_id = options.get('thread_link_id')
        if thread_link_id is not None:
            threads = LinkedEmailThread.objects.filter(pk=thread_link_id).select_related('connection')
        else:
            threads = (
                LinkedEmailThread.objects.filter(
                    is_active=True,
                    backfill_status='complete',
                    connection__status='active',
                )
                .select_related('connection')
                .order_by('connection_id', 'id')
            )

        total_threads = 0
        total_new_messages = 0
        failed_threads = 0
        skipped_connection_ids = set()

        for thread_link in threads:
            connection = thread_link.connection

            if connection.id in skipped_connection_ids:
                continue

            total_threads += 1
            try:
                new_count = email_sync.sync_linked_thread(thread_link)
                total_new_messages += new_count
                if new_count:
                    logger.info(
                        "sync_email_threads: thread_link=%s pulled %d new message(s).",
                        thread_link.id, new_count,
                    )
            except email_sync.ReauthRequired as e:
                failed_threads += 1
                skipped_connection_ids.add(connection.id)
                email_sync.mark_connection_needs_reauth(connection, str(e))
                logger.warning(
                    "sync_email_threads: connection=%s needs reconnect (%s) — "
                    "skipping its remaining threads for this run.",
                    connection.id, e,
                )
            except email_sync.EmailSyncError as e:
                failed_threads += 1
                logger.warning("sync_email_threads: thread_link=%s failed: %s", thread_link.id, e)
            except Exception:
                # Defense in depth - a bug in one thread's normalize/
                # classify path must never take down the other threads
                # (other orgs, other deals) still waiting their turn in
                # this same 5-minute run.
                failed_threads += 1
                logger.exception("sync_email_threads: unexpected error on thread_link=%s.", thread_link.id)

        elapsed = (timezone.now() - started_at).total_seconds()
        summary = (
            f"sync_email_threads: {total_threads} thread(s) checked, "
            f"{total_new_messages} new message(s), {failed_threads} failed/skipped, "
            f"{len(skipped_connection_ids)} connection(s) newly flagged for reconnect, "
            f"in {elapsed:.1f}s."
        )
        logger.info(summary)
        self.stdout.write(self.style.SUCCESS(summary))

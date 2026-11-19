# backend/ai_service/tasks.py
"""
Celery tasks wrapping the long-running (1-2+ minute) AI extraction pipelines
for both the CRE/Multifamily and institutional PPM underwriting flows.

WHY THIS EXISTS: both pipelines call out to an external AI API synchronously,
and on this VPS's limited Passenger worker pool that meant one slow AI call
blocked every other request - including simple page loads - for its entire
duration (see the TIMING BREAKDOWN log line in views.py, which has observed
total request times of 5+ minutes against a 300-second upstream timeout).
Moving the AI-bound work into a Celery task run by a separate worker process
takes it out of the request/response cycle entirely: the view enqueues and
returns almost immediately, and the frontend polls TaskStatusView (see
views.py) until the task completes.

FILE HANDLING: Django's UploadedFile objects are tied to the request and
cannot be serialized across the broker to a separate worker process. The
calling view (see views.py's _save_upload_to_temp) saves each upload to a
temp file on disk BEFORE enqueueing and passes the temp path + original
filename instead of the file object itself. These tasks re-wrap each temp
file in django.core.files.File (which supplies the .name/.size/.seek()
interface the existing parsing functions in parsers.py already expect) and
delete the temp file in a finally block once done, success or failure, so a
failed task never leaks a file on disk.
"""
import os
import logging
import contextlib

from celery import shared_task
from django.core.files import File
from django.contrib.auth import get_user_model

from .analysis_cache import store_analysis

logger = logging.getLogger(__name__)


def _cleanup_temp_files(*paths):
    for path in paths:
        try:
            if path:
                os.remove(path)
        except OSError:
            # Already gone, or never existed - not worth failing the task over.
            pass


def _load_user_and_org(uploaded_by_id):
    """
    Re-fetches the user (and, through it, their organization, via the same
    getattr(user, 'organization', None) pattern views.py uses elsewhere)
    inside the worker process - the original request.user object can't cross
    the broker, only its primary key can.
    """
    if not uploaded_by_id:
        return None, None
    User = get_user_model()
    user = User.objects.filter(pk=uploaded_by_id).first()
    if user is None:
        return None, None
    return user, getattr(user, 'organization', None)


@shared_task(bind=True, name='ai_service.run_cre_underwriting_task')
def run_cre_underwriting_task(self, om_path, om_name, t12_path, t12_name,
                                rent_roll_path, rent_roll_name, tier, uploaded_by_id,
                                project_id=None, asking_price=None):
    """
    Background counterpart to the CRE/Multifamily pipeline. Imports
    process_underwriting_files from views.py INSIDE the function body
    (not at module top) to avoid a circular import, since views.py imports
    this module at its own top level to enqueue these tasks.

    project_id (Feature 3): forwarded as-is from UnderwritePropertyView —
    see process_underwriting_files'/record_analysis_report's own docstrings
    for exactly how it's resolved. Defaults to None (keyword, with a default)
    rather than a required positional, so any task already enqueued before
    this parameter existed — e.g. still sitting in the broker across a
    deploy — keeps deserializing and running correctly.

    Sprint "Workflow Upgrades" Feature 1 (Missing Doc Check): any of
    om_path/t12_path/rent_roll_path (and its matching _name) may now be
    None — _save_upload_to_temp (views.py) returns (None, None) for a
    document the analyst didn't submit, rather than always writing all 3
    temp files. Opened conditionally below via contextlib.ExitStack, since
    a plain `with open(a) as x, open(b) as y:` can't skip one of its
    managers when that path is None.

    asking_price: forwarded as-is to process_underwriting_files — see its
    own docstring for exactly how it's used/persisted. None (the default)
    matches every task already enqueued before this parameter existed,
    same reasoning as project_id above.
    """
    from .views import process_underwriting_files

    uploaded_by, organization = _load_user_and_org(uploaded_by_id)

    try:
        with contextlib.ExitStack() as stack:
            om_file = None
            if om_path:
                om_file = File(stack.enter_context(open(om_path, 'rb')), name=om_name)
            t12_file = None
            if t12_path:
                t12_file = File(stack.enter_context(open(t12_path, 'rb')), name=t12_name)
            rent_roll_file = None
            if rent_roll_path:
                rent_roll_file = File(stack.enter_context(open(rent_roll_path, 'rb')), name=rent_roll_name)

            metrics, rent_roll_df, error, deal_scorecard = process_underwriting_files(
                om_file, t12_file, rent_roll_file, tier=tier,
                organization=organization, uploaded_by=uploaded_by,
                project_id=project_id, asking_price=asking_price,
            )

        if error:
            # error is the same DRF Response process_underwriting_files has
            # always returned on failure (quota/validation/rate-limit/server
            # error) - .data/.status_code carry everything TaskStatusView
            # needs to reproduce the exact same error shape the frontend's
            # existing catch blocks already know how to handle.
            return {'error': True, 'status_code': error.status_code, 'data': error.data}

        analysis_id = store_analysis(metrics, rent_roll_df)
        return {
            'error': False, 'metrics': metrics, 'tier': tier, 'analysis_id': analysis_id,
            # Sprint "Workflow Upgrades" Feature 1: the 10-second Kill-or-
            # Proceed Scorecard — TaskStatusView passes this dict straight
            # through to the frontend unchanged (see its own docstring).
            'deal_scorecard': deal_scorecard,
        }
    finally:
        _cleanup_temp_files(om_path, t12_path, rent_roll_path)


@shared_task(bind=True, name='ai_service.run_inbound_email_underwriting_task')
def run_inbound_email_underwriting_task(self, om_path, om_name, t12_path, t12_name,
                                          rent_roll_path, rent_roll_name, tier, uploaded_by_id,
                                          project_id, explicit_asking_price=None):
    """
    Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): thin
    wrapper around run_cre_underwriting_task above, for deals that
    arrived via the inbound email webhook (InboundEmailWebhookView in
    views.py) and therefore need TWO things no manual upload ever needs.
    Calls run_cre_underwriting_task DIRECTLY — a normal synchronous
    Python call, not .delay() — rather than duplicating its temp-file-
    handling/cleanup logic; this is the standard Celery pattern for
    invoking one task from inside another without a second broker
    round-trip (the @shared_task wrapper's bind=True still works exactly
    the same way on a direct call as it does on a normal .delay()'d one).

      1. asking_price pre-fill — there's no analyst at an upload screen to
         type a Target Purchase Price in by hand. If the inbound email's
         subject/body carried an explicit Target:/Price: $X tag
         (explicit_asking_price — resolved by the webhook view BEFORE
         enqueueing, by parsing the raw email), that value is forwarded
         exactly like a normal upload's asking_price and nothing below
         runs. Otherwise, once extraction succeeds, this falls back to
         the AI-extracted metrics['deal_valuation']['target_purchase_price']
         — the ONE place in this codebase that ever does that:
         process_underwriting_files/record_analysis_report deliberately
         NEVER do this for a manual upload (see
         PropertyProject.asking_price's own docstring in models.py) since
         a human is always available there to enter one instead. Never
         overwrites an asking_price that's already non-null on the
         project, so a later re-analysis of the same email-originated
         deal can't clobber a value an analyst has since edited by hand
         in the drawer.
      2. property_name upgrade — the webhook view creates this
         PropertyProject up front using the email's SUBJECT LINE as a
         placeholder title (so the deal shows up in the pipeline
         immediately, before extraction even starts), then passes its id
         in as project_id here. record_analysis_report's existing-project
         branch never touches property_name (see its own docstring) — by
         design, for the ordinary "re-upload & re-analyze" case, where
         overwriting a deal's established name on every re-run would be
         wrong. This is the one deliberate exception: once extraction
         succeeds, the placeholder subject is replaced with the AI's own
         extracted property name, exactly once, here.

    On an extraction FAILURE (result['error'] is True), neither of the
    above runs — the placeholder subject-line title and any Target:/
    Price: tag asking_price (already applied via the forwarded
    explicit_asking_price argument, if given) simply stand as-is; there's
    nothing further to resolve without real metrics.
    """
    result = run_cre_underwriting_task(
        om_path, om_name, t12_path, t12_name, rent_roll_path, rent_roll_name,
        tier, uploaded_by_id, project_id=project_id, asking_price=explicit_asking_price,
    )

    if result.get('error'):
        logger.warning(
            "Inbound email intake: extraction failed for project %s — "
            "leaving its placeholder title/asking_price exactly as the "
            "webhook view set them.", project_id,
        )
        return result

    try:
        from .models import PropertyProject

        project = PropertyProject.objects.filter(pk=project_id).first()
        if project is not None:
            update_fields = []
            metrics = result.get('metrics') or {}

            extracted_name = (metrics.get('property_metadata') or {}).get('property_name') or ''
            if extracted_name:
                project.property_name = extracted_name
                update_fields.append('property_name')

            if explicit_asking_price is None and project.asking_price is None:
                om_price = (metrics.get('deal_valuation') or {}).get('target_purchase_price')
                if om_price:
                    project.asking_price = om_price
                    update_fields.append('asking_price')

            if update_fields:
                project.save(update_fields=update_fields)
    except Exception:
        # Best-effort polish on top of an ALREADY-successful extraction —
        # never let a failure here make a successful run look failed.
        logger.exception(
            "Inbound email intake: post-processing (title/asking_price "
            "pre-fill) failed for project %s — extraction itself still "
            "succeeded and its metrics are unaffected.", project_id,
        )

    return result


@shared_task(bind=True, name='ai_service.run_ppm_underwriting_task')
def run_ppm_underwriting_task(self, ppm_path, ppm_name, manual_pages, tier, uploaded_by_id):
    """Background counterpart to the institutional PPM pipeline."""
    from .views import process_ppm_file

    uploaded_by, organization = _load_user_and_org(uploaded_by_id)

    try:
        with open(ppm_path, 'rb') as f_ppm:
            ppm_file = File(f_ppm, name=ppm_name)
            metrics, error = process_ppm_file(
                ppm_file, manual_pages=manual_pages, tier=tier,
                organization=organization, uploaded_by=uploaded_by,
            )

        if error:
            return {'error': True, 'status_code': error.status_code, 'data': error.data}

        return {'error': False, 'metrics': metrics, 'tier': tier}
    finally:
        _cleanup_temp_files(ppm_path)


@shared_task(bind=True, name='ai_service.run_email_thread_backfill_task', max_retries=0)
def run_email_thread_backfill_task(self, thread_link_id):
    """
    V2 Module 1 (Phase V2-3): the ONE-TIME full-history pull a thread
    gets right after an analyst tags it to a deal (ProjectEmailThreadsView.
    post() in views.py enqueues this instead of blocking the request on
    however many Gmail/Graph round trips a long thread's backfill takes).
    Unlike every other task in this file, this one is NOT an AI-bound
    pipeline — it's here, rather than a plain synchronous call, purely
    to keep the HTTP response fast; the actual backfill logic (and all
    of its own error handling — a failed backfill is recorded on
    LinkedEmailThread.backfill_status/backfill_error, never raised back
    to Celery to retry) lives entirely in email_sync.backfill_linked_thread.
    The 5-minute poll (sync_email_threads management command) is a
    completely separate, cron-driven entry point into the SAME
    underlying sync_linked_thread() — this task exists only for the
    one-time backfill's latency, not as part of the recurring sync.
    """
    from .models import LinkedEmailThread
    from . import email_sync

    try:
        thread_link = LinkedEmailThread.objects.select_related('connection', 'connection__user').get(pk=thread_link_id)
    except LinkedEmailThread.DoesNotExist:
        logger.warning("run_email_thread_backfill_task: LinkedEmailThread %s no longer exists.", thread_link_id)
        return

    email_sync.backfill_linked_thread(thread_link)

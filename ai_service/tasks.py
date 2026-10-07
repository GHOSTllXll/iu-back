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
                                rent_roll_path, rent_roll_name, tier, uploaded_by_id):
    """
    Background counterpart to the CRE/Multifamily pipeline. Imports
    process_underwriting_files from views.py INSIDE the function body
    (not at module top) to avoid a circular import, since views.py imports
    this module at its own top level to enqueue these tasks.
    """
    from .views import process_underwriting_files

    uploaded_by, organization = _load_user_and_org(uploaded_by_id)

    try:
        with open(om_path, 'rb') as f_om, open(t12_path, 'rb') as f_t12, open(rent_roll_path, 'rb') as f_rr:
            om_file = File(f_om, name=om_name)
            t12_file = File(f_t12, name=t12_name)
            rent_roll_file = File(f_rr, name=rent_roll_name)

            metrics, rent_roll_df, error = process_underwriting_files(
                om_file, t12_file, rent_roll_file, tier=tier,
                organization=organization, uploaded_by=uploaded_by,
            )

        if error:
            # error is the same DRF Response process_underwriting_files has
            # always returned on failure (quota/validation/rate-limit/server
            # error) - .data/.status_code carry everything TaskStatusView
            # needs to reproduce the exact same error shape the frontend's
            # existing catch blocks already know how to handle.
            return {'error': True, 'status_code': error.status_code, 'data': error.data}

        analysis_id = store_analysis(metrics, rent_roll_df)
        return {'error': False, 'metrics': metrics, 'tier': tier, 'analysis_id': analysis_id}
    finally:
        _cleanup_temp_files(om_path, t12_path, rent_roll_path)


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

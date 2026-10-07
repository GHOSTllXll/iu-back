"""
Celery application for the background-task queue. Added so the long-running
(1-2+ minute) AI extraction calls in ai_service's CRE and PPM pipelines no
longer run synchronously inside a Passenger request - see ai_service/tasks.py
for why, and ai_service/views.py's UnderwritePropertyView / PPMUnderwriteView
/ TaskStatusView for how the request/response cycle now hands this off.

Run the worker on the server with:
    celery -A config worker --loglevel=info

cPanel's Application Manager does not supervise this process - see the
deploy notes for the cron-based watchdog that keeps it alive.
"""
import os

from celery import Celery

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

app = Celery('config')
# Reads every CELERY_* setting from Django's settings.py (see that file for
# the actual values) - namespace='CELERY' means we write CELERY_BROKER_URL
# there instead of bare BROKER_URL, consistent with every other setting.
app.config_from_object('django.conf:settings', namespace='CELERY')
app.autodiscover_tasks()

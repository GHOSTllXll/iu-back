import pymysql
pymysql.install_as_MySQLdb()

# Load the Celery app on Django startup so @shared_task-decorated tasks
# in ai_service/tasks.py register correctly. See config/celery.py.
from .celery import app as celery_app

__all__ = ('celery_app',)

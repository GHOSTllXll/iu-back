#!/bin/bash
# ensure_celery_worker.sh - "poor man's supervisor" for the Celery worker.
#
# cPanel's Application Manager only supervises the Passenger-fronted Django
# app (restarting it on crash, etc.) - it has no concept of a standalone
# background process like a Celery worker, so nothing restarts the worker
# if it dies or the box reboots. This script is meant to be run every few
# minutes by a cPanel Cron Job: if the worker is already running, it exits
# immediately and does nothing; if not, it starts one. Mirrors the existing
# cleanup_expired_trials cron job's "source venv, cd, run" pattern already
# in use on this account.
#
# Concurrency is intentionally conservative (2 worker processes) given the
# server's 4 cores are shared with Apache/Passenger and MySQL - raise it if
# profiling later shows headroom.

source /home/blackbri/virtualenv/backend/3.13/bin/activate
cd /home/blackbri/backend

PIDFILE=/home/blackbri/backend/celery_worker.pid
LOGFILE=/home/blackbri/backend/celery_worker.log

if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
    exit 0
fi

# Stale pidfile (process gone but file left behind) would make kill -0 fail
# above and fall through here, which is correct - celery --detach overwrites
# it with the new process's PID.
celery -A config worker --loglevel=info --concurrency=2 \
    --detach --pidfile="$PIDFILE" --logfile="$LOGFILE"

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
# above and fall through here, which is correct - celery overwrites it with
# the new process's PID below.
#
# Bug fix (Nov 2026): this used to pass --detach and let Celery do its own
# double-fork daemonization. On this host/Python 3.13/Celery 5.4 combo that
# repeatedly (every single restart, 4/4 observed) left behind a ZOMBIE
# --detach parent process pegged at ~100% CPU indefinitely, while a SEPARATE
# worker process came up fine moments later under its own PID — meaning the
# --detach parent never actually handed off and exited the way it's
# supposed to. Harmless only because someone was watching and manually
# killing the stuck one each time; left alone (e.g. the cron watchdog
# firing unattended), these would accumulate forever on a shared box.
#
# Fix: stop asking Celery to daemonize itself at all, and do it ourselves
# with plain, well-understood shell tools instead of Celery's apparently
# fragile fork-and-poll-for-the-pidfile logic. setsid starts the worker in
# its own new session, fully detached from this script's process group —
# so it survives this script (and the cron job that launched it) exiting,
# without needing a second fork. nohup additionally ignores SIGHUP as a
# belt-and-suspenders measure, </dev/null detaches stdin, and redirecting
# stdout/stderr to LOGFILE catches anything printed before Celery's own
# --logfile handling takes over (e.g. warnings.warn() calls during import,
# which bypass structured logging). Backgrounding with & and letting the
# script then simply exit is what actually daemonizes it — once orphaned,
# the worker gets reparented to init like any background process. Celery
# still writes its own --pidfile exactly as before; since there's no extra
# fork anymore, that PID is the same one backgrounding would have given us
# via $!, so nothing else here needs to change.
setsid nohup celery -A config worker --loglevel=info --concurrency=2 \
    --pidfile="$PIDFILE" --logfile="$LOGFILE" \
    < /dev/null >> "$LOGFILE" 2>&1 &

# Email Thread Sync — cPanel Cron Setup (V2 Module 1, Phase V2-3)

This feature's "every 5 minutes" poll is a plain Django management
command (`sync_email_threads`), fired by cPanel's own Cron Jobs UI —
**not** a Celery beat daemon. See `ai_service/email_sync.py`'s module
docstring and `config/celery.py`'s own deploy-environment notes for why:
this host's Application Manager doesn't supervise a `celery beat`
process at all, only a cron-based watchdog that keeps the existing
Celery **worker** (used by the AI underwriting pipelines) alive. A command
invoked directly by cron needs no daemon and no watchdog — it runs, does
one pass, and exits.

## 1. Add the cron job

In cPanel → **Cron Jobs**:

- **Common Settings:** "Once Per 5 Minutes" if offered, otherwise set the
  Minute field to `*/5` and leave Hour/Day/Month/Weekday as `*`.
- **Command** (adjust the two paths for this account):

```
cd /home/<cpanel_user>/path/to/backend && /home/<cpanel_user>/path/to/venv/bin/python manage.py sync_email_threads >> /home/<cpanel_user>/logs/sync_email_threads.log 2>&1
```

Use the **same** virtualenv `python` binary and backend directory the
existing Celery-worker watchdog cron job already uses — don't invoke
the system `python3` directly, or a missing package (this feature adds
no new ones beyond what's already in `requirements.txt`: `cryptography`
and `requests`, both already present) would fail silently in the log.

Create the log directory first if it doesn't exist (`mkdir -p
/home/<cpanel_user>/logs`) — cPanel's cron does not create it for you,
and a missing directory means the redirect above silently drops output.

## 2. Set the required environment variables

Add these to whatever this app already uses to set environment
variables for the Django process (cPanel's "Setup Python App" env-vars
panel, or a `.env` file loaded the same way the existing
`RENTCAST_API_KEY` / `POSTMARK_INBOUND_WEBHOOK_*` vars are). All of them
default to blank/disabled — this feature stays off until they're set,
the same "blank means not configured yet" pattern as every other V2
integration (see `config/settings.py`'s own comments above each one):

| Variable | Value |
|---|---|
| `BACKEND_BASE_URL` | `https://inboundunderwriting.com` (this app's own public URL — used ONLY to build the OAuth `redirect_uri`; see below) |
| `GOOGLE_OAUTH_CLIENT_ID` | from Google Cloud Console (step 3) |
| `GOOGLE_OAUTH_CLIENT_SECRET` | from Google Cloud Console (step 3) |
| `MS_OAUTH_CLIENT_ID` | from the Microsoft Entra App Registration (step 4) |
| `MS_OAUTH_CLIENT_SECRET` | from the Microsoft Entra App Registration (step 4) |
| `MS_OAUTH_TENANT` | `common` (supports both Microsoft 365 work accounts and personal Outlook.com/Hotmail through the same app — leave as `common` unless this is deliberately locked down to one org's tenant later) |
| `EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY` | a Fernet key — generate ONE with the command below and never regenerate it once real users have connected a mailbox (regenerating it makes every already-stored token undecryptable, moving every connection to `status='error'`) |

Generate the encryption key once:

```
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

`FRONTEND_BASE_URL` is a pre-existing setting (used elsewhere already) —
the OAuth callback 302s the browser there (`/dashboard/settings`) once a
connection succeeds or fails; no new value needed unless it's wrong already.

## 3. Register the app with Google (Gmail)

In Google Cloud Console → APIs & Services → Credentials → create an
**OAuth 2.0 Client ID** of type **Web application**:

- **Authorized redirect URI:** `https://inboundunderwriting.com/api/ai/email-sync/oauth/callback/google/`
  (must match `BACKEND_BASE_URL` above character-for-character — see
  `email_oauth._redirect_uri()`)
- Enable the **Gmail API** for the project.
- The scope this app requests, `gmail.readonly`, is a Google-classified
  **RESTRICTED** scope (not just "sensitive") — before this goes live for
  real users, this project needs to pass Google's own security
  assessment for a production app storing that data server-side. That's
  a business/compliance step on Google's side, not something any code
  change here can satisfy. Until it's done, Google will cap this to a
  small number of test users added under "Test users" in the OAuth
  consent screen configuration.

## 4. Register the app with Microsoft (Microsoft 365 + Outlook.com)

In the Azure Portal → Microsoft Entra ID → App registrations → **New
registration**:

- **Supported account types:** "Accounts in any organizational directory
  and personal Microsoft accounts" (this is what lets ONE app serve both
  Microsoft 365 work accounts and personal Outlook.com/Hotmail accounts
  through the same `common` tenant — see `MS_OAUTH_TENANT` above).
- **Redirect URI** (platform: **Web**): `https://inboundunderwriting.com/api/ai/email-sync/oauth/callback/microsoft/`
- Under **Certificates & secrets**, create a **Client secret** — that
  value is `MS_OAUTH_CLIENT_SECRET`; the **Application (client) ID** on
  the Overview page is `MS_OAUTH_CLIENT_ID`.
- No API permissions need to be pre-granted in the portal — `Mail.Read`
  and `offline_access` are requested dynamically on the consent screen
  (see `email_oauth.MS_SCOPES`) and Microsoft's own delegated-permission
  model handles personal-account consent automatically.

## 5. Verify it's actually running

After the first cron fire (wait up to 5 minutes, or trigger it by hand
over SSH/cPanel Terminal with the same command the cron job uses, minus
the redirect), check the log:

```
tail -f /home/<cpanel_user>/logs/sync_email_threads.log
```

A healthy run looks like:

```
sync_email_threads: 3 thread(s) checked, 1 new message(s), 0 failed/skipped, 0 connection(s) newly flagged for reconnect, in 0.8s.
```

Nothing to check until at least one analyst has connected a mailbox
(Settings → Connected Email Accounts) and tagged at least one thread to
a deal (a deal's drawer → Linked Email Threads → "+ Link Email Thread").
Before either of those exist, every run will simply report "0 thread(s)
checked" — that's expected, not a failure.

To sync one specific thread by hand (useful while testing, bypassing
the `is_active`/`backfill_status` filters the real poll applies):

```
python manage.py sync_email_threads --thread-link-id <id>
```

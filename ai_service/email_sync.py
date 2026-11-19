# backend/ai_service/email_sync.py
"""
V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3): the
actual mailbox-reading logic — searching a connected mailbox for
candidate threads to link, pulling every message in a newly-linked
thread (full backfill, per the client's explicit decision), and the
incremental re-pull the 5-minute poll (sync_email_threads management
command) runs against every already-linked, active thread.

Provider differences are isolated to the handful of functions prefixed
_google_/_microsoft_ below; everything else (token refresh, attachment
classify-then-discard, body-text capping, SyncedEmailMessage upsert) is
provider-agnostic and runs identically either way.

ZERO-RETENTION BOUNDARY (read before touching this file): an attachment
classified as OM/T12/Rent Roll is classified IN MEMORY from bytes this
module fetched over HTTPS, and those bytes are then dropped — never
written to disk, never passed to DealVaultFile, never handed to the AI
extraction pipeline. This is intentionally narrower than the Postmark
inbound-intake path (_classify_inbound_attachments in views.py), which
extracts and keeps structured data from those same three document
types because THAT path is creating a brand-new deal from them. This
phase's job is syncing correspondence into an EXISTING deal's Activity
Feed, not re-running underwriting every time a broker emails an
updated T12 — the client's own instruction was "classify-then-discard,"
not "classify-then-analyze." An attachment that does NOT classify as
one of those three types is vaulted normally (DealVaultFile,
file_type='supporting_document') — a signed LOI or inspection report
arriving by email is exactly the kind of long-lived supporting document
that file type already exists for.
"""
import logging
import re
from datetime import datetime, timedelta, timezone as dt_timezone

from django.core.files.base import ContentFile
from django.utils import timezone

from . import email_oauth
from .document_classifier import (
    CLASSIFIABLE_EXTENSIONS,
    DOCUMENT_TYPE_OFFERING_MEMORANDUM,
    DOCUMENT_TYPE_TRAILING_12,
    DOCUMENT_TYPE_RENT_ROLL,
    DocumentClassificationError,
    classify_document,
)

logger = logging.getLogger(__name__)

MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024
MAX_ATTACHMENTS_PER_MESSAGE = 10
MAX_SYNCED_BODY_CHARS = 50_000

# Same allowlist Vault's own manual "Supporting Document" upload already
# enforces (ALLOWED_VAULT_SUPPORTING_DOC_EXTENSIONS in views.py) —
# duplicated here as a plain constant rather than imported, to avoid
# views.py <-> this module import ordering, but kept in sync by eye;
# never widened independently of that one.
ALLOWED_VAULT_EXTENSIONS = {
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.csv', '.png', '.jpg', '.jpeg',
}

_TYPE_TO_SLOT = {
    DOCUMENT_TYPE_OFFERING_MEMORANDUM: 'om',
    DOCUMENT_TYPE_TRAILING_12: 't12',
    DOCUMENT_TYPE_RENT_ROLL: 'rent_roll',
}


class EmailSyncError(Exception):
    """A per-connection or per-thread failure the caller should log and skip, never crash the whole poll over."""
    pass


class ReauthRequired(EmailSyncError):
    """The connection's refresh token no longer works — caller should mark status='error' and notify the user."""
    pass


def mark_connection_needs_reauth(connection, error_message: str) -> None:
    """
    Single shared place every ReauthRequired call site (thread search,
    backfill, the 5-minute sync_email_threads command) reports into —
    flips the connection to 'error' (Settings page and the 5-minute poll
    both treat that as "needs reconnect," never a crash) and fires one
    Notification(type='email_sync_error') so the user actually notices
    outside the Settings page. Deliberately only notifies on the
    active -> error TRANSITION, not every call: without that guard, a
    connection stuck in 'error' would re-notify the user every single
    5-minute poll tick forever over the same already-known problem.
    """
    from .models import Notification

    was_already_error = connection.status == 'error'
    connection.status = 'error'
    connection.last_error = str(error_message)[:2000]
    connection.save(update_fields=['status', 'last_error', 'updated_at'])

    if not was_already_error:
        Notification.objects.create(
            recipient=connection.user,
            notification_type='email_sync_error',
            project=None,
            message=(
                f"Your {connection.get_provider_display()} connection ({connection.email_address}) "
                f"needs to be reconnected before its linked deal threads can keep syncing."
            ),
        )


# ---------------------------------------------------------------------------
# Token refresh
# ---------------------------------------------------------------------------

def ensure_fresh_access_token(connection) -> str:
    """
    Returns a valid access token for `connection`, refreshing it first
    if it's expired (or within 2 minutes of expiring, to avoid a race
    against a slow downstream call). Persists the refreshed token back
    onto the connection row. Raises ReauthRequired — never lets a failed
    refresh surface as a generic exception — if the refresh itself
    fails, since that always means the SAME thing regardless of
    provider: the refresh_token no longer works (revoked, password
    changed, admin pulled app access, or this app's own encryption key
    rotated — see token_crypto.py) and this connection needs the user to
    reconnect.
    """
    needs_refresh = (
        connection.access_token_expires_at is None
        or connection.access_token_expires_at <= timezone.now() + timedelta(minutes=2)
    )
    if not needs_refresh:
        token = connection.get_access_token()
        if token:
            return token

    refresh_token = connection.get_refresh_token()
    if not refresh_token:
        raise ReauthRequired(f"No refresh token stored for connection {connection.id}.")

    try:
        if connection.provider == email_oauth.GOOGLE:
            result = email_oauth.refresh_google_token(refresh_token)
        else:
            result = email_oauth.refresh_microsoft_token(refresh_token)
    except email_oauth.EmailOAuthError as e:
        raise ReauthRequired(str(e))

    connection.set_access_token(result['access_token'], result['expires_at'])
    if result.get('refresh_token'):
        connection.set_refresh_token(result['refresh_token'])
    connection.save(update_fields=['encrypted_access_token', 'encrypted_refresh_token', 'access_token_expires_at', 'updated_at'])
    return result['access_token']


# ---------------------------------------------------------------------------
# Candidate thread search (for the "Link Email Thread" picker)
# ---------------------------------------------------------------------------

def search_candidate_threads(connection, query: str) -> list[dict]:
    """
    Returns [{provider_thread_id, subject, snippet, participant, last_message_at}],
    newest first — what EmailThreadSearchView in views.py hands the
    frontend's thread picker. Gmail has a real thread-list endpoint;
    Microsoft Graph doesn't, so its messages are deduped down to one row
    per conversationId (keeping the most recent message) client-side
    here.
    """
    access_token = ensure_fresh_access_token(connection)

    if connection.provider == email_oauth.GOOGLE:
        threads = email_oauth.search_google_threads(access_token, query)
        results = []
        for t in threads:
            try:
                full = email_oauth.get_google_thread(access_token, t['id'])
            except email_oauth.EmailOAuthError:
                continue
            messages = full.get('messages') or []
            if not messages:
                continue
            last = messages[-1]
            headers = {h['name'].lower(): h['value'] for h in (last.get('payload') or {}).get('headers', [])}
            results.append({
                'provider_thread_id': full['id'],
                'subject': headers.get('subject', '(no subject)'),
                'snippet': last.get('snippet', ''),
                'participant': headers.get('from', ''),
                'last_message_at': _gmail_internal_date(last),
            })
        return results

    messages = email_oauth.search_microsoft_threads(access_token, query)
    by_conversation = {}
    for m in messages:
        cid = m.get('conversationId')
        if not cid:
            continue
        existing = by_conversation.get(cid)
        if existing is None or m.get('receivedDateTime', '') > existing.get('receivedDateTime', ''):
            by_conversation[cid] = m
    return [
        {
            'provider_thread_id': cid,
            'subject': m.get('subject') or '(no subject)',
            'snippet': m.get('bodyPreview', ''),
            'participant': (m.get('from') or {}).get('emailAddress', {}).get('address', ''),
            'last_message_at': m.get('receivedDateTime'),
        }
        for cid, m in by_conversation.items()
    ]


# ---------------------------------------------------------------------------
# Pulling a thread's messages (backfill AND incremental re-pull both call
# this — a full re-fetch each time is simpler and more robust than a
# cursor, and SyncedEmailMessage's own (thread_link, provider_message_id)
# uniqueness makes a re-fetch idempotent for free).
# ---------------------------------------------------------------------------

def sync_linked_thread(thread_link) -> int:
    """
    Pulls every message currently in this thread, creates a
    SyncedEmailMessage for any one not already synced, and returns how
    many NEW messages were created. Updates thread_link.last_synced_at
    on success. Raises EmailSyncError/ReauthRequired on failure — the
    caller (management command) is responsible for catching that per
    thread so one broken thread/connection never stops the rest of the
    5-minute run.
    """
    from .models import SyncedEmailMessage

    connection = thread_link.connection
    access_token = ensure_fresh_access_token(connection)

    if connection.provider == email_oauth.GOOGLE:
        try:
            full = email_oauth.get_google_thread(access_token, thread_link.provider_thread_id)
        except email_oauth.EmailOAuthError as e:
            raise EmailSyncError(str(e))
        raw_messages = full.get('messages') or []
        normalize = _normalize_google_message
    else:
        try:
            raw_messages = email_oauth.get_microsoft_conversation_messages(access_token, thread_link.provider_thread_id)
        except email_oauth.EmailOAuthError as e:
            raise EmailSyncError(str(e))
        normalize = _normalize_microsoft_message

    new_count = 0
    for raw in raw_messages:
        normalized = normalize(raw, connection.email_address)
        if normalized is None:
            continue
        msg_obj, created = SyncedEmailMessage.objects.get_or_create(
            thread_link=thread_link,
            provider_message_id=normalized['provider_message_id'],
            defaults={
                'sent_at': normalized['sent_at'],
                'sender_email': normalized['sender_email'],
                'sender_name': normalized['sender_name'],
                'direction': normalized['direction'],
                'body_text': normalized['body_text'],
                'attachment_summary': [],
                'vault_file_ids': [],
            },
        )
        if created:
            new_count += 1
            if normalized['has_attachments']:
                summary, vault_ids = _process_attachments(
                    connection, access_token, thread_link, raw, normalized,
                )
                SyncedEmailMessage.objects.filter(pk=msg_obj.pk).update(
                    attachment_summary=summary, vault_file_ids=vault_ids,
                )

    thread_link.last_synced_at = timezone.now()
    thread_link.save(update_fields=['last_synced_at'])
    return new_count


def backfill_linked_thread(thread_link) -> None:
    """
    The one-time full-history pull a newly-linked thread gets (see
    LinkedEmailThread.backfill_status's own docstring). Identical
    mechanics to an incremental sync_linked_thread call — "backfill" is
    just this function's name for the first one, there's no separate
    code path to keep in sync.
    """
    try:
        sync_linked_thread(thread_link)
        thread_link.backfill_status = 'complete'
        thread_link.backfill_error = ''
        thread_link.save(update_fields=['backfill_status', 'backfill_error'])
    except ReauthRequired as e:
        mark_connection_needs_reauth(thread_link.connection, str(e))
        thread_link.backfill_status = 'failed'
        thread_link.backfill_error = f"Reconnect needed: {e}"[:2000]
        thread_link.save(update_fields=['backfill_status', 'backfill_error'])
    except Exception as e:
        logger.exception("Backfill failed for linked thread %s", thread_link.id)
        thread_link.backfill_status = 'failed'
        thread_link.backfill_error = str(e)[:2000]
        thread_link.save(update_fields=['backfill_status', 'backfill_error'])


# ---------------------------------------------------------------------------
# Attachment classify-then-discard / vault
# ---------------------------------------------------------------------------

def _classify_bytes(filename: str, raw_bytes: bytes) -> str:
    """Returns 'om' | 't12' | 'rent_roll' | 'other' | 'unclassified'. Mirrors _classify_inbound_attachments' own per-file logic in views.py."""
    import os
    ext = os.path.splitext(filename)[1].lower()
    if ext not in CLASSIFIABLE_EXTENSIONS:
        return 'other' if ext in ALLOWED_VAULT_EXTENSIONS else 'unclassified'
    try:
        wrapped = ContentFile(raw_bytes, name=filename)
        result = classify_document(wrapped)
    except DocumentClassificationError:
        return 'unclassified'
    return _TYPE_TO_SLOT.get(result.get('document_type'), 'other')


def _process_attachments(connection, access_token: str, thread_link, raw_message: dict, normalized: dict):
    """
    Returns (attachment_summary, vault_file_ids) — see
    SyncedEmailMessage's own docstring for the exact shape. Fetches
    each attachment's bytes, classifies in memory, and either discards
    (om/t12/rent_roll — zero retention) or vaults (everything else) it.
    Caps at MAX_ATTACHMENTS_PER_MESSAGE / MAX_ATTACHMENT_BYTES, same
    defensive bound as the Postmark inbound path's MAX_INBOUND_
    ATTACHMENTS/_BYTES, against a hostile or oversized message.
    """
    from .models import DealVaultFile

    summary = []
    vault_ids = []

    if connection.provider == email_oauth.GOOGLE:
        candidates = _iter_google_attachments(raw_message, access_token)
    else:
        candidates = _iter_microsoft_attachments(access_token, raw_message['id'])

    for filename, size_hint, fetch_bytes in candidates[:MAX_ATTACHMENTS_PER_MESSAGE]:
        if size_hint and size_hint > MAX_ATTACHMENT_BYTES:
            summary.append({'filename': filename, 'classified_as': None, 'action': 'skipped_too_large'})
            continue
        try:
            raw_bytes = fetch_bytes()
        except Exception:
            summary.append({'filename': filename, 'classified_as': None, 'action': 'skipped_error'})
            continue
        if not raw_bytes or len(raw_bytes) > MAX_ATTACHMENT_BYTES:
            summary.append({'filename': filename, 'classified_as': None, 'action': 'skipped_too_large'})
            continue

        classified_as = _classify_bytes(filename, raw_bytes)

        if classified_as in ('om', 't12', 'rent_roll'):
            # Zero-retention: bytes go out of scope here, never written anywhere.
            summary.append({'filename': filename, 'classified_as': classified_as, 'action': 'discarded'})
            continue

        if classified_as == 'unclassified':
            summary.append({'filename': filename, 'classified_as': 'unclassified', 'action': 'skipped_unrecognized_type'})
            continue

        vault_file = DealVaultFile.objects.create(
            project=thread_link.project,
            file=ContentFile(raw_bytes, name=filename),
            file_type='supporting_document',
            original_filename=filename,
            version_label=f"Email attachment — {normalized['sender_name'] or normalized['sender_email'] or 'unknown sender'}",
            uploaded_by=connection.user,
        )
        vault_ids.append(vault_file.id)
        summary.append({'filename': filename, 'classified_as': 'other', 'action': 'stored_in_vault'})

    return summary, vault_ids


# ---------------------------------------------------------------------------
# Provider-specific message/attachment extraction
# ---------------------------------------------------------------------------

def _gmail_internal_date(message: dict) -> str:
    """Gmail's internalDate is epoch milliseconds as a string — returns an ISO 8601 string."""
    try:
        ms = int(message.get('internalDate', '0'))
        return datetime.fromtimestamp(ms / 1000, tz=dt_timezone.utc).isoformat()
    except (TypeError, ValueError):
        return timezone.now().isoformat()


def _gmail_headers(message: dict) -> dict:
    return {h['name'].lower(): h['value'] for h in (message.get('payload') or {}).get('headers', [])}


_EMAIL_ADDR_RE = re.compile(r'<([^>]+)>')


def _parse_address(header_value: str) -> tuple[str, str]:
    """'Jane Doe <jane@cbre.com>' -> ('jane@cbre.com', 'Jane Doe'); a bare address has no name."""
    if not header_value:
        return '', ''
    match = _EMAIL_ADDR_RE.search(header_value)
    if match:
        name = header_value[:match.start()].strip().strip('"')
        return match.group(1).strip().lower(), name
    return header_value.strip().lower(), ''


def _strip_html(html: str) -> str:
    return re.sub(r'<[^>]+>', ' ', html or '')


def _normalize_google_message(message: dict, connection_email: str) -> dict | None:
    headers = _gmail_headers(message)
    sender_email, sender_name = _parse_address(headers.get('from', ''))
    if not sender_email:
        return None

    body_text = _gmail_body_text(message.get('payload') or {})
    has_attachments = bool(_google_attachment_parts(message))

    return {
        'provider_message_id': message['id'],
        'sent_at': _gmail_internal_date(message),
        'sender_email': sender_email,
        'sender_name': sender_name,
        'direction': 'outbound' if sender_email == (connection_email or '').lower() else 'inbound',
        'body_text': body_text[:MAX_SYNCED_BODY_CHARS],
        'has_attachments': has_attachments,
    }


def _gmail_body_text(payload: dict) -> str:
    """Walks payload.parts depth-first for a text/plain part; falls back to tag-stripped text/html; '' if neither exists."""
    import base64

    def walk(part, mime_type):
        if part.get('mimeType') == mime_type and (part.get('body') or {}).get('data'):
            raw = part['body']['data']
            try:
                return base64.urlsafe_b64decode(raw + '=' * (-len(raw) % 4)).decode('utf-8', errors='replace')
            except Exception:
                return ''
        for sub in part.get('parts') or []:
            found = walk(sub, mime_type)
            if found:
                return found
        return ''

    text = walk(payload, 'text/plain')
    if text:
        return text
    html = walk(payload, 'text/html')
    return _strip_html(html).strip()


def _google_attachment_parts(message: dict) -> list[tuple[str, str, int]]:
    """
    Returns [(filename, attachment_id, size_hint)] for every real
    attachment part in this message's payload — a part counts as a real
    attachment only when it has BOTH a filename AND a body.attachmentId
    (an inline body part, e.g. the text/plain/html part itself, has
    neither). No network call — just walks the already-fetched message
    JSON, so it's cheap enough to use for a has_attachments check too.
    """
    results = []

    def walk(part):
        filename = part.get('filename')
        body = part.get('body') or {}
        attachment_id = body.get('attachmentId')
        if filename and attachment_id:
            results.append((filename, attachment_id, body.get('size', 0)))
        for sub in part.get('parts') or []:
            walk(sub)

    walk(message.get('payload') or {})
    return results


def _iter_google_attachments(message: dict, access_token: str) -> list:
    """Returns [(filename, size_hint, fetch_bytes_callable)] — the callable makes the actual Gmail API call to fetch that one attachment's bytes, bound to the CALLER's already-fresh access_token (never a stashed/global one, so a token refreshed mid-run is always the one actually used)."""
    message_id = message['id']
    return [
        (
            filename,
            size_hint,
            lambda aid=attachment_id, mid=message_id, token=access_token: email_oauth.get_google_attachment(token, mid, aid),
        )
        for filename, attachment_id, size_hint in _google_attachment_parts(message)
    ]


def _normalize_microsoft_message(message: dict, connection_email: str) -> dict | None:
    sender = (message.get('from') or {}).get('emailAddress', {}) or {}
    sender_email = (sender.get('address') or '').strip().lower()
    if not sender_email:
        return None

    body = message.get('body') or {}
    if body.get('contentType') == 'text':
        body_text = body.get('content', '')
    else:
        body_text = _strip_html(body.get('content', ''))

    return {
        'provider_message_id': message['id'],
        'sent_at': message.get('receivedDateTime') or timezone.now().isoformat(),
        'sender_email': sender_email,
        'sender_name': sender.get('name', ''),
        'direction': 'outbound' if sender_email == (connection_email or '').lower() else 'inbound',
        'body_text': body_text.strip()[:MAX_SYNCED_BODY_CHARS],
        'has_attachments': bool(message.get('hasAttachments')),
    }


def _iter_microsoft_attachments(access_token: str, message_id: str) -> list:
    """Returns [(filename, size_hint, fetch_bytes_callable)] — Graph's list-attachments response already includes contentBytes for file attachments, so the 'fetch' is just a base64 decode of data already in hand."""
    import base64

    try:
        attachments = email_oauth.list_microsoft_attachments(access_token, message_id)
    except email_oauth.EmailOAuthError:
        return []

    results = []
    for att in attachments:
        if att.get('@odata.type') != '#microsoft.graph.fileAttachment':
            continue  # itemAttachment/referenceAttachment — not a file this policy applies to
        name = att.get('name') or 'attachment'
        size = att.get('size', 0)
        content_bytes = att.get('contentBytes', '')

        def fetch_bytes(cb=content_bytes):
            if not cb:
                return b''
            return base64.b64decode(cb)

        results.append((name, size, fetch_bytes))
    return results

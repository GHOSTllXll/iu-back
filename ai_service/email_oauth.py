# backend/ai_service/email_oauth.py
"""
V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3): the
OAuth 2.0 "Authorization Code" flow for Google (Gmail) and Microsoft
(Microsoft 365 + personal Outlook.com/Hotmail, via the same "common"
tenant app) — a thin, plain-`requests` wrapper, same style as
rentcast_client.py, deliberately WITHOUT the heavier google-api-python-
client/google-auth-oauthlib or msal SDKs: this only ever needs a handful
of REST calls, and a plain-requests implementation is both lighter to
install (relevant on this project's cPanel host — see config/celery.py's
own deploy-environment notes) and easier to test against mocked HTTP
responses without a real Google/Microsoft account.

Every endpoint/parameter below was confirmed against each provider's
OWN current documentation before writing this (not guessed) — see the
per-provider docstrings for exactly what was verified and what's a
documented-by-example-only, not-yet-live-tested best effort (flagged
explicitly where that applies, same "never present a guess as a fact"
discipline as rentcast_client.py).

Both providers are read-only in this phase: the only scope requested is
mail-read (Gmail `gmail.readonly`, Microsoft Graph `Mail.Read`) plus
`email`/`openid` to resolve the connected mailbox's own address for
display. This module never sends mail through a connected account.
"""
import base64
import json
import time
from datetime import timedelta
from urllib.parse import urlencode

import requests
from django.conf import settings
from django.utils import timezone

GOOGLE = 'google'
MICROSOFT = 'microsoft'
PROVIDER_CHOICES = (GOOGLE, MICROSOFT)

REQUEST_TIMEOUT_SECONDS = 15


class EmailOAuthError(Exception):
    """Any failure talking to a provider's OAuth or Graph/Gmail endpoints."""
    pass


def is_google_configured() -> bool:
    return bool(settings.GOOGLE_OAUTH_CLIENT_ID and settings.GOOGLE_OAUTH_CLIENT_SECRET)


def is_microsoft_configured() -> bool:
    return bool(settings.MS_OAUTH_CLIENT_ID and settings.MS_OAUTH_CLIENT_SECRET)


def is_configured(provider: str) -> bool:
    if provider == GOOGLE:
        return is_google_configured()
    if provider == MICROSOFT:
        return is_microsoft_configured()
    return False


def _redirect_uri(provider: str) -> str:
    """
    Built from BACKEND_BASE_URL (see settings.py's own docstring on that
    setting) — must match, character for character, whatever's
    registered as this app's "Authorized redirect URI" in Google Cloud
    Console / this app's "Redirect URI" in its Microsoft Entra App
    Registration. Never the FRONTEND_BASE_URL — the provider redirects
    back to THIS backend's callback view first (EmailOAuthCallbackView
    in views.py), which exchanges the code and only THEN 302s the
    browser on to the frontend's Settings page.
    """
    base = settings.BACKEND_BASE_URL.rstrip('/')
    return f"{base}/api/ai/email-sync/oauth/callback/{provider}/"


# ---------------------------------------------------------------------------
# Google (Gmail)
# ---------------------------------------------------------------------------
# Confirmed against https://developers.google.com/identity/protocols/oauth2/web-server
# and https://developers.google.com/gmail/api/auth/scopes directly before
# writing this:
#   - Authorization endpoint: https://accounts.google.com/o/oauth2/v2/auth
#   - Token endpoint: https://oauth2.googleapis.com/token (code exchange
#     AND refresh both POST here, distinguished only by grant_type)
#   - access_type=offline + prompt=consent on the authorization request
#     is what makes the FIRST code exchange return a refresh_token at
#     all — Google's own docs: omit either and a returning user who
#     already granted consent gets no refresh_token back at all.
#   - Read-only Gmail scope: https://www.googleapis.com/auth/gmail.readonly
#     (documented as a RESTRICTED scope, not just "sensitive" — Google's
#     own page notes a production app storing this data server-side
#     needs to pass Google's security assessment before verification;
#     that's a business/compliance step for whoever runs this in
#     production, not something this code can satisfy itself).
#   - Revocation endpoint: https://oauth2.googleapis.com/revoke (POST,
#     form-encoded `token` param) — Google's own docs page for this was
#     only partially readable during review; this module calls it best-
#     effort (see revoke_google) and NEVER lets a failed/unreachable
#     revoke call block disconnecting locally.
GOOGLE_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth'
GOOGLE_TOKEN_URL = 'https://oauth2.googleapis.com/token'
GOOGLE_REVOKE_URL = 'https://oauth2.googleapis.com/revoke'
GOOGLE_USERINFO_URL = 'https://www.googleapis.com/oauth2/v3/userinfo'
GOOGLE_SCOPES = 'openid email https://www.googleapis.com/auth/gmail.readonly'

GMAIL_API_BASE = 'https://gmail.googleapis.com/gmail/v1/users/me'


def build_google_authorization_url(state: str) -> str:
    params = {
        'client_id': settings.GOOGLE_OAUTH_CLIENT_ID,
        'redirect_uri': _redirect_uri(GOOGLE),
        'response_type': 'code',
        'scope': GOOGLE_SCOPES,
        'access_type': 'offline',
        'prompt': 'consent',
        'state': state,
    }
    return f"{GOOGLE_AUTH_URL}?{urlencode(params)}"


def exchange_google_code(code: str) -> dict:
    """
    Returns {'access_token', 'refresh_token' (may be absent if Google
    didn't issue one — see module docstring), 'expires_at' (a
    timezone-aware datetime), 'scope', 'email_address'}.
    """
    resp = requests.post(
        GOOGLE_TOKEN_URL,
        data={
            'client_id': settings.GOOGLE_OAUTH_CLIENT_ID,
            'client_secret': settings.GOOGLE_OAUTH_CLIENT_SECRET,
            'code': code,
            'grant_type': 'authorization_code',
            'redirect_uri': _redirect_uri(GOOGLE),
        },
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Google token exchange failed ({resp.status_code}): {resp.text[:300]}")
    payload = resp.json()

    email_address = _fetch_google_email(payload['access_token'])

    return {
        'access_token': payload['access_token'],
        'refresh_token': payload.get('refresh_token'),
        'expires_at': timezone.now() + timedelta(seconds=payload.get('expires_in', 3600)),
        'scope': payload.get('scope', ''),
        'email_address': email_address,
    }


def refresh_google_token(refresh_token: str) -> dict:
    """Returns {'access_token', 'expires_at', 'refresh_token' (always None — Google's refresh response never repeats it)."""
    resp = requests.post(
        GOOGLE_TOKEN_URL,
        data={
            'client_id': settings.GOOGLE_OAUTH_CLIENT_ID,
            'client_secret': settings.GOOGLE_OAUTH_CLIENT_SECRET,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
        },
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Google token refresh failed ({resp.status_code}): {resp.text[:300]}")
    payload = resp.json()
    return {
        'access_token': payload['access_token'],
        'expires_at': timezone.now() + timedelta(seconds=payload.get('expires_in', 3600)),
        'refresh_token': None,
    }


def revoke_google(token: str) -> None:
    """Best-effort — see module docstring. Never raises; a failed/unreachable revoke never blocks disconnecting locally."""
    try:
        requests.post(
            GOOGLE_REVOKE_URL,
            data={'token': token},
            headers={'Content-Type': 'application/x-www-form-urlencoded'},
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
    except requests.RequestException:
        pass


def _fetch_google_email(access_token: str) -> str:
    resp = requests.get(
        GOOGLE_USERINFO_URL,
        headers={'Authorization': f'Bearer {access_token}'},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        return ''
    return (resp.json() or {}).get('email', '') or ''


def _gmail_get(access_token: str, path: str, params: dict | None = None) -> dict:
    resp = requests.get(
        f"{GMAIL_API_BASE}{path}",
        headers={'Authorization': f'Bearer {access_token}'},
        params=params or {},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Gmail API {path} failed ({resp.status_code}): {resp.text[:300]}")
    return resp.json()


def search_google_threads(access_token: str, query: str, max_results: int = 20) -> list[dict]:
    """
    GET .../threads?q=<Gmail search syntax>&maxResults=N — confirmed
    endpoint/params. Returns a list of {id, snippet} (the list endpoint
    does NOT include subject/participants — see get_google_thread for
    the follow-up fetch this module's caller does per candidate; kept
    as two calls rather than N+1 full fetches up front, since a search
    is typically narrowed by the analyst before they pick one).
    """
    data = _gmail_get(access_token, '/threads', {'q': query, 'maxResults': max_results})
    return data.get('threads', []) or []


def get_google_thread(access_token: str, thread_id: str) -> dict:
    """GET .../threads/{id}?format=full — full message bodies/attachment-part metadata, confirmed endpoint/params."""
    return _gmail_get(access_token, f'/threads/{thread_id}', {'format': 'full'})


def get_google_attachment(access_token: str, message_id: str, attachment_id: str) -> bytes:
    """GET .../messages/{messageId}/attachments/{id} — confirmed endpoint. `data` is base64URL-encoded."""
    data = _gmail_get(access_token, f'/messages/{message_id}/attachments/{attachment_id}')
    raw = data.get('data', '')
    if not raw:
        return b''
    return base64.urlsafe_b64decode(raw + '=' * (-len(raw) % 4))


# ---------------------------------------------------------------------------
# Microsoft (Microsoft 365 + personal Outlook.com/Hotmail)
# ---------------------------------------------------------------------------
# Confirmed against https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow,
# .../v2-protocols-oidc, https://learn.microsoft.com/en-us/graph/search-query-parameter,
# and the Graph attachment-get/message-list-attachments pages directly:
#   - 'common' tenant supports BOTH work-or-school AND personal Microsoft
#     accounts through the SAME app registration (confirmed on the OIDC
#     page) — the app registration's own "supported account types" must
#     also be set to allow personal accounts, which is a console setting
#     this code can't verify from here.
#   - Authorize: https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize
#   - Token: https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
#     (code exchange AND refresh both POST here)
#   - offline_access MUST be in the requested scope to get a
#     refresh_token back at all (confirmed).
#   - Mail.Read is confirmed (from the Graph attachment-get page's own
#     permissions table) to work for BOTH "Delegated (work or school
#     account)" and "Delegated (personal Microsoft account)" — no
#     separate scope needed for personal accounts.
#   - GET /me/messages?$search="..." on messages is confirmed with a
#     verbatim example from Microsoft's own $search docs page.
#   - GET /me/messages?$filter=conversationId eq '<id>' to pull every
#     message in one conversation is backed by a Microsoft Q&A answer
#     using the Graph SDK's equivalent .Filter("conversationid eq ...")
#     call, NOT a verbatim example from Microsoft's own reference page
#     (that page documents $filter/$orderby interaction rules but never
#     names conversationId specifically) — treated here as a reasonable,
#     standard OData filter on a documented string property, but FLAGGED
#     as not independently confirmed against a live Graph call. If this
#     ever starts failing against a real connected account, the sync
#     task (email_sync.py) catches it per-thread, marks that thread's
#     poll as a logged failure, and never crashes the whole 5-minute run
#     over one thread.
#   - There is no documented endpoint to revoke a single refresh token
#     server-side on Microsoft's v2.0 endpoint — disconnecting simply
#     discards this app's own stored tokens (see revoke_microsoft,
#     which is a deliberate no-op) and tells the user they can also
#     remove this app's access from https://account.live.com/consent/Manage
#     or their organization's My Apps page if they want to revoke it on
#     Microsoft's side too.
MS_SCOPES = 'openid email offline_access https://graph.microsoft.com/Mail.Read'
GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0'


def _ms_authorize_url() -> str:
    return f"https://login.microsoftonline.com/{settings.MS_OAUTH_TENANT}/oauth2/v2.0/authorize"


def _ms_token_url() -> str:
    return f"https://login.microsoftonline.com/{settings.MS_OAUTH_TENANT}/oauth2/v2.0/token"


def build_microsoft_authorization_url(state: str) -> str:
    params = {
        'client_id': settings.MS_OAUTH_CLIENT_ID,
        'redirect_uri': _redirect_uri(MICROSOFT),
        'response_type': 'code',
        'scope': MS_SCOPES,
        'response_mode': 'query',
        'state': state,
    }
    return f"{_ms_authorize_url()}?{urlencode(params)}"


def exchange_microsoft_code(code: str) -> dict:
    resp = requests.post(
        _ms_token_url(),
        data={
            'client_id': settings.MS_OAUTH_CLIENT_ID,
            'client_secret': settings.MS_OAUTH_CLIENT_SECRET,
            'code': code,
            'grant_type': 'authorization_code',
            'redirect_uri': _redirect_uri(MICROSOFT),
            'scope': MS_SCOPES,
        },
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Microsoft token exchange failed ({resp.status_code}): {resp.text[:300]}")
    payload = resp.json()

    email_address = _fetch_microsoft_email(payload['access_token'])

    return {
        'access_token': payload['access_token'],
        'refresh_token': payload.get('refresh_token'),
        'expires_at': timezone.now() + timedelta(seconds=payload.get('expires_in', 3600)),
        'scope': payload.get('scope', ''),
        'email_address': email_address,
    }


def refresh_microsoft_token(refresh_token: str) -> dict:
    """Returns {'access_token', 'expires_at', 'refresh_token' (may be a NEW one Microsoft wants the old one replaced with — confirmed in docs)."""
    resp = requests.post(
        _ms_token_url(),
        data={
            'client_id': settings.MS_OAUTH_CLIENT_ID,
            'client_secret': settings.MS_OAUTH_CLIENT_SECRET,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
            'scope': MS_SCOPES,
        },
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Microsoft token refresh failed ({resp.status_code}): {resp.text[:300]}")
    payload = resp.json()
    return {
        'access_token': payload['access_token'],
        'expires_at': timezone.now() + timedelta(seconds=payload.get('expires_in', 3600)),
        'refresh_token': payload.get('refresh_token'),
    }


def revoke_microsoft(token: str) -> None:
    """No-op — see module docstring (no documented server-side single-token revocation on the v2.0 endpoint)."""
    pass


def _fetch_microsoft_email(access_token: str) -> str:
    resp = requests.get(
        f"{GRAPH_API_BASE}/me",
        headers={'Authorization': f'Bearer {access_token}'},
        params={'$select': 'mail,userPrincipalName'},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        return ''
    data = resp.json() or {}
    return data.get('mail') or data.get('userPrincipalName') or ''


def _graph_get(access_token: str, path: str, params: dict | None = None) -> dict:
    resp = requests.get(
        f"{GRAPH_API_BASE}{path}",
        headers={'Authorization': f'Bearer {access_token}'},
        params=params or {},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if resp.status_code != 200:
        raise EmailOAuthError(f"Graph API {path} failed ({resp.status_code}): {resp.text[:300]}")
    return resp.json()


def search_microsoft_threads(access_token: str, query: str, max_results: int = 20) -> list[dict]:
    """
    GET /me/messages?$search="<query>"&$top=N — confirmed endpoint/
    param for $search. Returns raw message objects (each one IS a
    candidate "thread" by its own conversationId — see email_sync.py's
    caller, which dedupes to one row per conversationId, keeping the
    most recent message, before showing the analyst a pick list).
    """
    params = {'$top': max_results, '$select': 'id,conversationId,subject,from,receivedDateTime,bodyPreview'}
    if query:
        params['$search'] = f'"{query}"'
    else:
        params['$orderby'] = 'receivedDateTime desc'
    data = _graph_get(access_token, '/me/messages', params)
    return data.get('value', []) or []


def get_microsoft_conversation_messages(access_token: str, conversation_id: str) -> list[dict]:
    """
    GET /me/messages?$filter=conversationId eq '<id>' — see this
    module's own docstring above on why this specific call is flagged
    as not independently confirmed against a live account.
    """
    params = {
        '$filter': f"conversationId eq '{conversation_id}'",
        '$orderby': 'receivedDateTime asc',
        '$select': 'id,conversationId,subject,from,toRecipients,receivedDateTime,bodyPreview,body,hasAttachments',
    }
    data = _graph_get(access_token, '/me/messages', params)
    return data.get('value', []) or []


def list_microsoft_attachments(access_token: str, message_id: str) -> list[dict]:
    """GET /me/messages/{id}/attachments — confirmed endpoint. File attachments include `contentBytes` (base64) directly, no second call needed."""
    data = _graph_get(access_token, f'/me/messages/{message_id}/attachments')
    return data.get('value', []) or []

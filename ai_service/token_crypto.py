# backend/ai_service/token_crypto.py
"""
V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3):
symmetric encryption for the OAuth access/refresh tokens
EmailOAuthConnection stores (see models.py) -- the client's explicit
"encrypted token storage" requirement. Uses `cryptography`'s Fernet
(AES-128-CBC + HMAC-SHA256, authenticated) -- already a dependency of
this project (requirements.txt pins cryptography==49.0.0 for other
uses), so this adds no new package.

Key comes from settings.EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY (env var
EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY) -- a urlsafe-base64, 32-byte Fernet
key, generated once with `Fernet.generate_key()` and set as a secret,
exactly the same "blank by default, feature quietly unavailable until
set" pattern as RENTCAST_API_KEY (see rentcast_client.is_configured()).
Deliberately a SEPARATE key from SECRET_KEY/SIMPLE_JWT's signing key --
rotating this key alone (by design, without touching Django's session/
JWT signing) is exactly the operation needed if it's ever compromised,
and reusing SECRET_KEY for storage encryption is a well-known anti-
pattern this module avoids on purpose.

Rotation: there is deliberately NO "re-encrypt everything" migration
here in v1 -- if EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY is ever rotated,
every existing EmailOAuthConnection's stored tokens become undecryptable
and that connection moves to status='error' on its next use (see
decrypt_token's caller in email_sync.py), surfacing to the affected
user as "reconnect your mailbox" -- the same recovery path as a
provider-side token revocation, not a new failure mode to handle
specially.
"""
from functools import lru_cache

from cryptography.fernet import Fernet, InvalidToken
from django.conf import settings


class TokenEncryptionUnavailable(Exception):
    """Raised when EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY isn't configured."""
    pass


class TokenDecryptionFailed(Exception):
    """
    Raised when a stored token can't be decrypted with the CURRENT key --
    either EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY was rotated/changed since this
    token was encrypted, or the stored ciphertext is corrupt. Callers
    should treat this exactly like an expired/revoked token: mark the
    connection status='error' and prompt a reconnect, never crash.
    """
    pass


@lru_cache(maxsize=1)
def _get_fernet() -> Fernet:
    key = getattr(settings, 'EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY', '') or ''
    if not key:
        raise TokenEncryptionUnavailable(
            "EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY isn't configured -- Email Thread "
            "Sync can't store OAuth tokens until it is. Generate one with "
            "`python -c \"from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())\"` "
            "and set it as an env var."
        )
    try:
        return Fernet(key.encode('utf-8') if isinstance(key, str) else key)
    except Exception as e:
        raise TokenEncryptionUnavailable(
            f"EMAIL_OAUTH_TOKEN_ENCRYPTION_KEY is set but isn't a valid Fernet key: {e}"
        )


def is_configured() -> bool:
    try:
        _get_fernet()
        return True
    except TokenEncryptionUnavailable:
        return False


def encrypt_token(plaintext: str) -> str:
    """Returns an opaque urlsafe-base64 ciphertext string, safe to store in a CharField/TextField."""
    if plaintext is None:
        return ''
    fernet = _get_fernet()
    return fernet.encrypt(plaintext.encode('utf-8')).decode('utf-8')


def decrypt_token(ciphertext: str) -> str:
    if not ciphertext:
        return ''
    fernet = _get_fernet()
    try:
        return fernet.decrypt(ciphertext.encode('utf-8')).decode('utf-8')
    except InvalidToken as e:
        raise TokenDecryptionFailed(str(e))

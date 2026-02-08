"""
Microsoft Graph API authentication via MSAL device code flow.

Provides token acquisition, caching, and refresh for accessing OneNote
via Microsoft Graph REST API.  Tokens are stored at
~/.cache/onenote-mcp/graph-token.json and automatically refreshed.
"""

import logging
import os
import sys
import threading
from pathlib import Path

import msal

log = logging.getLogger("onenote-mcp")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Microsoft Graph Explorer public client ID (no app registration needed).
_DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", _DEFAULT_CLIENT_ID)

SCOPES = [
    "Notes.Read",
    "Notes.ReadWrite",
    "Notes.Create",
    "User.Read",
]

AUTHORITY = "https://login.microsoftonline.com/consumers"

_CACHE_DIR = Path.home() / ".cache" / "onenote-mcp"
_TOKEN_PATH = _CACHE_DIR / "graph-token.json"

# ---------------------------------------------------------------------------
# MSAL token cache (persisted to disk)
# ---------------------------------------------------------------------------

_msal_cache = msal.SerializableTokenCache()

# Background auth state
_pending_flow: dict | None = None
_auth_thread: threading.Thread | None = None
_auth_error: str | None = None


def _load_cache() -> None:
    """Load the MSAL token cache from disk."""
    if _TOKEN_PATH.exists():
        try:
            _msal_cache.deserialize(_TOKEN_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning("Failed to load token cache: %s", e)


def _save_cache() -> None:
    """Persist the MSAL token cache to disk if it has changed."""
    if _msal_cache.has_state_changed:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        _TOKEN_PATH.write_text(_msal_cache.serialize(), encoding="utf-8")
        log.debug("Token cache saved to %s", _TOKEN_PATH)


def _build_app() -> msal.PublicClientApplication:
    """Build a public client application with the persistent cache."""
    _load_cache()
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=_msal_cache,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def get_access_token() -> str | None:
    """Return a valid access token, refreshing silently if needed.

    Returns None if no cached account exists or refresh fails.
    """
    app = _build_app()
    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache()

    if result and "access_token" in result:
        return result["access_token"]

    if result and "error" in result:
        log.warning("Token refresh failed: %s — %s",
                    result.get("error"), result.get("error_description"))
    return None


def _poll_for_token(app: msal.PublicClientApplication, flow: dict) -> None:
    """Background thread: poll until user completes sign-in."""
    global _auth_error
    try:
        result = app.acquire_token_by_device_flow(flow)
        _save_cache()
        if "access_token" in result:
            log.info("Background auth completed successfully")
            _auth_error = None
        else:
            _auth_error = result.get("error_description", result.get("error", "Unknown error"))
            log.warning("Background auth failed: %s", _auth_error)
    except Exception as e:
        _auth_error = str(e)
        log.warning("Background auth exception: %s", e)


def authenticate() -> str:
    """Start the device code flow.

    Returns immediately with the code and URL for the user to complete
    sign-in in a browser. Token acquisition polls in the background.
    Call check_auth() to verify completion.
    """
    global _pending_flow, _auth_thread, _auth_error

    app = _build_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device code flow failed: {flow.get('error_description', flow)}")

    _pending_flow = flow
    _auth_error = None

    # Start polling in background so the tool can return immediately
    _auth_thread = threading.Thread(target=_poll_for_token, args=(app, flow), daemon=True)
    _auth_thread.start()

    user_code = flow["user_code"]
    url = flow.get("verification_uri", "https://microsoft.com/devicelogin")
    log.info("Device code flow started: code=%s url=%s", user_code, url)

    return (
        f"Open this URL in your browser: {url}\n"
        f"Enter code: {user_code}\n\n"
        f"After signing in, use the 'check_auth' tool to verify."
    )


def check_auth() -> str:
    """Check if background authentication has completed."""
    global _auth_thread, _pending_flow

    # First check if we already have a valid token
    token = get_access_token()
    if token:
        _pending_flow = None
        _auth_thread = None
        return "Authenticated successfully! API features are now available."

    if _auth_thread is not None and _auth_thread.is_alive():
        code = _pending_flow.get("user_code", "?") if _pending_flow else "?"
        url = _pending_flow.get("verification_uri", "https://microsoft.com/devicelogin") if _pending_flow else "?"
        return (
            f"Still waiting for sign-in...\n"
            f"URL: {url}\n"
            f"Code: {code}\n\n"
            f"Complete sign-in in your browser, then run 'check_auth' again."
        )

    if _auth_error:
        return f"Authentication failed: {_auth_error}\nRun 'authenticate' to try again."

    return "No authentication in progress. Run 'authenticate' to start."


def clear_token() -> str:
    """Remove cached tokens (for re-authentication)."""
    if _TOKEN_PATH.exists():
        _TOKEN_PATH.unlink()
        log.info("Token cache removed: %s", _TOKEN_PATH)
        return f"Token cache removed: {_TOKEN_PATH}"
    return "No token cache to remove."

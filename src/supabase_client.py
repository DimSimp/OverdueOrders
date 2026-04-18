from __future__ import annotations

from typing import Optional

_client = None


def init_client(url: str, key: str) -> None:
    """Initialize the Supabase client. Call once at app startup after config is loaded."""
    global _client
    from supabase import create_client
    _client = create_client(url, key)


def get_client():
    """Return the initialized Supabase client.

    Raises RuntimeError if init_client() has not been called yet.
    """
    if _client is None:
        raise RuntimeError(
            "Supabase client is not initialized. "
            "Add supabase.url and supabase.key to config.json."
        )
    return _client


def is_initialized() -> bool:
    """Return True if the client has been initialized."""
    return _client is not None

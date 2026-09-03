"""JWT decoding helpers."""

from __future__ import annotations

import json
import time
from base64 import urlsafe_b64decode


def _decode_audience(token: str) -> str:
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return "unknown"
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        decoded = json.loads(urlsafe_b64decode(payload))
        return decoded.get("aud", "unknown")
    except Exception:
        return "unknown"


def _decode_exp(token: str) -> float:
    """Extract exp claim from JWT without full verification."""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return time.time() + 3600
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        decoded = json.loads(urlsafe_b64decode(payload))
        return float(decoded.get("exp", time.time() + 3600))
    except Exception:
        return time.time() + 3600

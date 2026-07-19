"""Shared-token authentication guard for API routes."""
from __future__ import annotations

import hmac
import os
from functools import wraps
from typing import Callable, TypeVar

from flask import jsonify, request

F = TypeVar("F", bound=Callable[..., object])


def require_service_token(fn: F) -> F:
    """Gate a route behind Authorization: Bearer <RECAP_SERVICE_TOKEN>.

    Fails closed — an unset token never authorizes anything. This is the same
    shared secret the dashboard uses for its Flask -> Next.js webhook, reused
    here for the Next.js -> Flask direction."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        token = os.environ.get('RECAP_SERVICE_TOKEN')
        auth = request.headers.get('Authorization', '')
        if not token or not auth.startswith('Bearer ') or not hmac.compare_digest(auth[7:], token):
            return jsonify({'error': 'Unauthorized'}), 401
        return fn(*args, **kwargs)
    return wrapper  # type: ignore[return-value]

"""
Session-based auth. Single admin user from env for MVP-1.

Env:
  DUBSTUDIO_ADMIN_EMAIL (default: admin@local)
  DUBSTUDIO_ADMIN_PASSWORD (default: admin)
  DUBSTUDIO_SESSION_SECRET (generated at startup if missing)
"""
from __future__ import annotations

import os
import secrets

from fastapi import Request


ADMIN_EMAIL = os.environ.get("DUBSTUDIO_ADMIN_EMAIL", "admin@local")
ADMIN_PASSWORD = os.environ.get("DUBSTUDIO_ADMIN_PASSWORD", "admin")
SESSION_SECRET = os.environ.get("DUBSTUDIO_SESSION_SECRET") or secrets.token_urlsafe(32)


def is_authenticated(request: Request) -> bool:
    return request.session.get("user") == ADMIN_EMAIL


def login(request: Request, email: str, password: str) -> bool:
    if email == ADMIN_EMAIL and password == ADMIN_PASSWORD:
        request.session["user"] = email
        return True
    return False


def logout(request: Request) -> None:
    request.session.clear()

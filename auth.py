"""
Сессионная авторизация поверх таблицы User. При первом запуске, если таблица
пустая, сидируется bootstrap-админ из env-дефолтов (см. _bootstrap_admin).

Env:
  DUBSTUDIO_ADMIN_EMAIL  (default: Sofi)
  DUBSTUDIO_ADMIN_PASSWORD (default: C@rrots!)
  DUBSTUDIO_SESSION_SECRET (fallback — автогенерится в .session_secret)
"""

from __future__ import annotations

import os
import secrets
from pathlib import Path
from typing import Optional

import bcrypt
from fastapi import Request
from sqlmodel import Session, select

from db import engine
from models import User


BOOTSTRAP_EMAIL = os.environ.get("DUBSTUDIO_ADMIN_EMAIL", "Sofi")
BOOTSTRAP_PASSWORD = os.environ.get("DUBSTUDIO_ADMIN_PASSWORD", "C@rrots!")


def _session_secret() -> str:
    env = os.environ.get("DUBSTUDIO_SESSION_SECRET")
    if env:
        return env
    p = Path(__file__).parent / ".session_secret"
    if p.exists():
        return p.read_text().strip()
    s = secrets.token_urlsafe(32)
    p.write_text(s)
    return s


SESSION_SECRET = _session_secret()


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), hashed.encode("utf-8"))
    except (ValueError, TypeError):
        return False


def seed_bootstrap_admin() -> None:
    with Session(engine) as s:
        existing = s.exec(select(User)).first()
        if existing:
            return
        admin = User(
            email=BOOTSTRAP_EMAIL,
            password_hash=hash_password(BOOTSTRAP_PASSWORD),
            role="admin",
        )
        s.add(admin)
        s.commit()


def is_authenticated(request: Request) -> bool:
    return bool(request.session.get("user_id"))


def current_role(request: Request) -> Optional[str]:
    return request.session.get("role")


def is_admin(request: Request) -> bool:
    return current_role(request) == "admin"


def login(request: Request, email: str, password: str) -> bool:
    email = (email or "").strip()
    if not email or not password:
        return False
    with Session(engine) as s:
        user = s.exec(select(User).where(User.email == email)).first()
        if not user or not verify_password(password, user.password_hash):
            return False
        request.session["user_id"] = user.id
        request.session["user"] = user.email
        request.session["role"] = user.role
    return True


def logout(request: Request) -> None:
    request.session.clear()

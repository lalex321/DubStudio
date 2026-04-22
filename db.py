from __future__ import annotations

import os
from pathlib import Path

from sqlmodel import Session, SQLModel, create_engine

_DB_URL = os.environ.get("DUBSTUDIO_DB_URL", f"sqlite:///{Path(__file__).parent / 'dubstudio.db'}")

engine = create_engine(
    _DB_URL,
    echo=False,
    connect_args={"check_same_thread": False} if _DB_URL.startswith("sqlite") else {},
)


def init_db() -> None:
    SQLModel.metadata.create_all(engine)


def get_session() -> Session:
    return Session(engine)

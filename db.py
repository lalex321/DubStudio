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
    _migrate_sqlite()


def _migrate_sqlite() -> None:
    """Лёгкие миграции: SQLModel.create_all не добавляет новые колонки к
    существующим таблицам. Делаем ALTER TABLE вручную для известных полей."""
    if not _DB_URL.startswith("sqlite"):
        return
    from sqlalchemy import text

    with engine.begin() as conn:
        cols = [r[1] for r in conn.execute(text("PRAGMA table_info(wordcount)"))]
        if "edited" not in cols:
            conn.execute(text("ALTER TABLE wordcount ADD COLUMN edited BOOLEAN DEFAULT 0"))


def get_session() -> Session:
    return Session(engine)

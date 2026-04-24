from __future__ import annotations

import os

from sqlmodel import Session, SQLModel, create_engine

from paths import DATA_DIR

_DB_URL = os.environ.get("DUBSTUDIO_DB_URL", f"sqlite:///{DATA_DIR / 'dubstudio.db'}")

engine = create_engine(
    _DB_URL,
    echo=False,
    connect_args={"check_same_thread": False} if _DB_URL.startswith("sqlite") else {},
)


def init_db() -> None:
    SQLModel.metadata.create_all(engine)
    _migrate_sqlite()
    _cleanup_junk_characters()


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
        char_cols = [r[1] for r in conn.execute(text("PRAGMA table_info(character)"))]
        if "acknowledged" not in char_cols:
            conn.execute(text("ALTER TABLE character ADD COLUMN acknowledged BOOLEAN DEFAULT 1"))
        actor_cols = [r[1] for r in conn.execute(text("PRAGMA table_info(actor)"))]
        if "tg" not in actor_cols:
            conn.execute(text("ALTER TABLE actor ADD COLUMN tg TEXT DEFAULT ''"))


# Netflix sheets contain rows for "PRINCIPAL PHOTOGRAPHY", "GRAPHICS INSERTS",
# "MAIN TITLE" that were ingested as characters before the import filter
# was added. Kept in sync with _JUNK_CHARACTER_NAMES in app.py.
_JUNK_CHARACTER_NAMES = ("PRINCIPAL PHOTOGRAPHY", "GRAPHICS INSERTS", "MAIN TITLE")


def _cleanup_junk_characters() -> None:
    """One-shot cleanup of production-metadata rows that slipped into the
    character table before the import filter existed. Idempotent — after
    the first run the SELECTs return empty on every start."""
    from sqlalchemy import text

    with engine.begin() as conn:
        for name in _JUNK_CHARACTER_NAMES:
            ids = [
                r[0] for r in conn.execute(
                    text("SELECT id FROM character WHERE UPPER(TRIM(name)) = :n"),
                    {"n": name},
                )
            ]
            for cid in ids:
                conn.execute(
                    text("DELETE FROM wordcount WHERE character_id = :cid"),
                    {"cid": cid},
                )
                conn.execute(
                    text("DELETE FROM assignment WHERE character_id = :cid"),
                    {"cid": cid},
                )
                conn.execute(
                    text("DELETE FROM character WHERE id = :cid"),
                    {"cid": cid},
                )


def get_session() -> Session:
    return Session(engine)

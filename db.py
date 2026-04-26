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
    _migrate_session_actors()
    _migrate_assignment_multi_actor()
    _cleanup_junk_characters()
    _seed_default_room()


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
        proj_cols = [r[1] for r in conn.execute(text("PRAGMA table_info(project)"))]
        if "end_date" not in proj_cols:
            conn.execute(text("ALTER TABLE project ADD COLUMN end_date DATE"))
        if "color" not in proj_cols:
            conn.execute(text("ALTER TABLE project ADD COLUMN color TEXT DEFAULT ''"))


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


def _migrate_session_actors() -> None:
    """RecordingSession.actor_id was a single FK in the first cut of the
    calendar feature; sessions are now M:N with actors via SessionActor.
    Detect the legacy column, fan out its values into SessionActor, then
    rebuild the recordingsession table without the column.

    SQLite doesn't support DROP COLUMN before 3.35 in a way SQLModel can
    drive, so we do the safe table-rebuild dance:
      1) read all old rows (including actor_id)
      2) DROP TABLE recordingsession
      3) recreate via SQLModel.metadata.create_all on the new schema
      4) re-insert rows + matching SessionActor entries
    """
    if not _DB_URL.startswith("sqlite"):
        return
    from sqlalchemy import text

    with engine.begin() as conn:
        cols = [r[1] for r in conn.execute(text("PRAGMA table_info(recordingsession)"))]
        if not cols:
            return  # table doesn't exist yet — fresh install, nothing to migrate
        if "actor_id" not in cols:
            return  # already on the new schema

        rows = list(conn.execute(text(
            "SELECT id, project_id, actor_id, room_id, starts_at, ends_at, "
            "status, target_words, episode_numbers, notes, created_at, updated_at "
            "FROM recordingsession"
        )))
        conn.execute(text("DROP TABLE recordingsession"))

    # Recreate the table from the current model definitions.
    SQLModel.metadata.create_all(engine)

    if not rows:
        return
    with engine.begin() as conn:
        for r in rows:
            conn.execute(
                text(
                    "INSERT INTO recordingsession "
                    "(id, project_id, room_id, starts_at, ends_at, status, "
                    "target_words, episode_numbers, notes, created_at, updated_at) "
                    "VALUES (:id, :project_id, :room_id, :starts_at, :ends_at, "
                    ":status, :target_words, :episode_numbers, :notes, "
                    ":created_at, :updated_at)"
                ),
                {
                    "id": r[0], "project_id": r[1], "room_id": r[3],
                    "starts_at": r[4], "ends_at": r[5], "status": r[6],
                    "target_words": r[7], "episode_numbers": r[8],
                    "notes": r[9], "created_at": r[10], "updated_at": r[11],
                },
            )
            if r[2] is not None:  # had an actor_id — preserve it
                conn.execute(
                    text(
                        "INSERT INTO sessionactor (session_id, actor_id) "
                        "VALUES (:sid, :aid)"
                    ),
                    {"sid": r[0], "aid": r[2]},
                )


def _migrate_assignment_multi_actor() -> None:
    """Originally Assignment had UNIQUE(project_id, character_id) — one
    actor per character. Массовка scenes need multiple actors per
    character, so the constraint is now UNIQUE(character_id, actor_id).
    Detect the old shape by inspecting the table's CREATE SQL and rebuild
    if needed. Existing rows transfer 1:1.
    """
    if not _DB_URL.startswith("sqlite"):
        return
    from sqlalchemy import text

    with engine.begin() as conn:
        row = conn.execute(text(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='assignment'"
        )).first()
        if not row:
            return
        ddl = (row[0] or "").lower().replace(" ", "")
        # Legacy constraint was UNIQUE(project_id, character_id). If the
        # CREATE SQL no longer mentions that pair, we're already on the
        # new shape — nothing to do.
        if "project_id,character_id" not in ddl:
            return

        rows = list(conn.execute(text(
            "SELECT id, project_id, character_id, actor_id FROM assignment"
        )))
        conn.execute(text("DROP TABLE assignment"))

    SQLModel.metadata.create_all(engine)

    if not rows:
        return
    with engine.begin() as conn:
        seen: set[tuple[int, int]] = set()
        for r in rows:
            key = (r[2], r[3])
            if key in seen:
                continue
            seen.add(key)
            conn.execute(
                text(
                    "INSERT INTO assignment (id, project_id, character_id, actor_id) "
                    "VALUES (:id, :pid, :cid, :aid)"
                ),
                {"id": r[0], "pid": r[1], "cid": r[2], "aid": r[3]},
            )


def _seed_default_room() -> None:
    """Create the first room on a fresh install so the calendar picker
    isn't empty. Idempotent — does nothing if any room already exists."""
    from sqlalchemy import text

    with engine.begin() as conn:
        existing = conn.execute(text("SELECT COUNT(*) FROM room")).scalar() or 0
        if existing == 0:
            conn.execute(
                text(
                    "INSERT INTO room (name, sort_order, color) "
                    "VALUES ('Studio A', 0, '')"
                )
            )


def get_session() -> Session:
    return Session(engine)

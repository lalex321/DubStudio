from __future__ import annotations

from datetime import date, datetime, timezone
from typing import Optional

from sqlmodel import Field, SQLModel, UniqueConstraint


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _today() -> date:
    return date.today()


class User(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    email: str = Field(unique=True, index=True)
    password_hash: str
    role: str = "assistant_director"
    created_at: datetime = Field(default_factory=_utcnow)


class Project(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    number: int = Field(unique=True, index=True)
    title: str
    start_date: date = Field(default_factory=_today)
    # Optional production end date; used to render a horizontal band on
    # the planning calendar from start_date to end_date. Null = open.
    end_date: Optional[date] = None
    # Hex color used as the calendar event/band tint. Empty string means
    # auto-derived from the project id at render time.
    color: str = ""
    created_at: datetime = Field(default_factory=_utcnow)


class Episode(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("project_id", "number"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: int = Field(foreign_key="project.id", index=True)
    number: int
    uploaded_filename: str
    imported_at: datetime = Field(default_factory=_utcnow)


class Character(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("project_id", "name"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: int = Field(foreign_key="project.id", index=True)
    name: str
    acknowledged: bool = True


class Actor(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    name: str = Field(unique=True, index=True)
    # Telegram handle without a leading @. Used later for notification
    # delivery; empty string means no handle captured yet.
    tg: str = ""


class Assignment(SQLModel, table=True):
    """Character ↔ Actor link inside a project. A character can be cast
    with multiple actors (массовка / групповые сцены), so the uniqueness
    is on (character_id, actor_id) — same actor can't be added twice to
    the same character — but a character can have many actor rows."""

    __table_args__ = (UniqueConstraint("character_id", "actor_id"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: int = Field(foreign_key="project.id", index=True)
    character_id: int = Field(foreign_key="character.id", index=True)
    actor_id: int = Field(foreign_key="actor.id", index=True)


class WordCount(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("episode_id", "character_id"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    episode_id: int = Field(foreign_key="episode.id", index=True)
    character_id: int = Field(foreign_key="character.id", index=True)
    dialog_wc: int = 0
    transcription_wc: int = 0
    total_wc: int = 0
    edited: bool = False


class Room(SQLModel, table=True):
    """Studio recording room. Sessions are booked into a room (or none,
    for an unassigned booking)."""

    id: Optional[int] = Field(default=None, primary_key=True)
    name: str = Field(unique=True, index=True)
    sort_order: int = 0
    color: str = ""


SESSION_STATUSES = ("planned", "done", "cancelled", "no_show")


class RecordingSession(SQLModel, table=True):
    """A scheduled recording slot in the studio. starts_at / ends_at are
    stored UTC; the calendar renders them in the browser's timezone.

    Actors are stored in the SessionActor junction table — a session can
    have one or more actors (e.g. a dialogue scene with both speakers
    booked together).

    episode_numbers is a comma-separated list (e.g. "3,5,7") for the
    episodes this session covers — kept as a plain string until we
    actually need to query inside it. notes is freeform helper text."""

    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: int = Field(foreign_key="project.id", index=True)
    room_id: Optional[int] = Field(default=None, foreign_key="room.id", index=True)
    starts_at: datetime = Field(index=True)
    ends_at: datetime
    status: str = "planned"
    target_words: Optional[int] = None
    episode_numbers: str = ""
    notes: str = ""
    created_at: datetime = Field(default_factory=_utcnow)
    updated_at: datetime = Field(default_factory=_utcnow)


class SessionActor(SQLModel, table=True):
    """Many-to-many between RecordingSession and Actor: one session can
    book multiple actors (typical for dialogue scenes or group walla)."""

    __table_args__ = (UniqueConstraint("session_id", "actor_id"),)

    id: Optional[int] = Field(default=None, primary_key=True)
    session_id: int = Field(foreign_key="recordingsession.id", index=True)
    actor_id: int = Field(foreign_key="actor.id", index=True)

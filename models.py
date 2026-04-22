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


class Actor(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    name: str = Field(unique=True, index=True)


class Assignment(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("project_id", "character_id"),)

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

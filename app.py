from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Optional
from urllib.parse import quote

from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func
from sqlmodel import Session, select
from starlette.middleware.sessions import SessionMiddleware

from auth import (
    SESSION_SECRET,
    hash_password,
    is_admin,
    is_authenticated,
    login,
    logout,
    seed_bootstrap_admin,
)
from db import engine, init_db
from models import (
    Actor,
    Assignment,
    Character,
    Episode,
    Project,
    RecordingSession,
    Room,
    SESSION_STATUSES,
    SessionActor,
    User,
    WordCount,
)
from parser import (
    PROFILES,
    EpisodeData,
    derive_common_name,
    derive_show_title,
)
from paths import BUNDLE_DIR
from writer import build_actor_report_xlsx, build_project_xlsx


BASE_DIR = BUNDLE_DIR
TEMPLATES = Jinja2Templates(directory=str(BASE_DIR / "templates"))

_SECURE_COOKIES = os.environ.get("DUBSTUDIO_SECURE_COOKIES", "").lower() in ("1", "true", "yes")

app = FastAPI(title="DubStudio")
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    https_only=_SECURE_COOKIES,
    same_site="lax",
)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


def _flash(request: Request, message: str, level: str = "error") -> None:
    request.session["flash"] = {"msg": message, "level": level}


def _pop_flash(request: Request):
    f = request.session.pop("flash", None)
    if f is None:
        return None
    if isinstance(f, str):  # legacy payloads left in session
        return {"msg": f, "level": "error"}
    return f


def _flash_import_summary(
    request: Request, files_total: int, episodes_imported: int, warnings: list[str]
) -> None:
    if episodes_imported == 0:
        msg = f"Import cancelled: none of {files_total} files could be recognized as episodes."
        if warnings:
            msg += " Reasons: " + "; ".join(warnings[:5])
        _flash(request, msg, level="error")
        return
    msg = f"Imported episodes: {episodes_imported} of {files_total} files."
    level = "success"
    if warnings:
        msg += " Warnings: " + "; ".join(warnings[:5])
        if len(warnings) > 5:
            msg += f" (+{len(warnings) - 5})"
        level = "info"
    _flash(request, msg, level=level)


def _attachment_header(filename: str) -> str:
    ascii_fallback = filename.encode("ascii", "replace").decode("ascii").replace('"', "")
    return f'attachment; filename="{ascii_fallback}"; filename*=UTF-8\'\'{quote(filename)}'


@app.on_event("startup")
def _startup() -> None:
    init_db()
    seed_bootstrap_admin()


def _require_auth(request: Request) -> Optional[RedirectResponse]:
    if not is_authenticated(request):
        return RedirectResponse("/login", status_code=303)
    return None


def _require_admin(request: Request):
    if not is_authenticated(request):
        return RedirectResponse("/login", status_code=303)
    if not is_admin(request):
        raise HTTPException(403, "Admins only")
    return None


def _render(request: Request, name: str, ctx: dict) -> HTMLResponse:
    ctx = {
        "user": request.session.get("user"),
        "role": request.session.get("role"),
        "is_admin": is_admin(request),
        "flash": _pop_flash(request),
        **ctx,
    }
    return TEMPLATES.TemplateResponse(request, name, ctx)


# ---------- root / auth ----------
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    if is_authenticated(request):
        return RedirectResponse("/projects", status_code=303)
    return RedirectResponse("/login", status_code=303)


@app.get("/login", response_class=HTMLResponse)
async def login_form(request: Request):
    if is_authenticated(request):
        return RedirectResponse("/projects", status_code=303)
    return _render(request, "login.html", {"title": "Sign in"})


@app.post("/login", response_class=HTMLResponse)
async def login_submit(
    request: Request,
    email: str = Form(...),
    password: str = Form(...),
):
    if login(request, email, password):
        return RedirectResponse("/projects", status_code=303)
    return _render(
        request,
        "login.html",
        {"title": "Sign in", "error": "Wrong username or password", "email": email},
    )


@app.post("/logout")
async def logout_route(request: Request):
    logout(request)
    return RedirectResponse("/login", status_code=303)


@app.get("/healthz")
async def healthz() -> dict:
    return {"ok": True}


# ---------- projects list ----------
@app.get("/projects", response_class=HTMLResponse)
async def projects_list(request: Request):
    if (r := _require_auth(request)):
        return r
    with Session(engine) as s:
        projects = list(s.exec(select(Project).order_by(Project.number)).all())
        ep_counts_rows = s.exec(
            select(Episode.project_id, func.count(Episode.id)).group_by(Episode.project_id)
        ).all()
        ep_counts = {pid: cnt for pid, cnt in ep_counts_rows}

    rows = [
        {
            "id": p.id,
            "number": p.number,
            "title": p.title,
            "episodes": ep_counts.get(p.id, 0),
            "start_date": p.start_date,
        }
        for p in projects
    ]
    return _render(request, "projects.html", {"title": "Projects", "projects": rows})


# ---------- calendar ----------
@app.get("/calendar", response_class=HTMLResponse)
async def calendar_page(request: Request):
    if (r := _require_auth(request)):
        return r
    with Session(engine) as s:
        projects = list(s.exec(select(Project).order_by(Project.title)).all())
        actors = list(s.exec(select(Actor).order_by(Actor.name)).all())
        rooms = list(s.exec(select(Room).order_by(Room.sort_order, Room.name)).all())
    ctx = {
        "title": "Calendar",
        "projects": [
            {"id": p.id, "title": p.title, "color": p.color or ""}
            for p in projects
        ],
        "actors": [{"id": a.id, "name": a.name} for a in actors],
        "rooms": [{"id": r.id, "name": r.name} for r in rooms],
    }
    return _render(request, "calendar.html", ctx)


# ---------- new project (upload) ----------
@app.post("/projects/delete")
async def projects_delete(request: Request):
    if (r := _require_admin(request)):
        return r
    form = await request.form()
    raw_ids = form.getlist("project_ids")
    ids: list[int] = []
    for v in raw_ids:
        try:
            ids.append(int(v))
        except (TypeError, ValueError):
            continue
    if not ids:
        return RedirectResponse("/projects", status_code=303)

    deleted = 0
    with Session(engine) as s:
        for pid in ids:
            project = s.get(Project, pid)
            if not project:
                continue
            ep_ids = [
                e.id for e in s.exec(select(Episode).where(Episode.project_id == pid)).all()
            ]
            if ep_ids:
                for wc in s.exec(
                    select(WordCount).where(WordCount.episode_id.in_(ep_ids))
                ).all():
                    s.delete(wc)
                for e in s.exec(select(Episode).where(Episode.project_id == pid)).all():
                    s.delete(e)
            for a in s.exec(
                select(Assignment).where(Assignment.project_id == pid)
            ).all():
                s.delete(a)
            for c in s.exec(
                select(Character).where(Character.project_id == pid)
            ).all():
                s.delete(c)
            s.delete(project)
            deleted += 1
        s.commit()

    _flash(request, f"Projects deleted: {deleted}", level="success" if deleted else "error")
    return RedirectResponse("/projects", status_code=303)


def _episodes_from_payload(payload_episodes: list[dict]) -> dict[int, EpisodeData]:
    """JSON from import_parser.js → dict[episode_num, EpisodeData]."""
    out: dict[int, EpisodeData] = {}
    for item in payload_episodes:
        try:
            ep_num = int(item.get("episode_num"))
        except (TypeError, ValueError):
            continue
        raw_rows = item.get("rows") or []
        rows: list[tuple] = []
        for r in raw_rows:
            if not isinstance(r, list):
                continue
            padded = list(r) + [None] * max(0, 8 - len(r))
            rows.append(tuple(padded[:8]))
        out[ep_num] = EpisodeData(
            number=ep_num,
            filename=str(item.get("filename") or f"episode-{ep_num}.xlsx"),
            rows=rows,
            total=None,
            show_title=str(item.get("show_title") or "").strip(),
        )
    return out


@app.post("/api/import-json")
async def import_json(request: Request):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    project_id = body.get("project_id")
    warnings: list[str] = list(body.get("warnings") or [])
    payload_eps = body.get("episodes") or []
    files_total = int(body.get("files_total") or len(payload_eps) + len(warnings))

    episodes = _episodes_from_payload(payload_eps)
    profile = PROFILES["default"]

    if project_id is None:
        if not episodes:
            _flash_import_summary(request, files_total, 0, warnings)
            return JSONResponse({"ok": False, "reason": "no-episodes"}, status_code=400)
        override = str(body.get("project_title") or "").strip()
        if override:
            title = override
        else:
            title = derive_show_title(episodes)
            if not title:
                title = derive_common_name([d.filename for d in episodes.values()], profile)
            if not title:
                title = "Untitled"
        with Session(engine) as s:
            current_max = s.exec(select(func.coalesce(func.max(Project.number), 0))).one()
            project = Project(number=current_max + 1, title=title)
            s.add(project)
            s.commit()
            s.refresh(project)
            _ingest_episodes(s, project.id, episodes, profile, mark_new_unacknowledged=False)
            s.commit()
            new_project_id = project.id
        _flash_import_summary(request, files_total, len(episodes), warnings)
        return JSONResponse({"ok": True, "project_id": new_project_id})

    try:
        pid = int(project_id)
    except (TypeError, ValueError):
        raise HTTPException(400, "Invalid project_id")
    with Session(engine) as s:
        project = s.get(Project, pid)
        if not project:
            raise HTTPException(404, "Project not found")
        if episodes:
            _ingest_episodes(s, pid, episodes, profile)
            s.commit()
    _flash_import_summary(request, files_total, len(episodes), warnings)
    return JSONResponse({"ok": True, "project_id": pid})


@app.get("/projects/new", response_class=HTMLResponse)
async def project_new_form(request: Request):
    if (r := _require_auth(request)):
        return r
    return _render(request, "project_new.html", {"title": "New project"})




def _to_int(value) -> int:
    if value is None or value == "":
        return 0
    try:
        return int(value)
    except (TypeError, ValueError):
        try:
            return int(float(str(value).replace(",", ".").strip()))
        except (TypeError, ValueError):
            return 0


# Production-metadata rows in Netflix word-count sheets that should not be
# treated as characters. Compared case-insensitively against the trimmed
# character-name cell. Keep in sync with import_parser.js and the cleanup
# migration in db._cleanup_junk_characters.
_JUNK_CHARACTER_NAMES: set[str] = {
    "PRINCIPAL PHOTOGRAPHY",
    "GRAPHICS INSERTS",
    "MAIN TITLE",
}


def _ingest_episodes(
    session: Session,
    project_id: int,
    episodes: dict,
    profile,
    mark_new_unacknowledged: bool = True,
) -> None:
    for ep_num, data in episodes.items():
        episode = session.exec(
            select(Episode).where(
                Episode.project_id == project_id, Episode.number == ep_num
            )
        ).first()
        if episode:
            episode.uploaded_filename = data.filename
            episode.imported_at = datetime.now(timezone.utc)
            session.add(episode)
            session.flush()
        else:
            episode = Episode(
                project_id=project_id,
                number=ep_num,
                uploaded_filename=data.filename,
            )
            session.add(episode)
            session.flush()

        xlsx_rows: dict[str, tuple[int, int, int]] = {}
        for r in data.rows:
            name_cell = r[profile.col_character]
            name = str(name_cell).strip() if name_cell else ""
            # Netflix word-count sheets include production-metadata rows
            # (main title card, photography, graphics inserts) that are
            # not characters — drop them before counting.
            if name.upper() in _JUNK_CHARACTER_NAMES:
                continue
            if not name:
                name = "(unnamed)"
            d = _to_int(r[profile.col_dialog])
            t = _to_int(r[profile.col_transcription])
            tot = _to_int(r[profile.col_total])
            if name in xlsx_rows:
                prev = xlsx_rows[name]
                xlsx_rows[name] = (prev[0] + d, prev[1] + t, prev[2] + tot)
            else:
                xlsx_rows[name] = (d, t, tot)

        char_by_name: dict[str, Character] = {}
        for name in xlsx_rows:
            char = session.exec(
                select(Character).where(
                    Character.project_id == project_id, Character.name == name
                )
            ).first()
            if not char:
                char = Character(
                    project_id=project_id,
                    name=name,
                    acknowledged=not mark_new_unacknowledged,
                )
                session.add(char)
                session.flush()
            char_by_name[name] = char

        existing_wcs = list(
            session.exec(select(WordCount).where(WordCount.episode_id == episode.id)).all()
        )
        existing_by_char_id = {wc.character_id: wc for wc in existing_wcs}
        new_char_ids = {c.id for c in char_by_name.values()}

        for name, (d, t, tot) in xlsx_rows.items():
            char = char_by_name[name]
            wc = existing_by_char_id.get(char.id)
            if wc:
                if wc.edited:
                    continue
                wc.dialog_wc = d
                wc.transcription_wc = t
                wc.total_wc = tot
                session.add(wc)
            else:
                session.add(
                    WordCount(
                        episode_id=episode.id,
                        character_id=char.id,
                        dialog_wc=d,
                        transcription_wc=t,
                        total_wc=tot,
                    )
                )

        for wc in existing_wcs:
            if wc.character_id not in new_char_ids and not wc.edited:
                session.delete(wc)

        session.flush()


# ---------- project detail ----------
@app.get("/projects/{project_id}", response_class=HTMLResponse)
async def project_detail(request: Request, project_id: int):
    if (r := _require_auth(request)):
        return r
    with Session(engine) as s:
        project = s.get(Project, project_id)
        if not project:
            raise HTTPException(404, "Project not found")

        episodes = list(
            s.exec(
                select(Episode).where(Episode.project_id == project_id).order_by(Episode.number)
            ).all()
        )
        ep_numbers = [e.number for e in episodes]

        characters = list(
            s.exec(select(Character).where(Character.project_id == project_id)).all()
        )

        wc_rows = s.exec(
            select(WordCount, Episode.number)
            .join(Episode, Episode.id == WordCount.episode_id)
            .where(Episode.project_id == project_id)
        ).all()

        assignment_rows = s.exec(
            select(Assignment, Actor.name)
            .join(Actor, Actor.id == Assignment.actor_id)
            .where(Assignment.project_id == project_id)
        ).all()
        # character_id → list of (actor_id, actor_name) in input order.
        actors_by_char: dict[int, list[tuple[int, str]]] = {}
        for a, actor_name in assignment_rows:
            actors_by_char.setdefault(a.character_id, []).append((a.actor_id, actor_name))

        actors_all = list(s.exec(select(Actor).order_by(Actor.name)).all())

        edited_count = s.exec(
            select(func.count(WordCount.id))
            .join(Episode, Episode.id == WordCount.episode_id)
            .where(Episode.project_id == project_id, WordCount.edited == True)  # noqa: E712
        ).one()

    wc_by_char_ep: dict[int, dict[int, tuple[int, int, int]]] = {}
    for wc, ep_num in wc_rows:
        wc_by_char_ep.setdefault(wc.character_id, {})[ep_num] = (
            wc.dialog_wc,
            wc.transcription_wc,
            wc.total_wc,
        )

    rows = []
    ep_totals: dict[int, dict[str, int]] = {n: {"dialog": 0, "transcription": 0} for n in ep_numbers}
    grand = {"dialog": 0, "transcription": 0}
    for c in sorted(characters, key=lambda x: x.name.upper()):
        actors_for_char = actors_by_char.get(c.id, [])
        per_ep = []
        totals = {"dialog": 0, "transcription": 0}
        episodes_count = {"dialog": 0, "transcription": 0}
        for n in ep_numbers:
            d, t, _tot = wc_by_char_ep.get(c.id, {}).get(n, (0, 0, 0))
            per_ep.append({"dialog": d, "transcription": t})
            totals["dialog"] += d
            totals["transcription"] += t
            ep_totals[n]["dialog"] += d
            ep_totals[n]["transcription"] += t
            if d > 0:
                episodes_count["dialog"] += 1
            if t > 0:
                episodes_count["transcription"] += 1
        grand["dialog"] += totals["dialog"]
        grand["transcription"] += totals["transcription"]
        actor_ids = [aid for aid, _ in actors_for_char]
        actor_names = [n for _, n in actors_for_char]
        rows.append(
            {
                "character_id": c.id,
                "character": c.name,
                "acknowledged": c.acknowledged,
                # Joined display: comma-separated. Empty string when no actors.
                "actor_name": ", ".join(actor_names),
                # Structured fields for the JS layer (tooltip, edit roundtrip).
                "actor_ids": actor_ids,
                "actor_names": actor_names,
                "per_ep": per_ep,
                "totals": totals,
                "episodes_count": episodes_count,
            }
        )

    footer = {
        "per_ep": [ep_totals[n] for n in ep_numbers],
        "totals": grand,
    }

    # Bookkeeper report: same computation as /projects/{id}/report xlsx,
    # but reusing the already-loaded rows[] so the third tab on the project
    # page doesn't hit the database again.
    actor_totals: dict[str, int] = {}
    for r in rows:
        names = r.get("actor_names") or []
        # Skip массовка (multi-actor characters) — payout rule pending.
        # Single-actor rows attribute their full transcription to that actor.
        if len(names) != 1:
            continue
        actor_totals[names[0]] = actor_totals.get(names[0], 0) + r["totals"]["transcription"]
    report_rows = [
        {"actor": name, "words": words}
        for name, words in sorted(actor_totals.items(), key=lambda x: x[0].lower())
    ]
    report_total = sum(x["words"] for x in report_rows)

    return _render(
        request,
        "project.html",
        {
            "title": project.title,
            "project": project,
            "episodes": ep_numbers,
            "rows": rows,
            "footer": footer,
            "actors": [a.name for a in actors_all],
            "edited_count": edited_count,
            "report_rows": report_rows,
            "report_total": report_total,
        },
    )


# ---------- export xlsx ----------
@app.get("/projects/{project_id}/export")
async def project_export(request: Request, project_id: int):
    if (r := _require_auth(request)):
        return r
    with Session(engine) as s:
        project = s.get(Project, project_id)
        if not project:
            raise HTTPException(404, "Project not found")
        episodes = list(
            s.exec(
                select(Episode).where(Episode.project_id == project_id).order_by(Episode.number)
            ).all()
        )
        ep_numbers = [e.number for e in episodes]
        characters = list(
            s.exec(select(Character).where(Character.project_id == project_id)).all()
        )
        wc_rows = s.exec(
            select(WordCount, Episode.number)
            .join(Episode, Episode.id == WordCount.episode_id)
            .where(Episode.project_id == project_id)
        ).all()
        actor_rows = s.exec(
            select(Assignment, Actor.name, Character.name)
            .join(Actor, Actor.id == Assignment.actor_id)
            .join(Character, Character.id == Assignment.character_id)
            .where(Assignment.project_id == project_id)
        ).all()

    char_names = sorted((c.name for c in characters), key=str.upper)
    # Multiple actors per character collapse to a comma-joined string for
    # the xlsx export (one row per character, like the original Netflix sheet).
    _by_char: dict[str, list[str]] = {}
    for _, actor_name, char_name in actor_rows:
        _by_char.setdefault(char_name, []).append(actor_name)
    actor_by_char = {k: ", ".join(v) for k, v in _by_char.items()}
    char_name_by_id = {c.id: c.name for c in characters}

    transcription: dict[str, dict[int, int]] = {}
    dialogue: dict[str, dict[int, int]] = {}
    for wc, ep_num in wc_rows:
        name = char_name_by_id.get(wc.character_id)
        if not name:
            continue
        if wc.transcription_wc:
            transcription.setdefault(name, {})[ep_num] = wc.transcription_wc
        if wc.dialog_wc:
            dialogue.setdefault(name, {})[ep_num] = wc.dialog_wc

    blob = build_project_xlsx(
        show_title=project.title,
        ep_numbers=ep_numbers,
        characters=char_names,
        actor_by_char=actor_by_char,
        transcription=transcription,
        dialogue=dialogue,
    )
    safe_title = (project.title or f"project-{project_id}").strip() or f"project-{project_id}"
    filename = f"{safe_title} - Word Count Summary.xlsx"
    return Response(
        content=blob,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _attachment_header(filename)},
    )


# ---------- actor report ----------
@app.get("/projects/{project_id}/report")
async def project_report(request: Request, project_id: int):
    if (r := _require_auth(request)):
        return r
    with Session(engine) as s:
        project = s.get(Project, project_id)
        if not project:
            raise HTTPException(404, "Project not found")
        episode_ids = [
            e.id for e in s.exec(
                select(Episode).where(Episode.project_id == project_id)
            ).all()
        ]
        wc_by_char: dict[int, int] = {}
        if episode_ids:
            for wc in s.exec(
                select(WordCount).where(WordCount.episode_id.in_(episode_ids))
            ).all():
                wc_by_char[wc.character_id] = wc_by_char.get(wc.character_id, 0) + wc.transcription_wc
        # Multi-actor characters (массовка) are skipped — payout rule TBD.
        # Single-actor characters attribute their full transcription to
        # that actor.
        actors_per_char: dict[int, list[int]] = {}
        for a in s.exec(
            select(Assignment).where(Assignment.project_id == project_id)
        ).all():
            actors_per_char.setdefault(a.character_id, []).append(a.actor_id)
        actor_total: dict[int, int] = {}
        for cid, aids in actors_per_char.items():
            if len(aids) != 1:
                continue
            words = wc_by_char.get(cid, 0)
            actor_total[aids[0]] = actor_total.get(aids[0], 0) + words
        if not actor_total:
            rows = []
        else:
            actor_names = {
                a.id: a.name for a in s.exec(
                    select(Actor).where(Actor.id.in_(list(actor_total.keys())))
                ).all()
            }
            rows = sorted(
                ((actor_names[aid], total) for aid, total in actor_total.items() if aid in actor_names),
                key=lambda r: r[0].lower(),
            )

    blob = build_actor_report_xlsx(project.title, rows)
    safe_title = (project.title or f"project-{project_id}").strip() or f"project-{project_id}"
    filename = f"{safe_title} - Report.xlsx"
    return Response(
        content=blob,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _attachment_header(filename)},
    )




# ---------- actor assignment API ----------
def _parse_actor_names(raw: str) -> list[str]:
    """Split a comma-separated cell value into trimmed, deduped names.
    Case-insensitive dedup but preserves the first-seen casing."""
    seen: set[str] = set()
    out: list[str] = []
    for piece in (raw or "").split(","):
        n = piece.strip()
        if not n:
            continue
        key = n.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(n)
    return out


@app.post("/api/projects/{project_id}/characters/{character_id}/actor")
async def set_actor(request: Request, project_id: int, character_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    # Accept either {"name": "A, B"} (legacy single field, now also used
    # for multi) or {"names": [...]} for explicit list.
    if isinstance(body.get("names"), list):
        wanted = [str(x).strip() for x in body["names"] if str(x).strip()]
        # dedup with same rules as the parser
        wanted = _parse_actor_names(",".join(wanted))
    else:
        wanted = _parse_actor_names(body.get("name") or "")

    with Session(engine) as s:
        char = s.get(Character, character_id)
        if not char or char.project_id != project_id:
            raise HTTPException(404, "Character not found")

        # Multi-actor cast is only meaningful for WALLA / crowd characters.
        # On a regular character we silently keep just the first name —
        # protects against stray commas in copy-paste from creating
        # phantom assignments and matches the UI rules on the front-end.
        is_walla = "WALLA" in (char.name or "").upper()
        if not is_walla and len(wanted) > 1:
            wanted = wanted[:1]

        existing = list(s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id,
                Assignment.character_id == character_id,
            )
        ).all())

        # Empty list — drop all assignments for this character.
        if not wanted:
            for a in existing:
                s.delete(a)
            s.commit()
            return JSONResponse({"actor_ids": [], "actor_names": [], "actor_name": ""})

        # Resolve names to actor rows, auto-creating missing ones.
        all_actors = list(s.exec(select(Actor)).all())
        by_lower = {a.name.lower(): a for a in all_actors}
        resolved: list[Actor] = []
        for n in wanted:
            actor = by_lower.get(n.lower())
            if not actor:
                actor = Actor(name=n)
                s.add(actor)
                s.flush()
                by_lower[n.lower()] = actor
            resolved.append(actor)

        # Replace assignments: drop all old, flush so the UNIQUE
        # (character_id, actor_id) constraint isn't tripped if the same
        # actor is being re-added, then re-insert.
        for a in existing:
            s.delete(a)
        s.flush()
        seen_ids: set[int] = set()
        for actor in resolved:
            if actor.id in seen_ids:
                continue
            seen_ids.add(actor.id)
            s.add(
                Assignment(
                    project_id=project_id,
                    character_id=character_id,
                    actor_id=actor.id,
                )
            )
        s.commit()

        ids = [a.id for a in resolved]
        names = [a.name for a in resolved]
        return JSONResponse({
            "actor_ids": ids,
            "actor_names": names,
            # Backwards-compat field for any caller still reading actor_name.
            "actor_name": ", ".join(names),
        })


@app.post("/api/projects/{project_id}/characters/{character_id}/acknowledge")
async def acknowledge_character(request: Request, project_id: int, character_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    with Session(engine) as s:
        char = s.get(Character, character_id)
        if not char or char.project_id != project_id:
            raise HTTPException(404, "Character not found")
        char.acknowledged = True
        s.add(char)
        s.commit()
    return JSONResponse({"ok": True})


@app.post("/api/projects/{project_id}/characters/{character_id}/merge")
async def merge_character(request: Request, project_id: int, character_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    try:
        target_id = int(body.get("target_id") or 0)
    except (TypeError, ValueError):
        raise HTTPException(400, "Invalid target_id")
    if not target_id or target_id == character_id:
        raise HTTPException(400, "Invalid target_id")

    with Session(engine) as s:
        source = s.get(Character, character_id)
        target = s.get(Character, target_id)
        if not source or source.project_id != project_id:
            raise HTTPException(404, "Character not found")
        if not target or target.project_id != project_id:
            raise HTTPException(404, "Target character not found")

        src_wcs = list(s.exec(select(WordCount).where(WordCount.character_id == source.id)).all())
        tgt_wcs_by_ep = {
            wc.episode_id: wc
            for wc in s.exec(select(WordCount).where(WordCount.character_id == target.id)).all()
        }

        for swc in src_wcs:
            twc = tgt_wcs_by_ep.get(swc.episode_id)
            if twc:
                if swc.edited and not twc.edited:
                    twc.dialog_wc = swc.dialog_wc
                    twc.transcription_wc = swc.transcription_wc
                    twc.total_wc = swc.total_wc
                    twc.edited = True
                    s.add(twc)
                s.delete(swc)
            else:
                swc.character_id = target.id
                s.add(swc)

        src_assigns = list(s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id, Assignment.character_id == source.id
            )
        ).all())
        tgt_actor_ids = {
            a.actor_id for a in s.exec(
                select(Assignment).where(
                    Assignment.project_id == project_id,
                    Assignment.character_id == target.id,
                )
            ).all()
        }
        # Move source assignments onto target unless that actor is already
        # cast on target (UNIQUE(character_id, actor_id) would fail).
        for sa in src_assigns:
            if sa.actor_id in tgt_actor_ids:
                s.delete(sa)
            else:
                sa.character_id = target.id
                tgt_actor_ids.add(sa.actor_id)
                s.add(sa)

        target.acknowledged = True
        s.add(target)
        s.delete(source)
        s.commit()

    return JSONResponse({"ok": True, "target_id": target_id})


@app.post("/api/projects/{project_id}/characters/{character_id}/delete")
async def delete_character(request: Request, project_id: int, character_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    with Session(engine) as s:
        char = s.get(Character, character_id)
        if not char or char.project_id != project_id:
            raise HTTPException(404, "Character not found")
        for wc in s.exec(select(WordCount).where(WordCount.character_id == character_id)).all():
            s.delete(wc)
        for a in s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id, Assignment.character_id == character_id
            )
        ).all():
            s.delete(a)
        s.delete(char)
        s.commit()
    return JSONResponse({"ok": True})


@app.post("/api/projects/{project_id}/wordcount")
async def set_wordcount(request: Request, project_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    character_id = int(body.get("character_id") or 0)
    episode_number = int(body.get("episode_number") or 0)
    metric = body.get("metric")
    if metric not in ("transcription", "dialog"):
        raise HTTPException(400, "Unknown metric")
    try:
        value = max(0, int(body.get("value") or 0))
    except (TypeError, ValueError):
        value = 0

    with Session(engine) as s:
        char = s.get(Character, character_id)
        if not char or char.project_id != project_id:
            raise HTTPException(404, "Character not found")
        episode = s.exec(
            select(Episode).where(
                Episode.project_id == project_id, Episode.number == episode_number
            )
        ).first()
        if not episode:
            raise HTTPException(404, "Episode not found")
        wc = s.exec(
            select(WordCount).where(
                WordCount.episode_id == episode.id, WordCount.character_id == character_id
            )
        ).first()
        if not wc:
            wc = WordCount(episode_id=episode.id, character_id=character_id)
        if metric == "transcription":
            wc.transcription_wc = value
        else:
            wc.dialog_wc = value
        wc.edited = True
        s.add(wc)
        s.commit()
    return JSONResponse({"value": value, "edited": True})


# ---------- admin ----------
def _admin_context(session: Session, tab: str, error: str = "") -> dict:
    users = list(session.exec(select(User).order_by(User.id)).all())
    actors = list(session.exec(select(Actor).order_by(func.lower(Actor.name))).all())
    actor_usage = dict(
        session.exec(
            select(Assignment.actor_id, func.count(Assignment.id)).group_by(Assignment.actor_id)
        ).all()
    )
    return {
        "title": "Admin",
        "tab": tab,
        "users": users,
        "actors": [
            {"id": a.id, "name": a.name, "tg": a.tg or "", "used": actor_usage.get(a.id, 0)}
            for a in actors
        ],
        "error": error,
    }


@app.get("/admin", response_class=HTMLResponse)
async def admin_home(request: Request, tab: str = "actors"):
    if (r := _require_admin(request)):
        return r
    tab = tab if tab in ("users", "actors") else "actors"
    with Session(engine) as s:
        ctx = _admin_context(s, tab)
    return _render(request, "admin.html", ctx)


@app.post("/admin/users")
async def admin_user_create(
    request: Request,
    email: str = Form(...),
    password: str = Form(...),
    role: str = Form("assistant_director"),
):
    if (r := _require_admin(request)):
        return r
    email = email.strip()
    password = password.strip()
    if role not in ("admin", "assistant_director"):
        role = "assistant_director"
    if not email or not password:
        return RedirectResponse("/admin?tab=users", status_code=303)
    with Session(engine) as s:
        exists = s.exec(select(User).where(User.email == email)).first()
        if exists:
            ctx = _admin_context(s, "users", error=f"User \"{email}\" already exists")
            return _render(request, "admin.html", ctx)
        s.add(User(email=email, password_hash=hash_password(password), role=role))
        s.commit()
    return RedirectResponse("/admin?tab=users", status_code=303)


@app.post("/admin/users/{user_id}/delete")
async def admin_user_delete(request: Request, user_id: int):
    if (r := _require_admin(request)):
        return r
    current_user_id = request.session.get("user_id")
    with Session(engine) as s:
        user = s.get(User, user_id)
        if not user:
            return RedirectResponse("/admin?tab=users", status_code=303)
        if user.id == current_user_id:
            ctx = _admin_context(s, "users", error="Cannot delete yourself")
            return _render(request, "admin.html", ctx)
        if user.role == "admin":
            admin_count = s.exec(
                select(func.count(User.id)).where(User.role == "admin")
            ).one()
            if admin_count <= 1:
                ctx = _admin_context(s, "users", error="Cannot delete the last admin")
                return _render(request, "admin.html", ctx)
        s.delete(user)
        s.commit()
    return RedirectResponse("/admin?tab=users", status_code=303)


@app.post("/admin/actors")
async def admin_actor_create(request: Request, name: str = Form(...)):
    if (r := _require_admin(request)):
        return r
    name = name.strip()
    if not name:
        return RedirectResponse("/admin?tab=actors", status_code=303)
    with Session(engine) as s:
        name_lower = name.lower()
        existing = next(
            (a for a in s.exec(select(Actor)).all() if a.name.lower() == name_lower),
            None,
        )
        if existing:
            ctx = _admin_context(s, "actors", error=f"Actor \"{name}\" already exists")
            return _render(request, "admin.html", ctx)
        s.add(Actor(name=name))
        s.commit()
    return RedirectResponse("/admin?tab=actors", status_code=303)


@app.post("/admin/actors/{actor_id}/rename")
async def admin_actor_rename(request: Request, actor_id: int, name: str = Form(...)):
    if (r := _require_admin(request)):
        return r
    # AJAX callers (admin.html autosave-on-blur) send Accept: application/json
    # so we can report clashes without a full page re-render.
    wants_json = "application/json" in request.headers.get("accept", "").lower()
    name = name.strip()
    if not name:
        if wants_json:
            return JSONResponse({"ok": False, "error": "name is required"}, status_code=400)
        return RedirectResponse("/admin?tab=actors", status_code=303)
    with Session(engine) as s:
        actor = s.get(Actor, actor_id)
        if not actor:
            if wants_json:
                return JSONResponse({"ok": False, "error": "not found"}, status_code=404)
            return RedirectResponse("/admin?tab=actors", status_code=303)
        name_lower = name.lower()
        clash = next(
            (
                a for a in s.exec(select(Actor).where(Actor.id != actor_id)).all()
                if a.name.lower() == name_lower
            ),
            None,
        )
        if clash:
            msg = f'Name "{name}" already taken'
            if wants_json:
                return JSONResponse({"ok": False, "error": msg}, status_code=409)
            ctx = _admin_context(s, "actors", error=msg)
            return _render(request, "admin.html", ctx)
        actor.name = name
        s.add(actor)
        s.commit()
    if wants_json:
        return JSONResponse({"ok": True, "name": name})
    return RedirectResponse("/admin?tab=actors", status_code=303)


@app.post("/admin/actors/{actor_id}/tg")
async def admin_actor_set_tg(request: Request, actor_id: int, tg: str = Form("")):
    if (r := _require_admin(request)):
        return r
    # Normalize: strip whitespace and a leading @ so the column holds bare
    # handles (some Telegram APIs want the @, others don't — we store the
    # canonical shape and let the notifier glue @ back on).
    tg = tg.strip().lstrip("@").strip()
    with Session(engine) as s:
        actor = s.get(Actor, actor_id)
        if not actor:
            return JSONResponse({"ok": False, "error": "not found"}, status_code=404)
        actor.tg = tg
        s.add(actor)
        s.commit()
    return JSONResponse({"ok": True, "tg": tg})


@app.post("/admin/actors/{actor_id}/delete")
async def admin_actor_delete(request: Request, actor_id: int):
    if (r := _require_admin(request)):
        return r
    with Session(engine) as s:
        actor = s.get(Actor, actor_id)
        if not actor:
            return RedirectResponse("/admin?tab=actors", status_code=303)
        for a in s.exec(select(Assignment).where(Assignment.actor_id == actor_id)).all():
            s.delete(a)
        s.delete(actor)
        s.commit()
    return RedirectResponse("/admin?tab=actors", status_code=303)


# ---------- calendar API: rooms + sessions ----------
def _parse_iso_dt(s: str) -> Optional[datetime]:
    """Parse an ISO-8601 datetime, accept the JS 'Z' suffix as +00:00."""
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except (TypeError, ValueError):
        return None


def _session_to_dict(rs: RecordingSession, actor_ids: list[int]) -> dict:
    return {
        "id": rs.id,
        "project_id": rs.project_id,
        "actor_ids": list(actor_ids),
        "room_id": rs.room_id,
        "starts_at": rs.starts_at.isoformat(),
        "ends_at": rs.ends_at.isoformat(),
        "status": rs.status,
        "target_words": rs.target_words,
        "episode_numbers": rs.episode_numbers,
        "notes": rs.notes,
    }


def _load_actor_ids(s: Session, session_id: int) -> list[int]:
    return [
        sa.actor_id for sa in s.exec(
            select(SessionActor)
            .where(SessionActor.session_id == session_id)
            .order_by(SessionActor.id)
        ).all()
    ]


def _replace_session_actors(s: Session, session_id: int, actor_ids: list[int]) -> None:
    """Drop existing actor links for the session and insert the given list
    in order. Caller is responsible for validating each actor exists."""
    for sa in s.exec(
        select(SessionActor).where(SessionActor.session_id == session_id)
    ).all():
        s.delete(sa)
    # Flush so the deletes hit the DB before we insert the (potentially
    # overlapping) new rows, otherwise the UNIQUE(session_id, actor_id)
    # constraint trips when re-adding the same actor.
    s.flush()
    seen: set[int] = set()
    for aid in actor_ids:
        if aid in seen:
            continue
        seen.add(aid)
        s.add(SessionActor(session_id=session_id, actor_id=aid))


@app.get("/api/rooms")
async def api_list_rooms(request: Request):
    if not is_authenticated(request):
        raise HTTPException(401)
    with Session(engine) as s:
        rooms = list(
            s.exec(select(Room).order_by(Room.sort_order, Room.name)).all()
        )
    return JSONResponse(
        [{"id": r.id, "name": r.name, "color": r.color} for r in rooms]
    )


@app.get("/api/sessions")
async def api_list_sessions(
    request: Request,
    from_: Optional[str] = None,
    to: Optional[str] = None,
    project_id: Optional[int] = None,
    actor_id: Optional[int] = None,
    room_id: Optional[int] = None,
):
    """List sessions overlapping [from, to]. FullCalendar passes the
    visible window via its `events` callback as `start` / `end` query
    params; we accept `from` (since `from` is reserved in Python) under
    the `from_` parameter — clients should send `?from=...&to=...`.

    actor_id filter joins through SessionActor so a session is included
    when ANY of its booked actors matches."""
    if not is_authenticated(request):
        raise HTTPException(401)
    if from_ is None:
        from_ = request.query_params.get("from")
    start = _parse_iso_dt(from_) if from_ else None
    end = _parse_iso_dt(to) if to else None
    with Session(engine) as s:
        stmt = select(RecordingSession)
        if start is not None:
            stmt = stmt.where(RecordingSession.ends_at >= start)
        if end is not None:
            stmt = stmt.where(RecordingSession.starts_at <= end)
        if project_id is not None:
            stmt = stmt.where(RecordingSession.project_id == project_id)
        if room_id is not None:
            stmt = stmt.where(RecordingSession.room_id == room_id)
        if actor_id is not None:
            matching = s.exec(
                select(SessionActor.session_id).where(SessionActor.actor_id == actor_id)
            ).all()
            stmt = stmt.where(RecordingSession.id.in_(list(matching)))
        rows = list(s.exec(stmt.order_by(RecordingSession.starts_at)).all())
        actors_by_session: dict[int, list[int]] = {}
        if rows:
            for sa in s.exec(
                select(SessionActor)
                .where(SessionActor.session_id.in_([rs.id for rs in rows]))
                .order_by(SessionActor.id)
            ).all():
                actors_by_session.setdefault(sa.session_id, []).append(sa.actor_id)
    return JSONResponse([
        _session_to_dict(rs, actors_by_session.get(rs.id, [])) for rs in rows
    ])


def _validate_refs(
    s: Session, *, project_id: int, actor_ids: list[int], room_id: Optional[int]
):
    if not s.get(Project, project_id):
        raise HTTPException(400, "project not found")
    if not actor_ids:
        raise HTTPException(400, "at least one actor is required")
    for aid in actor_ids:
        if not s.get(Actor, aid):
            raise HTTPException(400, f"actor {aid} not found")
    if room_id is not None and not s.get(Room, room_id):
        raise HTTPException(400, "room not found")


def _coerce_actor_ids(raw) -> list[int]:
    """Accept either a list of ints/strings or None; return list of int.
    Raises 400 on garbage input."""
    if raw is None:
        return []
    if not isinstance(raw, (list, tuple)):
        raise HTTPException(400, "actor_ids must be an array")
    out: list[int] = []
    for v in raw:
        try:
            out.append(int(v))
        except (TypeError, ValueError):
            raise HTTPException(400, f"invalid actor id: {v!r}")
    return out


@app.post("/api/sessions")
async def api_create_session(request: Request):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    starts = _parse_iso_dt(body.get("starts_at"))
    ends = _parse_iso_dt(body.get("ends_at"))
    if not starts or not ends or starts >= ends:
        raise HTTPException(400, "invalid time range")
    try:
        project_id = int(body["project_id"])
    except (KeyError, TypeError, ValueError):
        raise HTTPException(400, "project_id is required")
    actor_ids = _coerce_actor_ids(body.get("actor_ids"))
    room_id_raw = body.get("room_id")
    room_id = int(room_id_raw) if room_id_raw not in (None, "") else None
    status = body.get("status") or "planned"
    if status not in SESSION_STATUSES:
        raise HTTPException(400, f"invalid status (allowed: {', '.join(SESSION_STATUSES)})")
    target_words = body.get("target_words")
    if target_words in ("", None):
        target_words = None
    else:
        try:
            target_words = int(target_words)
        except (TypeError, ValueError):
            raise HTTPException(400, "target_words must be an integer")
    with Session(engine) as s:
        _validate_refs(s, project_id=project_id, actor_ids=actor_ids, room_id=room_id)
        rs = RecordingSession(
            project_id=project_id,
            room_id=room_id,
            starts_at=starts,
            ends_at=ends,
            status=status,
            target_words=target_words,
            episode_numbers=str(body.get("episode_numbers") or "").strip(),
            notes=str(body.get("notes") or ""),
        )
        s.add(rs)
        s.flush()
        _replace_session_actors(s, rs.id, actor_ids)
        s.commit()
        s.refresh(rs)
        return JSONResponse(_session_to_dict(rs, _load_actor_ids(s, rs.id)))


@app.patch("/api/sessions/{session_id}")
async def api_update_session(request: Request, session_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    with Session(engine) as s:
        rs = s.get(RecordingSession, session_id)
        if not rs:
            raise HTTPException(404, "session not found")
        if "starts_at" in body:
            v = _parse_iso_dt(body["starts_at"])
            if not v:
                raise HTTPException(400, "invalid starts_at")
            rs.starts_at = v
        if "ends_at" in body:
            v = _parse_iso_dt(body["ends_at"])
            if not v:
                raise HTTPException(400, "invalid ends_at")
            rs.ends_at = v
        if rs.starts_at >= rs.ends_at:
            raise HTTPException(400, "starts_at must precede ends_at")
        if "project_id" in body:
            rs.project_id = int(body["project_id"])
        if "room_id" in body:
            v = body["room_id"]
            rs.room_id = int(v) if v not in (None, "") else None
        if "status" in body:
            if body["status"] not in SESSION_STATUSES:
                raise HTTPException(400, "invalid status")
            rs.status = body["status"]
        if "target_words" in body:
            v = body["target_words"]
            if v in ("", None):
                rs.target_words = None
            else:
                rs.target_words = int(v)
        if "episode_numbers" in body:
            rs.episode_numbers = str(body["episode_numbers"]).strip()
        if "notes" in body:
            rs.notes = str(body["notes"])
        actor_ids = _load_actor_ids(s, rs.id)
        if "actor_ids" in body:
            actor_ids = _coerce_actor_ids(body["actor_ids"])
        _validate_refs(
            s,
            project_id=rs.project_id,
            actor_ids=actor_ids,
            room_id=rs.room_id,
        )
        if "actor_ids" in body:
            _replace_session_actors(s, rs.id, actor_ids)
        rs.updated_at = datetime.now(timezone.utc)
        s.add(rs)
        s.commit()
        s.refresh(rs)
        return JSONResponse(_session_to_dict(rs, _load_actor_ids(s, rs.id)))


@app.delete("/api/sessions/{session_id}")
async def api_delete_session(request: Request, session_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    with Session(engine) as s:
        rs = s.get(RecordingSession, session_id)
        if not rs:
            raise HTTPException(404, "session not found")
        for sa in s.exec(
            select(SessionActor).where(SessionActor.session_id == session_id)
        ).all():
            s.delete(sa)
        s.delete(rs)
        s.commit()
    return JSONResponse({"ok": True})


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
        actor_by_char: dict[int, tuple[int, str]] = {
            a.character_id: (a.actor_id, actor_name) for a, actor_name in assignment_rows
        }

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
        actor = actor_by_char.get(c.id)
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
        rows.append(
            {
                "character_id": c.id,
                "character": c.name,
                "acknowledged": c.acknowledged,
                "actor_id": actor[0] if actor else None,
                "actor_name": actor[1] if actor else "",
                "per_ep": per_ep,
                "totals": totals,
                "episodes_count": episodes_count,
            }
        )

    footer = {
        "per_ep": [ep_totals[n] for n in ep_numbers],
        "totals": grand,
    }

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
    actor_by_char = {char_name: actor_name for _, actor_name, char_name in actor_rows}
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
        actor_total: dict[int, int] = {}
        for a in s.exec(
            select(Assignment).where(Assignment.project_id == project_id)
        ).all():
            actor_total[a.actor_id] = actor_total.get(a.actor_id, 0) + wc_by_char.get(a.character_id, 0)
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
                key=lambda r: (-r[1], r[0].lower()),
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
@app.post("/api/projects/{project_id}/characters/{character_id}/actor")
async def set_actor(request: Request, project_id: int, character_id: int):
    if not is_authenticated(request):
        raise HTTPException(401)
    body = await request.json()
    name = (body.get("name") or "").strip()

    with Session(engine) as s:
        char = s.get(Character, character_id)
        if not char or char.project_id != project_id:
            raise HTTPException(404, "Character not found")

        existing = s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id,
                Assignment.character_id == character_id,
            )
        ).first()

        if not name:
            if existing:
                s.delete(existing)
                s.commit()
            return JSONResponse({"actor_id": None, "actor_name": ""})

        name_lower = name.lower()
        actor = next(
            (a for a in s.exec(select(Actor)).all() if a.name.lower() == name_lower),
            None,
        )
        if not actor:
            actor = Actor(name=name)
            s.add(actor)
            s.flush()

        if existing:
            existing.actor_id = actor.id
            s.add(existing)
        else:
            s.add(
                Assignment(
                    project_id=project_id,
                    character_id=character_id,
                    actor_id=actor.id,
                )
            )
        s.commit()
        return JSONResponse({"actor_id": actor.id, "actor_name": actor.name})


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

        src_assign = s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id, Assignment.character_id == source.id
            )
        ).first()
        tgt_assign = s.exec(
            select(Assignment).where(
                Assignment.project_id == project_id, Assignment.character_id == target.id
            )
        ).first()
        if src_assign:
            if not tgt_assign:
                src_assign.character_id = target.id
                s.add(src_assign)
            else:
                s.delete(src_assign)

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
        "actors": [{"id": a.id, "name": a.name, "used": actor_usage.get(a.id, 0)} for a in actors],
        "error": error,
    }


@app.get("/admin", response_class=HTMLResponse)
async def admin_home(request: Request, tab: str = "users"):
    if (r := _require_admin(request)):
        return r
    tab = tab if tab in ("users", "actors") else "users"
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
    name = name.strip()
    if not name:
        return RedirectResponse("/admin?tab=actors", status_code=303)
    with Session(engine) as s:
        actor = s.get(Actor, actor_id)
        if not actor:
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
            ctx = _admin_context(s, "actors", error=f"Name \"{name}\" already taken")
            return _render(request, "admin.html", ctx)
        actor.name = name
        s.add(actor)
        s.commit()
    return RedirectResponse("/admin?tab=actors", status_code=303)


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

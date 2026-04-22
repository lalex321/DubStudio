# DubStudio

Веб-приложение для студии дубляжа: проекты-фильмы, версии wordcount-файлов,
назначение актёров к персонажам, учёт объёма озвученных слов.

Развивает идеи [Pivot](https://github.com/lalex321/pivot) — Pivot остаётся как
простой разовый конвертер xlsx, DubStudio даёт stateful-workflow.

## Стек

- **Backend**: FastAPI + SQLModel + SQLite (Postgres на Render)
- **Frontend**: Jinja2 + vanilla JS, тёмная тема
- **Без React** — намеренно.

## Локальный запуск

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

uvicorn app:app --reload --port 8001
```

Открой `http://localhost:8001`.

Порт `8001` выбран, чтобы не мешать Pivot-у (он на `8000`).

## Статус

Скелет. MVP-1:

- [ ] Один пользователь-админ (логин/пароль из env)
- [ ] CRUD проектов (фильмов)
- [ ] Загрузка xlsx серий в проект → парсинг в БД
- [ ] Просмотр проекта: грид персонаж × серия + колонка Actor (editable)
- [ ] Расчёт зарплат — позже
- [ ] Роли / многопользовательский режим — позже

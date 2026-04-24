"""
Единый источник путей для DubStudio.

- `BUNDLE_DIR` — откуда читать read-only ресурсы (templates/, static/).
  В PyInstaller-сборке это `sys._MEIPASS`, иначе — корень репо.
- `DATA_DIR` — куда писать mutable данные (SQLite-файл, .session_secret).
  В PyInstaller-сборке — platform-specific app-data dir (создаётся при
  первом запуске). В обычном режиме — корень репо (как раньше).

В dev-режиме и под Render всё работает как и было: код выше просто
возвращает `Path(__file__).parent`, env-переменные имеют приоритет.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _bundle_dir() -> Path:
    if _is_frozen():
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).parent


def _data_dir() -> Path:
    if not _is_frozen():
        return Path(__file__).parent

    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home()) / "DubStudio"
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / "DubStudio"
    else:
        base = Path.home() / ".local" / "share" / "DubStudio"
    base.mkdir(parents=True, exist_ok=True)
    return base


BUNDLE_DIR: Path = _bundle_dir()
DATA_DIR: Path = _data_dir()

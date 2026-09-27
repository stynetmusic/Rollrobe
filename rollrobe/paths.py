"""Пути к программе, базе и файлам прайса."""

from __future__ import annotations

import sys
from pathlib import Path


def app_dir() -> Path:
    """Папка программы: рядом с main.py или с exe, не временный каталог сборки."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_dirs() -> list[Path]:
    """Где искать xlsx: папка программы и, при сборке, распакованные данные."""
    folders = [app_dir()]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled = Path(meipass)
        if bundled not in folders:
            folders.append(bundled)
    return folders


def find_data_file(name: str) -> Path | None:
    for folder in resource_dirs():
        candidate = folder / name
        if candidate.is_file():
            return candidate
    return None


def database_path() -> Path:
    return app_dir() / "data" / "rollrobe.sqlite"

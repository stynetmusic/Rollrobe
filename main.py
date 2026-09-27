#!/usr/bin/env python3
"""Запуск Rollrobe."""

import sys


def main() -> None:
    try:
        from rollrobe.app import main as run_app
    except ModuleNotFoundError as exc:
        missing = exc.name or ""
        if missing.split(".", 1)[0] in {"customtkinter", "openpyxl", "rollrobe"}:
            print(
                "Не хватает зависимости. Установите: python3 -m pip install -r requirements.txt",
                file=sys.stderr,
            )
            print(exc, file=sys.stderr)
            raise SystemExit(1) from exc
        raise
    run_app()


if __name__ == "__main__":
    main()

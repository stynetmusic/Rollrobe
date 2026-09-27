"""SQLite: прайсы, позиции, сметы. Файл data/rollrobe.sqlite рядом с программой."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from rollrobe.money import line_total, q2

UNIT_CODES = (0, 2, 3, 4)
DEFAULT_UNIT_LABELS = {0: "ед. 0", 2: "ед. 2", 3: "ед. 3", 4: "ед. 4"}

SCHEMA = """
CREATE TABLE IF NOT EXISTS price_lists (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    filename TEXT NOT NULL UNIQUE,
    is_active INTEGER NOT NULL DEFAULT 0,
    imported_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY,
    list_id INTEGER NOT NULL REFERENCES price_lists(id) ON DELETE CASCADE,
    code TEXT NOT NULL,
    name TEXT NOT NULL,
    color TEXT NOT NULL,
    unit INTEGER NOT NULL,
    file_price TEXT NOT NULL,
    user_price TEXT NOT NULL,
    user_edited INTEGER NOT NULL DEFAULT 0,
    UNIQUE (list_id, code)
);

CREATE TABLE IF NOT EXISTS estimates (
    id INTEGER PRIMARY KEY,
    number TEXT NOT NULL DEFAULT '',
    estimate_date TEXT NOT NULL DEFAULT '',
    client TEXT NOT NULL DEFAULT '',
    object_name TEXT NOT NULL DEFAULT '',
    width_mm TEXT NOT NULL DEFAULT '',
    height_mm TEXT NOT NULL DEFAULT '',
    comment TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS estimate_lines (
    id INTEGER PRIMARY KEY,
    estimate_id INTEGER NOT NULL REFERENCES estimates(id) ON DELETE CASCADE,
    position INTEGER NOT NULL,
    code TEXT NOT NULL,
    name TEXT NOT NULL,
    color TEXT NOT NULL,
    unit INTEGER NOT NULL,
    qty TEXT NOT NULL,
    price TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_items_list ON items(list_id);
CREATE INDEX IF NOT EXISTS idx_lines_estimate ON estimate_lines(estimate_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_one_active_price
    ON price_lists(is_active) WHERE is_active = 1;
"""


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _money(value: Decimal) -> str:
    return f"{q2(value):.2f}"


class Database:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.conn.executescript(SCHEMA)

    def close(self) -> None:
        self.conn.close()

    def lists(self) -> list[dict]:
        rows = self.conn.execute(
            "SELECT id, name, filename, is_active, imported_at FROM price_lists ORDER BY id"
        ).fetchall()
        return [_list_dict(row) for row in rows]

    def active_list(self) -> dict | None:
        row = self.conn.execute(
            "SELECT id, name, filename, is_active, imported_at "
            "FROM price_lists WHERE is_active = 1"
        ).fetchone()
        return _list_dict(row) if row else None

    def find_list(self, filename: str) -> dict | None:
        stem = Path(filename).stem
        row = self.conn.execute(
            "SELECT id, name, filename, is_active, imported_at "
            "FROM price_lists WHERE filename = ? OR name = ?",
            (filename, stem),
        ).fetchone()
        return _list_dict(row) if row else None

    def create_list(self, name: str, filename: str) -> int:
        with self.conn:
            cursor = self.conn.execute(
                "INSERT INTO price_lists (name, filename, is_active, imported_at) "
                "VALUES (?, ?, 0, ?)",
                (name, filename, _now()),
            )
        return int(cursor.lastrowid)

    def touch_list(self, list_id: int, filename: str) -> None:
        with self.conn:
            self.conn.execute(
                "UPDATE price_lists SET name = ?, filename = ?, imported_at = ? WHERE id = ?",
                (Path(filename).stem, filename, _now(), list_id),
            )

    def set_active(self, list_id: int) -> None:
        with self.conn:
            self.conn.execute("UPDATE price_lists SET is_active = 0")
            cursor = self.conn.execute(
                "UPDATE price_lists SET is_active = 1 WHERE id = ?",
                (list_id,),
            )
            if cursor.rowcount != 1:
                raise LookupError("Прайс не найден")

    def upsert_items(self, list_id: int, rows: list) -> tuple[int, int]:
        """Обновить file_price. user_price не трогать, если позицию уже правили."""
        existing = {
            row[0]
            for row in self.conn.execute("SELECT code FROM items WHERE list_id = ?", (list_id,))
        }
        inserted = sum(1 for row in rows if row.code not in existing)
        updated = len(rows) - inserted
        payload = [
            (
                list_id,
                row.code,
                row.name,
                row.color,
                row.unit,
                _money(row.file_price),
                _money(row.file_price),
            )
            for row in rows
        ]
        with self.conn:
            self.conn.executemany(
                """
                INSERT INTO items (
                    list_id, code, name, color, unit, file_price, user_price, user_edited
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 0)
                ON CONFLICT(list_id, code) DO UPDATE SET
                    name = excluded.name,
                    color = excluded.color,
                    unit = excluded.unit,
                    file_price = excluded.file_price,
                    user_price = CASE
                        WHEN items.user_edited = 1 THEN items.user_price
                        ELSE excluded.file_price
                    END
                """,
                payload,
            )
        return inserted, updated

    def items(self, list_id: int) -> list[dict]:
        rows = self.conn.execute(
            """
            SELECT id, list_id, code, name, color, unit, file_price, user_price, user_edited
            FROM items WHERE list_id = ?
            """,
            (list_id,),
        ).fetchall()
        return [_item_dict(row) for row in rows]

    def set_user_price(self, item_id: int, price: Decimal) -> None:
        price = q2(price)
        row = self.conn.execute(
            "SELECT file_price FROM items WHERE id = ?",
            (item_id,),
        ).fetchone()
        if row is None:
            raise LookupError("Позиция не найдена")
        edited = 0 if price == Decimal(row["file_price"]) else 1
        with self.conn:
            self.conn.execute(
                "UPDATE items SET user_price = ?, user_edited = ? WHERE id = ?",
                (_money(price), edited, item_id),
            )

    def reset_user_price(self, item_id: int) -> None:
        with self.conn:
            cursor = self.conn.execute(
                "UPDATE items SET user_price = file_price, user_edited = 0 WHERE id = ?",
                (item_id,),
            )
            if cursor.rowcount != 1:
                raise LookupError("Позиция не найдена")

    def replace_user_prices(self, list_id: int) -> int:
        with self.conn:
            cursor = self.conn.execute(
                "UPDATE items SET user_price = file_price, user_edited = 0 WHERE list_id = ?",
                (list_id,),
            )
        return int(cursor.rowcount)

    def unit_labels(self) -> dict[int, str]:
        labels = dict(DEFAULT_UNIT_LABELS)
        rows = self.conn.execute(
            "SELECT key, value FROM settings WHERE key LIKE 'unit_%'"
        ).fetchall()
        for row in rows:
            try:
                code = int(str(row["key"]).split("_", 1)[1])
            except (IndexError, ValueError):
                continue
            text = str(row["value"]).strip()
            if code in labels and text:
                labels[code] = text
        return labels

    def set_unit_labels(self, labels: dict[int, str]) -> None:
        with self.conn:
            for code in UNIT_CODES:
                text = labels.get(code, "").strip() or DEFAULT_UNIT_LABELS[code]
                self.conn.execute(
                    """
                    INSERT INTO settings (key, value) VALUES (?, ?)
                    ON CONFLICT(key) DO UPDATE SET value = excluded.value
                    """,
                    (f"unit_{code}", text),
                )

    def save_estimate(self, estimate_id: int | None, header: dict, lines: list[dict]) -> int:
        now = _now()
        with self.conn:
            if estimate_id is None:
                cursor = self.conn.execute(
                    """
                    INSERT INTO estimates (
                        number, estimate_date, client, object_name, width_mm, height_mm,
                        comment, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        header["number"],
                        header["date"],
                        header["client"],
                        header["object"],
                        header["width"],
                        header["height"],
                        header["comment"],
                        now,
                        now,
                    ),
                )
                estimate_id = int(cursor.lastrowid)
            else:
                cursor = self.conn.execute(
                    """
                    UPDATE estimates SET
                        number = ?, estimate_date = ?, client = ?, object_name = ?,
                        width_mm = ?, height_mm = ?, comment = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        header["number"],
                        header["date"],
                        header["client"],
                        header["object"],
                        header["width"],
                        header["height"],
                        header["comment"],
                        now,
                        estimate_id,
                    ),
                )
                if cursor.rowcount != 1:
                    raise LookupError("Смета не найдена")
                self.conn.execute(
                    "DELETE FROM estimate_lines WHERE estimate_id = ?",
                    (estimate_id,),
                )
            self.conn.executemany(
                """
                INSERT INTO estimate_lines (
                    estimate_id, position, code, name, color, unit, qty, price
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        estimate_id,
                        index,
                        line["code"],
                        line["name"],
                        line["color"],
                        int(line["unit"]),
                        _money(line["qty"]),
                        _money(line["price"]),
                    )
                    for index, line in enumerate(lines)
                ],
            )
        return estimate_id

    def list_estimates(self) -> list[dict]:
        estimates = self.conn.execute(
            "SELECT * FROM estimates ORDER BY updated_at DESC, id DESC"
        ).fetchall()
        buckets: dict[int, list[sqlite3.Row]] = {}
        for line in self.conn.execute(
            "SELECT estimate_id, qty, price FROM estimate_lines"
        ):
            buckets.setdefault(int(line["estimate_id"]), []).append(line)
        result = []
        for row in estimates:
            lines = buckets.get(int(row["id"]), [])
            total = Decimal("0.00")
            for line in lines:
                total += line_total(Decimal(line["qty"]), Decimal(line["price"]))
            item = _header_dict(row)
            item["total"] = q2(total)
            item["positions"] = len(lines)
            result.append(item)
        return result

    def load_estimate(self, estimate_id: int) -> tuple[dict, list[dict]] | None:
        row = self.conn.execute(
            "SELECT * FROM estimates WHERE id = ?",
            (estimate_id,),
        ).fetchone()
        if row is None:
            return None
        lines = []
        for line in self.conn.execute(
            """
            SELECT code, name, color, unit, qty, price
            FROM estimate_lines
            WHERE estimate_id = ?
            ORDER BY position, id
            """,
            (estimate_id,),
        ):
            lines.append(
                {
                    "code": line["code"],
                    "name": line["name"],
                    "color": line["color"],
                    "unit": int(line["unit"]),
                    "qty": Decimal(line["qty"]),
                    "price": Decimal(line["price"]),
                }
            )
        return _header_dict(row), lines


def _list_dict(row: sqlite3.Row) -> dict:
    return {
        "id": int(row["id"]),
        "name": row["name"],
        "filename": row["filename"],
        "is_active": bool(row["is_active"]),
        "imported_at": row["imported_at"],
    }


def _item_dict(row: sqlite3.Row) -> dict:
    return {
        "id": int(row["id"]),
        "list_id": int(row["list_id"]),
        "code": row["code"],
        "name": row["name"],
        "color": row["color"],
        "unit": int(row["unit"]),
        "file_price": Decimal(row["file_price"]),
        "user_price": Decimal(row["user_price"]),
        "user_edited": bool(row["user_edited"]),
    }


def _header_dict(row: sqlite3.Row) -> dict:
    return {
        "id": int(row["id"]),
        "number": row["number"],
        "date": row["estimate_date"],
        "client": row["client"],
        "object": row["object_name"],
        "width": row["width_mm"],
        "height": row["height_mm"],
        "comment": row["comment"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }

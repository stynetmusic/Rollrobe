"""Поиск позиции по коду, имени и цвету."""

from __future__ import annotations

from decimal import Decimal


def filter_items(
    items: list[dict],
    query: str,
    *,
    show_zero: bool,
    limit: int,
) -> dict:
    """Вернуть показанные строки и счётчики.

    Пустой запрос оставляет весь список. Ноль прячется, если show_zero выключен.
    Точное совпадение кода, имени или цвета ставится выше, чем простое вхождение.
    """
    needle = query.strip().casefold()
    if needle:
        ranked: list[tuple[int, str, str, str, dict]] = []
        for item in items:
            code = item["code"].casefold()
            name = item["name"].casefold()
            color = item["color"].casefold()
            if needle not in code and needle not in name and needle not in color:
                continue
            if needle in (code, name, color):
                group = 0
            elif code.startswith(needle) or name.startswith(needle) or color.startswith(needle):
                group = 1
            else:
                group = 2
            ranked.append((group, name, color, code, item))
        ranked.sort(key=lambda row: row[:4])
        matched = [row[4] for row in ranked]
    else:
        matched = sorted(
            items,
            key=lambda item: (
                item["name"].casefold(),
                item["color"].casefold(),
                item["code"].casefold(),
            ),
        )

    hidden = 0
    visible: list[dict] = []
    for item in matched:
        price: Decimal = item["user_price"]
        if not show_zero and price == 0:
            hidden += 1
            continue
        visible.append(item)

    return {
        "matched": len(matched),
        "hidden": hidden,
        "visible_total": len(visible),
        "shown": visible[:limit],
    }

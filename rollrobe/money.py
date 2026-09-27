"""Деньги и количество. Считаем Decimal, не float."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

CENT = Decimal("0.01")
MAX_AMOUNT = Decimal("1000000000")


class MoneyError(ValueError):
    """Пользователь ввёл не число."""


def q2(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def parse_decimal(value: object) -> Decimal:
    """Принять 2296,26 или 2296.26 и округлить до копейки."""
    if isinstance(value, Decimal):
        number = value
    elif isinstance(value, bool) or value is None:
        raise MoneyError("пустое значение")
    elif isinstance(value, int):
        number = Decimal(value)
    elif isinstance(value, float):
        number = Decimal(str(value))
    else:
        text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
        if text == "":
            raise MoneyError("пустое значение")
        try:
            number = Decimal(text)
        except InvalidOperation as exc:
            raise MoneyError("не число") from exc
    if not number.is_finite():
        raise MoneyError("не число")
    if number < 0:
        raise MoneyError("отрицательное число")
    if number > MAX_AMOUNT:
        raise MoneyError("слишком большое число")
    return q2(number)


def parse_price_cell(value: object) -> Decimal:
    """Ячейка прайса. Пусто и ложь AluRoll — это цена 0, строку всё равно берём."""
    if value is None or value is False:
        return Decimal("0.00")
    if isinstance(value, str) and not value.strip():
        return Decimal("0.00")
    return parse_decimal(value)


def line_total(qty: Decimal, price: Decimal) -> Decimal:
    """Сумма строки: количество × цена, до копейки."""
    return q2(q2(qty) * q2(price))


def sum_lines(lines: list[dict]) -> Decimal:
    total = Decimal("0.00")
    for line in lines:
        total += line_total(line["qty"], line["price"])
    return q2(total)


def format_number(value: Decimal) -> str:
    """2 296,26 — как в прайсе, с пробелом между разрядами."""
    rounded = q2(value)
    sign = "-" if rounded < 0 else ""
    whole, frac = f"{abs(rounded):.2f}".split(".")
    groups: list[str] = []
    while whole:
        groups.append(whole[-3:])
        whole = whole[:-3]
    return sign + " ".join(reversed(groups)) + "," + frac


def format_edit(value: Decimal) -> str:
    """Число для поля ввода, без пробелов: 2296,26."""
    return f"{q2(value):.2f}".replace(".", ",")

"""Чистая логика материального и скользящего баланса.

Модуль не зависит от Telegram и базы данных, поэтому формулы можно проверять
обычными модульными тестами.
"""

import math
import re
from collections.abc import Iterable, Mapping, Sequence
from datetime import date, datetime, timedelta, timezone
from itertools import pairwise

FIELDS = [
    "kv4",
    "kv4d",
    "kv14",
    "kv32",
    "kv34",
    "kv34a",
    "kv102",
    "kv24p",
    "kv24hv",
    "kv28a1",
    "kv3",
    "kv3d",
    "kv15",
    "kv19",
    "kv31",
    "kv33",
    "kv101",
    "kv28a2",
    "kv44",
    "kv44d",
    "kv46",
    "kv46d",
    "kv74",
    "kv74d",
    "kv65mps",
    "kv65cpo",
    "kv66mps",
    "kv66cpo",
    "kv84mps",
    "kv84cpo",
    "kv63",
    "kv61",
]


MONTH_NAMES = {
    1: ("январь", "января", "қаңтар"),
    2: ("февраль", "февраля", "ақпан"),
    3: ("март", "марта", "наурыз"),
    4: ("апрель", "апреля", "сәуір"),
    5: ("май", "мая", "мамыр"),
    6: ("июнь", "июня", "маусым"),
    7: ("июль", "июля", "шілде"),
    8: ("август", "августа", "тамыз"),
    9: ("сентябрь", "сентября", "қыркүйек"),
    10: ("октябрь", "октября", "қазан"),
    11: ("ноябрь", "ноября", "қараша"),
    12: ("декабрь", "декабря", "желтоқсан"),
}


def _number(value) -> float:
    try:
        number = float(value or 0.0)
    except (TypeError, ValueError):
        return 0.0
    return number if math.isfinite(number) else 0.0


def bal1(data: Mapping) -> float | None:
    """Баланс первой очереди, % от конвейера 4."""
    base = _number(data.get("kv4"))
    if not base:
        return None
    result = (
        _number(data.get("kv102"))
        + _number(data.get("kv34"))
        + _number(data.get("kv24p"))
        + _number(data.get("kv24hv"))
        - _number(data.get("kv28a1"))
        - base
    )
    return result / base * 100


def bal2(data: Mapping) -> float | None:
    """Баланс второй очереди, % от конвейера 3."""
    base = _number(data.get("kv3"))
    if not base:
        return None
    result = (
        _number(data.get("kv101"))
        + _number(data.get("kv33"))
        - _number(data.get("kv28a2"))
        - base
    )
    return result / base * 100


def balc1(data: Mapping) -> float | None:
    """Баланс сепарации первой очереди, % от конвейера 4."""
    base = _number(data.get("kv4"))
    if not base:
        return None
    result = (
        _number(data.get("kv102"))
        + _number(data.get("kv24hv"))
        + _number(data.get("kv24p"))
        + _number(data.get("kv32"))
        - _number(data.get("kv28a1"))
        - _number(data.get("kv14"))
    )
    return result / base * 100


def balc2(data: Mapping) -> float | None:
    """Баланс сепарации второй очереди, % от конвейера 3."""
    base = _number(data.get("kv3"))
    if not base:
        return None
    result = (
        _number(data.get("kv101"))
        + _number(data.get("kv31"))
        - _number(data.get("kv28a2"))
        - _number(data.get("kv15"))
    )
    return result / base * 100


def calculate_balances(data: Mapping) -> dict:
    return {
        "b1": bal1(data),
        "bc1": balc1(data),
        "b2": bal2(data),
        "bc2": balc2(data),
    }


def sum_period(rows: Iterable[Mapping], fields: Sequence[str] = FIELDS) -> dict:
    """Сначала суммирует тоннаж, не усредняя суточные проценты."""
    rows = list(rows)
    return {key: sum(_number(row.get(key)) for row in rows) for key in fields}


def row_date(row: Mapping) -> date:
    raw = str(row.get("report_date") or "").strip()
    if raw:
        try:
            return date.fromisoformat(raw[:10])
        except ValueError:
            pass
    return date(int(row["year"]), int(row["month"]), int(row["day_num"]))


def format_period_label(rows: Sequence[Mapping]) -> str:
    if not rows:
        return "—"
    dates = sorted(row_date(row) for row in rows)
    first, last = dates[0], dates[-1]
    if first == last:
        return first.strftime("%d.%m.%Y")
    if first.year == last.year and first.month == last.month:
        return f"{first.day:02d}–{last.day:02d}.{last.month:02d}.{last.year}"
    if first.year == last.year:
        return f"{first.day:02d}.{first.month:02d}–{last.day:02d}.{last.month:02d}.{last.year}"
    return f"{first:%d.%m.%Y}–{last:%d.%m.%Y}"


def is_consecutive_period(rows: Sequence[Mapping]) -> bool:
    dates = sorted(row_date(row) for row in rows)
    return all(
        current - previous == timedelta(days=1) for previous, current in pairwise(dates)
    )


def rolling_snapshots(
    rows: Sequence[Mapping],
    periods: Sequence[int] = (1, 2, 3),
    fields: Sequence[str] = FIELDS,
) -> list:
    """Возвращает расчёты по последним 1/2/3 доступным завершённым суткам."""
    ordered = sorted(rows, key=row_date)
    snapshots = []
    for days in periods:
        if days < 1 or len(ordered) < days:
            continue
        selected = ordered[-days:]
        total = sum_period(selected, fields)
        snapshots.append(
            {
                "days": days,
                "rows": selected,
                "total": total,
                "balances": calculate_balances(total),
                "label": format_period_label(selected),
                "consecutive": is_consecutive_period(selected),
            }
        )
    return snapshots


def balance_status(
    value: float | None,
    warn_pct: float = 2.0,
    crit_pct: float = 5.0,
) -> str:
    if value is None or not math.isfinite(value):
        return "none"
    absolute = abs(value)
    if absolute <= warn_pct:
        return "ok"
    if absolute <= crit_pct:
        return "warn"
    return "crit"


def infer_report_period(period_text: str, reference=None) -> tuple:
    """Определяет год и месяц из подписи отчёта.

    Если в подписи указан только месяц словами, год выбирается относительно
    даты загрузки. Например, отчёт «за декабрь», загруженный в январе,
    относится к предыдущему году.
    """
    if reference is None:
        ref_date = datetime.now(timezone.utc).date()
    elif isinstance(reference, datetime):
        ref_date = reference.date()
    else:
        ref_date = reference

    text = str(period_text or "").strip().lower().replace("ё", "е")

    # Форматы с числовым месяцем: 07.2026, 2026-07 и полные даты.
    month_year = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[./-](20\d{2})(?!\d)", text)
    if month_year:
        return int(month_year.group(2)), int(month_year.group(1)), True

    year_month = re.search(r"(?<!\d)(20\d{2})[./-](0?[1-9]|1[0-2])(?!\d)", text)
    if year_month:
        return int(year_month.group(1)), int(year_month.group(2)), True

    year_match = re.search(r"(?<!\d)(20\d{2})(?!\d)", text)
    for month, names in MONTH_NAMES.items():
        if not any(
            re.search(rf"(?<!\w){re.escape(name)}(?!\w)", text) for name in names
        ):
            continue
        if year_match:
            year = int(year_match.group(1))
        else:
            year = ref_date.year - (1 if month > ref_date.month else 0)
        return year, month, True

    return ref_date.year, ref_date.month, False

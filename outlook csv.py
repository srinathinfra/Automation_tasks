"""
outlook_csv.py
--------------
Reads the CSV that Outlook produces from File > Open & Export > Import/Export
> Export to a file > Comma Separated Values, and turns it into clean per-day
meeting hours.

Deliberately forgiving, because that export varies by Outlook version and
Windows locale:
  * figures out for itself whether dates are MM/DD/YYYY or DD/MM/YYYY
  * copes with 12-hour and 24-hour clocks
  * handles a combined "Start"/"End" datetime column as well as the usual
    split Start Date / Start Time pair
  * tries UTF-8, then Windows-1252, then Latin-1 encodings

It also does three things that stop your meeting hours from lying to you:
  * double-booked meetings are merged, so two overlapping 1-hour calls count
    as the elapsed time, not 2 hours
  * anything crossing midnight is split across the two days
  * cancelled items, all-day items and Free-time blocks are dropped
"""

from __future__ import annotations

import csv
import datetime as dt
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

CANCELLED_RE = re.compile(r"^\s*(cancell?ed)\s*:", re.IGNORECASE)

# Outlook writes "Show time as" as an integer code, not a word. Newer builds
# and Graph-based exports use the words instead, so accept both.
SHOW_AS_CODES = {
    "0": "Free",
    "1": "Tentative",
    "2": "Busy",
    "3": "Out of Office",
    "4": "Working Elsewhere",
}

DATE_FORMATS_DAY_FIRST = ["%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%b-%Y", "%Y-%m-%d"]
DATE_FORMATS_MONTH_FIRST = ["%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%b %d, %Y", "%Y-%m-%d"]
TIME_FORMATS = ["%I:%M:%S %p", "%I:%M %p", "%H:%M:%S", "%H:%M", "%I:%M:%S%p", "%I:%M%p"]


class OutlookParseError(RuntimeError):
    pass


@dataclass
class Meeting:
    subject: str
    start: dt.datetime
    end: dt.datetime
    organizer: str = ""
    location: str = ""
    show_as: str = ""
    categories: str = ""

    @property
    def date(self) -> dt.date:
        return self.start.date()

    @property
    def hours(self) -> float:
        return round((self.end - self.start).total_seconds() / 3600.0, 4)

    @property
    def time_range(self) -> str:
        return f"{self.start:%H:%M}-{self.end:%H:%M}"


# --------------------------------------------------------------------------
# Low-level parsing helpers
# --------------------------------------------------------------------------
def _norm_header(h: str) -> str:
    return re.sub(r"\s+", " ", (h or "").replace("\ufeff", "").strip().lower())


def _read_rows(path: str) -> List[Dict[str, str]]:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=encoding, newline="") as fh:
                reader = csv.reader(fh)
                try:
                    raw_header = next(reader)
                except StopIteration:
                    raise OutlookParseError(f"{path} is empty.")
                header = [_norm_header(h) for h in raw_header]
                rows = []
                for values in reader:
                    if not any((v or "").strip() for v in values):
                        continue
                    row = {}
                    for i, name in enumerate(header):
                        if name:
                            row[name] = values[i].strip() if i < len(values) else ""
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    raise OutlookParseError(f"Could not decode {path}: {last_error}")


def _pick(row: Dict[str, str], *names: str) -> str:
    for n in names:
        if n in row and row[n]:
            return row[n]
    return ""


def _looks_boolean_true(value: str) -> bool:
    return (value or "").strip().lower() in ("true", "yes", "1", "y")


def normalise_show_as(value: str) -> str:
    """Turn either '0' or 'Free' into the word 'Free'."""
    raw = (value or "").strip()
    if not raw:
        return ""
    if raw in SHOW_AS_CODES:
        return SHOW_AS_CODES[raw]
    lowered = raw.lower()
    for label in SHOW_AS_CODES.values():
        if lowered == label.lower():
            return label
    if lowered in ("oof", "outofoffice", "out of office"):
        return "Out of Office"
    return raw


def detect_day_first(samples: Sequence[str], default_day_first: bool = False) -> bool:
    """Work out whether numeric dates are DD/MM or MM/DD by finding a value
    that can only be read one way."""
    for raw in samples:
        parts = re.split(r"[/-]", (raw or "").strip().split(" ")[0])
        if len(parts) < 2 or not parts[0].isdigit() or not parts[1].isdigit():
            continue
        first, second = int(parts[0]), int(parts[1])
        if first > 12 and second <= 12:
            return True   # 25/07 - must be day first
        if second > 12 and first <= 12:
            return False  # 07/25 - must be month first
    return default_day_first


def _parse_date(raw: str, day_first: bool) -> Optional[dt.date]:
    raw = (raw or "").strip()
    if not raw:
        return None
    formats = DATE_FORMATS_DAY_FIRST if day_first else DATE_FORMATS_MONTH_FIRST
    for fmt in formats:
        try:
            return dt.datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def _parse_time(raw: str) -> Optional[dt.time]:
    raw = (raw or "").strip().replace("\u202f", " ")
    if not raw:
        return None
    for fmt in TIME_FORMATS:
        try:
            return dt.datetime.strptime(raw, fmt).time()
        except ValueError:
            continue
    return None


def _parse_datetime(raw: str, day_first: bool) -> Optional[dt.datetime]:
    """For exports that put date and time in one column."""
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        return dt.datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        pass
    bits = raw.split(" ", 1)
    day = _parse_date(bits[0], day_first)
    if day is None:
        return None
    clock = _parse_time(bits[1]) if len(bits) > 1 else dt.time(0, 0)
    return dt.datetime.combine(day, clock or dt.time(0, 0))


# --------------------------------------------------------------------------
# Main entry point
# --------------------------------------------------------------------------
def parse_outlook_csv(
    path: str,
    date_from: Optional[dt.date] = None,
    date_to: Optional[dt.date] = None,
    date_format: str = "auto",
    skip_all_day: bool = True,
    skip_free: bool = True,
    skip_tentative: bool = False,
    skip_cancelled: bool = True,
    min_minutes: int = 5,
    exclude_subjects: Iterable[str] = (),
) -> Tuple[List[Meeting], List[str]]:
    """Return (meetings, warnings)."""
    rows = _read_rows(path)
    if not rows:
        return [], [f"{path} contained a header but no appointments."]

    warnings: List[str] = []

    if date_format.strip().lower() == "auto":
        samples = [_pick(r, "start date", "start") for r in rows[:400]]
        day_first = detect_day_first(samples, default_day_first=False)
    else:
        day_first = date_format.strip().lower() in ("eu", "uk", "dayfirst", "dd/mm/yyyy")

    excludes = [e.strip().lower() for e in exclude_subjects if e and e.strip()]
    meetings: List[Meeting] = []
    unparsed = 0

    for row in rows:
        subject = _pick(row, "subject", "title") or "(no subject)"

        if skip_cancelled and CANCELLED_RE.match(subject):
            continue
        if skip_all_day and _looks_boolean_true(_pick(row, "all day event", "all day")):
            continue

        show_as = normalise_show_as(_pick(row, "show time as", "show as", "busy status"))
        if skip_free and show_as == "Free":
            continue
        if skip_tentative and show_as == "Tentative":
            continue
        if any(x in subject.lower() for x in excludes):
            continue

        # Split date/time columns first, combined column as fallback.
        start_date = _parse_date(_pick(row, "start date"), day_first)
        if start_date is not None:
            start_time = _parse_time(_pick(row, "start time")) or dt.time(0, 0)
            start = dt.datetime.combine(start_date, start_time)
            end_date = _parse_date(_pick(row, "end date"), day_first) or start_date
            end_time = _parse_time(_pick(row, "end time")) or start_time
            end = dt.datetime.combine(end_date, end_time)
        else:
            start = _parse_datetime(_pick(row, "start", "start time"), day_first)
            end = _parse_datetime(_pick(row, "end", "end time"), day_first)
            if start is None or end is None:
                unparsed += 1
                continue

        if end <= start:
            continue
        if (end - start).total_seconds() < min_minutes * 60:
            continue

        meetings.append(
            Meeting(
                subject=subject,
                start=start,
                end=end,
                organizer=_pick(row, "meeting organizer", "organizer"),
                location=_pick(row, "location"),
                show_as=show_as,
                categories=_pick(row, "categories"),
            )
        )

    if unparsed:
        warnings.append(
            f"{unparsed} row(s) had dates this script could not read. If the whole "
            f"file looks wrong, set date_format to 'us' or 'eu' in config.ini."
        )

    meetings = split_across_midnight(meetings)

    if date_from and date_to:
        meetings = [m for m in meetings if date_from <= m.date <= date_to]

    meetings.sort(key=lambda m: (m.start, m.end))
    return meetings, warnings


def split_across_midnight(meetings: List[Meeting]) -> List[Meeting]:
    """An event running 23:00 to 01:00 becomes two events, one per day."""
    out: List[Meeting] = []
    for m in meetings:
        if m.start.date() == m.end.date():
            out.append(m)
            continue
        cursor = m.start
        while cursor.date() < m.end.date():
            midnight = dt.datetime.combine(
                cursor.date() + dt.timedelta(days=1), dt.time(0, 0)
            )
            out.append(
                Meeting(m.subject, cursor, midnight, m.organizer, m.location,
                        m.show_as, m.categories)
            )
            cursor = midnight
        if cursor < m.end:
            out.append(
                Meeting(m.subject, cursor, m.end, m.organizer, m.location,
                        m.show_as, m.categories)
            )
    return out


def merged_hours_by_day(meetings: List[Meeting]) -> Dict[dt.date, float]:
    """Total elapsed busy hours per day, with overlaps counted only once.

    Without this, a day where two 1-hour calls sit on top of each other reads
    as 2 hours of meetings when you were really only busy for 1.
    """
    by_day: Dict[dt.date, List[Tuple[dt.datetime, dt.datetime]]] = {}
    for m in meetings:
        by_day.setdefault(m.date, []).append((m.start, m.end))

    totals: Dict[dt.date, float] = {}
    for day, spans in by_day.items():
        spans.sort()
        merged: List[List[dt.datetime]] = []
        for start, end in spans:
            if merged and start <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], end)
            else:
                merged.append([start, end])
        seconds = sum((e - s).total_seconds() for s, e in merged)
        totals[day] = round(seconds / 3600.0, 4)
    return totals


def raw_hours_by_day(meetings: List[Meeting]) -> Dict[dt.date, float]:
    """Straight sum of every meeting, overlaps double-counted."""
    totals: Dict[dt.date, float] = {}
    for m in meetings:
        totals[m.date] = round(totals.get(m.date, 0.0) + m.hours, 4)
    return totals


def overlap_hours_by_day(meetings: List[Meeting]) -> Dict[dt.date, float]:
    raw = raw_hours_by_day(meetings)
    merged = merged_hours_by_day(meetings)
    return {d: round(raw[d] - merged.get(d, 0.0), 4) for d in raw}

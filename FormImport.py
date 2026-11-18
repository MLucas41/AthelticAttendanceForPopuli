"""
FormImport.py

Reads the Excel export of the Microsoft Forms trip request and turns each
response (row) into a Trip: who is traveling (name + Populi ID) and when.

This file only reads the spreadsheet. It never talks to Populi, so it can be
tested on its own:

    uv run FormImport.py "Baseball_Travel_Roster_and_Dates.xlsx"

Needs openpyxl:  uv add openpyxl
"""

import re
from dataclasses import dataclass, field
from datetime import date, datetime, time

import openpyxl

# The form only asks for dates, not times, so a trip is treated as whole days:
# the team is gone from the start of the leave date through the end of the
# return date. Set this to False if players are normally back in time for the
# return date's classes (the trip then ends at the very start of that day).
RETURN_DAY_INCLUSIVE = True

# Column headers are matched by keyword, so small wording changes in the form
# won't break the import. Each entry is (field name, keywords - any one matches).
COLUMNS = {
    "players": ("traveling", "which players"),
    "leave": ("leaving",),
    "return": ("returning",),
    "who": ("name",),          # who submitted the form (optional)
}

# "Last, First (12345678)" - the ID is the last (...) group of digits, so names
# that contain their own parentheses still parse.
ENTRY = re.compile(r"^(?P<name>.+?)\s*\((?P<id>\d+)\)\s*$")

DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y")


@dataclass
class Trip:
    row: int                          # spreadsheet row number, for error messages
    submitted_by: str
    players: list[dict]               # [{"id": "24557570", "first": "Malikhi", "last": "Emery-Henderson"}]
    start: datetime
    end: datetime
    warnings: list[str] = field(default_factory=list)

    @property
    def label(self):
        """One-line description for a pick list."""
        d0, d1 = self.start.strftime("%b %d"), self.end.strftime("%b %d")
        dates = d0 if self.start.date() == self.end.date() else f"{d0} - {d1}"
        who = f"{self.submitted_by}: " if self.submitted_by else ""
        n = len(self.players)
        return f"Row {self.row}: {who}{dates} ({n} player{'' if n == 1 else 's'})"


def parse_players(cell):
    """Split the semicolon-separated player cell. Returns (players, warnings)."""
    players, warnings, seen = [], [], set()
    for part in str(cell or "").split(";"):
        part = part.strip()
        if not part:
            continue                  # the cell ends with a trailing ";"
        m = ENTRY.match(part)
        if not m:
            warnings.append(f"Could not read a Populi ID from: {part!r}")
            continue
        pid = m.group("id")
        if pid in seen:
            continue                  # same player picked twice
        seen.add(pid)
        last, _, first = m.group("name").partition(", ")
        if not first:                 # no comma: treat the whole thing as a first name
            first, last = last, ""
        players.append({"id": pid, "first": first.strip(), "last": last.strip()})
    return players, warnings


def _to_date(value, what, row):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Row {row}: could not read the {what} date {value!r}")


def _find_columns(headers):
    """Map each field to a column index by matching header keywords."""
    found = {}
    for key, keywords in COLUMNS.items():
        for i, header in enumerate(headers):
            h = str(header or "").lower()
            if i not in found.values() and any(k in h for k in keywords):
                found[key] = i
                break
    missing = [k for k in ("players", "leave", "return") if k not in found]
    if missing:
        raise ValueError(
            f"Could not find the {', '.join(missing)} column(s). "
            f"Headers in the file: {[str(h) for h in headers if h]}"
        )
    return found


def load_trips(path):
    """Read every response in the export and return a list of Trip objects."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    if len(rows) < 2:
        raise ValueError("The file has no responses in it.")

    cols = _find_columns(rows[0])
    trips = []
    for row_num, row in enumerate(rows[1:], start=2):
        if not any(v is not None and str(v).strip() for v in row):
            continue                  # blank row

        players, warnings = parse_players(row[cols["players"]])
        if not players:
            raise ValueError(f"Row {row_num}: no players were selected.")

        leave = _to_date(row[cols["leave"]], "leaving", row_num)
        back = _to_date(row[cols["return"]], "returning", row_num)
        if back < leave:
            raise ValueError(f"Row {row_num}: the return date is before the leave date.")

        start = datetime.combine(leave, time.min)
        end = datetime.combine(back, time(23, 59, 59) if RETURN_DAY_INCLUSIVE else time.min)
        if end <= start:              # same-day trip with the exclusive setting
            end = datetime.combine(back, time(23, 59, 59))

        who = ""
        if "who" in cols and row[cols["who"]]:
            who = str(row[cols["who"]]).strip()

        trips.append(Trip(row_num, who, players, start, end, warnings))
    return trips


if __name__ == "__main__":
    import sys

    for trip in load_trips(sys.argv[1]):
        print(trip.label)
        print("  from", trip.start, "to", trip.end)
        for p in trip.players:
            print("  ", p["id"], p["last"], "|", p["first"])
        for w in trip.warnings:
            print("  WARNING:", w)
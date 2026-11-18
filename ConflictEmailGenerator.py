"""
ConflictEmailGenerator.py

Finds which athletes will miss class during a team trip and drafts an email to
each affected professor.

How it works:
    1. Read a roster CSV exported from Populi (one row per player).
    2. For each player, pull their class schedule for the term from Populi.
    3. Compare every class meeting against the trip's departure/return times.
    4. Group the missed classes by course section and print one email per
       section, listing the students who will be absent and addressed to the
       section's primary instructor.

Nothing is sent or changed in Populi. The script only reads data and prints drafts.
"""

import os
import csv
import PopuliPuller
import win32com.client
import time as timely
from dotenv import load_dotenv
from datetime import datetime, timedelta, time, date
from dataclasses import dataclass, field
from collections import defaultdict
from pathlib import Path

# Load the API key from the .env file into environment variables
load_dotenv()

URL = "https://corban.populiweb.com/api2/"         # Populi API base URL
TERM_ID = PopuliPuller.getTermID("Current Term")      # Populi's numeric ID for the term being checked

ROSTER_DIR = Path("RosterCSVs")
_player_cache = {}

# Courses that are never reported to professors (everyone takes them, so they'd add noise).
# Names are compared in lowercase.
SKIP_COURSES = {"chapel"}

# Converts Populi's two-letter day codes to Python weekday numbers (Monday = 0 ... Sunday = 6)
DAY_CODES = {"mo": 0, "tu": 1, "we": 2, "th": 3, "fr": 4, "sa": 5, "su": 6}

_players_by_id = {}   # Populi ID -> Player, so each schedule is downloaded once

def subsequentAccessTest():
    "The thing the thing"

def _openApp(App):
    app = win32com.client.Dispatch(app)
    app.Visible = True
    print("App opener ran")
    return

def _testDraftCount():
    
    outlook = win32com.client.Dispatch("Outlook.Application")

    namespace = outlook.GetNamespace("MAPI")
    drafts_folder = namespace.GetDefaultFolder(16)
    draft_count = drafts_folder.Items.Count

    print(f"Number of Drafts: {draft_count}")

    return

def _test_draft_saving():
    outlook = win32com.client.Dispatch("Outlook.Application")
    print("Outlook version:", outlook.Version)
    print("Default store:", outlook.Session.DefaultStore.DisplayName)

    mail = outlook.CreateItem(0)
    mail.To = "lucas.mertlich@gmail.com"
    mail.Subject = "TEST"
    mail.Body = "This is to test the functionality of the draft saving"
    mail.Save()

    print("Saved in folder:", mail.Parent.FolderPath)
    print("Drafts folder item count:", outlook.Session.GetDefaultFolder(16).Items.Count)  # 16 = Drafts

    mail.Display()   # also pops the message open so you can see it exists

# ---------------------------------------------------------------------------
# Creates drafts in the outlook app
# ---------------------------------------------------------------------------
def create_outlook_drafts(drafts, send=False, sendtoLuke=True):
    """Create one Outlook message per draft. Saves to Drafts unless send=True.

    Returns (number created, list of course names skipped for having no email).
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    folder = outlook.Session.GetDefaultFolder(16)          # Drafts
    tag = "Absence " + datetime.now().strftime("%Y-%m-%d %H%M")

    created, skipped, failed = 0, [], []
    for d in drafts:
        if not d["to"]:
            skipped.append(d["course"])
            continue
        try:
            mail = outlook.CreateItem(0)
            mail.To = d["to"]
            mail.Subject = d["subject"]
            mail.Body = d["body"]
            mail.Categories = tag
            if send:
                mail.Send()
            else:
                if sendtoLuke:
                    mail.To = "lukem@corban.edu"
                    mail.Send()
                mail.Save()
            created += 1
        except Exception as e:
            failed.append((d["course"], str(e)))

    timely.sleep(5)   # give Outlook a moment to catch up before counting
    found = folder.Items.Restrict(f"[Categories] = '{tag}'").Count
    print(f"Tag: {tag}")
    print(f"Tried {len(drafts)}: created {created}, skipped {len(skipped)}, "
          f"failed {len(failed)}, found in Drafts with tag: {found}")
    for course, err in failed:
        print("FAILED:", course, "-", err)
    return created, skipped

# ---------------------------------------------------------------------------
# Gets the team names from the CSV files
# ---------------------------------------------------------------------------
def list_teams():
    """Team names, taken from the CSV filenames (Baseball.csv -> 'Baseball')."""
    return sorted(p.stem for p in ROSTER_DIR.glob("*.csv"))

# ---------------------------------------------------------------------------
# Gets the team names from the CSV files
# ---------------------------------------------------------------------------
def load_players(team):
    if team not in _player_cache:
        with open(ROSTER_DIR / f"{team}.csv", newline="", encoding="utf-8-sig") as f:
            _player_cache[team] = build_players(csv.DictReader(f), TERM_ID)
    return _player_cache[team]

def load_players_from_trip(trip):
    players = []
    for entry in trip.players:
        pid = entry["id"]
        if pid not in _players_by_id:
            p = Player(populi_id=pid, first=entry["first"], last=entry["last"])
            for c in PopuliPuller.GetTermScheduleFromPopuliID(pid, TERM_ID):
                p.courses.append(Course.from_raw(c["name"], c["offering_id"], c["days"], c["start"], c["end"]))
            _players_by_id[pid] = p
        players.append(_players_by_id[pid])
    return players

def generate_drafts_for_trip(trip):
    return generate_drafts_for_players(load_players_from_trip(trip), trip.start, trip.end)
# ---------------------------------------------------------------------------
# Generates email drafts based on the professor, rather than the classes
# ---------------------------------------------------------------------------
def generate_drafts_for_players(team, trip_start, trip_end):
    """Return one draft per professor (or per course, when no professor email exists)."""
    players = load_players(team)
    conflicts = find_conflicts(players, trip_start, trip_end)

    # Group the conflicts by professor email
    groups = {}   # key -> {"faculty": dict, "courses": [(course_name, students), ...]}
    for offering_id, info in conflicts.items():
        faculty = PopuliPuller.get_faculty(offering_id, TERM_ID)
        email = (faculty.get("email") or "").strip().lower()
        # No email: keep each course separate so none get merged under a blank address
        key = email or f"no-email-{offering_id}"
        group = groups.setdefault(key, {"faculty": faculty, "courses": []})
        group["courses"].append((info["name"], info["students"]))

    drafts = []
    for key, g in groups.items():
        faculty, courses = g["faculty"], g["courses"]
        to = (faculty.get("email") or "").strip() or None
        n = len(courses)
        last = faculty.get("last_name")

        if n == 1:
            label = courses[0][0]                        # course name, as before
            subject = f"Athletic absences due to travel for {team}: from {format_date_range(trip_start, trip_end)} in {courses[0][0]}"
        else:
            label = f"{last or faculty.get('name') or 'Unknown professor'} ({n} courses)"
            subject = f"Athletic absences due to travel for {team}: from {format_date_range(trip_start, trip_end)}"

        drafts.append({
            "course": label,       # shown in the app's list (key name kept so app.py is unchanged)
            "to": to,
            "subject": subject,
            "body": build_group_body(courses, trip_start, trip_end, faculty),
        })

    drafts.sort(key=lambda d: d["course"].lower())
    return drafts

# ---------------------------------------------------------------------------
# Builds the body of each email to include each course the professor teaches, and the students that will miss those classes
# ---------------------------------------------------------------------------
def build_group_body(courses, trip_start, trip_end, faculty):
    """One email body covering every course a professor has with absences.

    courses: list of (course_name, students) pairs
    """
    last = faculty.get("last_name")
    greeting = f"Hello Professor {last}," if last else "Hello Professor,"
    date_range = format_date_range(trip_start, trip_end)

    if len(courses) == 1:
        course_name, students = courses[0]
        intro = (
            f"This email is to notify you that the following students will be absent from "
            f"{course_name} from {date_range} for their sport. "
            f"Their attendance should already be updated.\n\n"
        )
        listing = "\n".join(sorted(students))
    else:
        intro = (
            f"This email is to notify you that the following students will be absent from "
            f"the classes listed below from {date_range} for their sport. "
            f"Their attendance should already be updated.\n\n"
        )
        sections = []
        for course_name, students in sorted(courses, key=lambda c: c[0].lower()):
            sections.append(course_name + "\n" + "\n".join(f"  {s}" for s in sorted(students)))
        listing = "\n\n".join(sections)

    return f"{greeting}\n\n{intro}{listing}\n\nThank you!"

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Course:
    """One scheduled meeting pattern of a course section (e.g. "MWF 10:00-10:50").

    A course that meets at two different times (like a lecture plus a lab) is
    stored as two Course objects with the same offering_id.
    frozen=True makes instances read-only so they can't be changed by accident.
    """
    name: str
    offering_id: str            # Populi's ID for this specific section; used to look up the professor
    days: tuple[int, ...]       # weekday numbers the class meets, e.g. (0, 2, 4) for Mon/Wed/Fri
    start: time | None          # time of day the class starts (None if it has no meeting time)
    end: time | None            # time of day the class ends (None if it has no meeting time)

    @classmethod
    def from_raw(cls, name, offering_id, days, start, end):
        """Build a Course from the raw strings that StudentIDtoSchedule returns.

        Populi gives days like ['mo', 'we'] and times like '10:30AM'. This converts
        them to weekday numbers and real time objects once, here, so the rest of
        the code never has to deal with text.
        Courses with no meeting time (internships, advising, clubs) come back as
        None values and become a Course with no meetings.
        """
        if not isinstance(days, (list, tuple)) or "No meetingt time" in (start, end):
            return cls(name, None, (), None, None)
        return cls(
            name=name,
            offering_id=offering_id,
            days=tuple(DAY_CODES[d] for d in days),
            start=datetime.strptime(start, "%I:%M%p").time(),   # "10:30AM" -> 10:30
            end=datetime.strptime(end, "%I:%M%p").time(),
        )

    @property
    def has_meetings(self):
        """True if this course actually has scheduled days and times."""
        return bool(self.days) and self.start is not None
    
    def missed_between(self, trip_start: datetime, trip_end: datetime):
        """Yield (class_start, class_end) for each meeting that overlaps the trip.

        Walks through every calendar day from departure to return. On each day
        the class meets, it builds that day's start/end datetimes and checks
        whether they overlap the trip window. A class counts as missed if any
        part of it falls inside the trip, so leaving at 10:30 AM still misses
        a class that runs 10:00-10:50.
        """
        if not self.has_meetings:
            return  # nothing scheduled, so nothing can be missed
        day = trip_start.date()
        while day <= trip_end.date():
            if day.weekday() in self.days:
                s = datetime.combine(day, self.start)   # this day's class start
                e = datetime.combine(day, self.end)     # this day's class end
                # Overlap test: the class starts before the trip ends AND ends after it begins
                if s < trip_end and e > trip_start:
                    yield s, e
            day += timedelta(days=1)

@dataclass
class Player:
    """A student athlete and the list of courses they're enrolled in this term."""
    populi_id: str              # Populi's internal person ID (from the roster CSV)
    first: str = ""
    last: str = ""
    courses: list[Course] = field(default_factory=list)   # each player gets their own list

    @property
    def full_name(self):
        """First and last name joined, with extra whitespace trimmed."""
        return f"{self.first} {self.last}".strip()

    def missed_classes(self, trip_start, trip_end):
        """Return this player's missed class sessions as (start, course name, end) tuples, in date order.

        (Not currently called by main(); find_conflicts does this work for the whole roster.)
        """
        rows = []
        for course in self.courses:
            for s, e in course.missed_between(trip_start, trip_end):
                rows.append((s, course.name, e))
        rows.sort()
        return rows


# ---------------------------------------------------------------------------
# Finding conflicts
# ---------------------------------------------------------------------------

def find_conflicts(players, trip_start, trip_end):
    """Group every missed class by course section, across the whole roster.

    Returns a dict keyed by offering_id (the section) so that two sections with
    the same name but different professors stay separate:
        {offering_id: {"name": str, "students": {student_name: [(start, end), ...]}}}
    """
    conflicts = {}
    for p in players:
        for course in p.courses:
            if course.name.lower() in SKIP_COURSES:
                continue   # e.g. Chapel: don't report it
            sessions = list(course.missed_between(trip_start, trip_end))
            if not sessions:
                continue   # no overlap with the trip, so no conflict
            # Create the section's entry the first time a student misses it
            entry = conflicts.setdefault(
                course.offering_id, {"name": course.name, "students": {}}
            )
            # Add this student's missed sessions (extend, in case the section meets more than once)
            entry["students"].setdefault(p.full_name, []).extend(sessions)
    return conflicts


# ---------------------------------------------------------------------------
# Building the email text
# ---------------------------------------------------------------------------

def _fmt_day(dt):
    """Format a date like 'Friday, October 23'."""
    return f"{dt.strftime('%A, %B')} {dt.day}"      # avoids %-d / %#d platform differences

def format_date_range(start, end):
    """Format the trip dates: one date if it's a single day, otherwise 'X to Y'."""
    if start.date() == end.date():
        return _fmt_day(start)
    return f"{_fmt_day(start)} to {_fmt_day(end)}"

def build_email(course_name, students, trip_start, trip_end, faculty):
    """Return the text of one email for one course section.

    course_name: name of the section
    students:    dict of student names who will miss it (only the names are used)
    faculty:     dict from StudentIDtoSchedule.get_faculty, with "last_name" and "email"
                 (either can be missing; the email shows a placeholder instead)
    """
    last = faculty.get("last_name")
    # If Populi has no email on file, say so in the draft rather than leaving it blank
    if(faculty.get("email") == None):
        email = "No Email Found for Professor"
    else:
        email = faculty.get("email")
    
    # Use "Professor Smith" when we know the last name, otherwise a generic greeting
    greeting = f"Hello Professor {last}," if last else "Hello Professor,"
    names = "\n".join(sorted(students))     # one student per line, alphabetical
    date_range = format_date_range(trip_start, trip_end)
    return (
        f"Email: {email}\n\n"
        f"Subject: Athletic absences - {course_name}\n\n"
        f"{greeting}\n\n"
        f"This email is to notify you that the following students will be absent from "
        f"{course_name} from {date_range} for their sport. "
        f"Their attendance should already be updated.\n\n"
        f"{names}\n\n"
        f"Thank you!"
    )


# ---------------------------------------------------------------------------
# Loading the roster
# ---------------------------------------------------------------------------

def build_players(csv_rows, term_id):
    """Turn roster CSV rows into Player objects with their term schedules filled in.

    Makes the Populi API calls (through StudentIDtoSchedule), so this is the
    slow step on the first run. Later runs are faster because meeting times
    are cached on disk.
    """
    players = []
    for row in csv_rows:
        # Column names must match the CSV header exactly
        p = Player(
            populi_id=row["Person Populi ID"],
            first=row.get("First Name", ""),
            last=row.get("Last Name", ""),
        )
        # Fetch this student's schedule and convert each meeting into a Course
        for c in PopuliPuller.GetTermScheduleFromPopuliID(p.populi_id, term_id):
            p.courses.append(Course.from_raw(c["name"], c["offering_id"], c["days"], c["start"], c["end"]))
        players.append(p)
        # Progress line so a long run visibly makes progress (prints the whole CSV row as well as the name)
        print(f"[{row}] {p.first} {p.last}: {len(p.courses)} courses", flush=True)
    return players

def getPeople():
    """Read the baseball roster CSV and return a list of Players with schedules loaded."""
    # utf-8-sig strips the invisible marker some exports add to the start of the file
    with open("RosterCSVs/Baseball.csv", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)   # each row becomes a dict keyed by the column headers
        players = build_players(reader, TERM_ID)

    return players


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Load the roster and every player's schedule
    players = getPeople()

    # The trip window: when the team leaves and when they get back
    trip_start = datetime(2026, 10, 23, 13, 0)    # Fri Oct 23, 1:00 PM
    trip_end = datetime(2026, 10, 25, 20, 0)      # Sun Oct 25, 8:00 PM

    # Find every course section where at least one player will miss class
    conflicts = find_conflicts(players, trip_start, trip_end)
    missing = []   # sections where Populi has no professor email
    # Print one email draft per section, in alphabetical order by course name
    for offering_id, info in sorted(conflicts.items(), key=lambda kv: kv[1]["name"]):
        faculty = PopuliPuller.get_faculty(offering_id, TERM_ID)
        if not faculty.get("email"):
            missing.append(info["name"])
        print(build_email(info["name"], info["students"], trip_start, trip_end, faculty))
        print("\n" + "-" * 40 + "\n")
    # Summary of emails that need to be sent manually
    if missing:
        print("No emails found for:", ", ".join(missing))

if __name__ == "__main__": main()
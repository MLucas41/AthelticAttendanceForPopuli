"""
PopuliPuller.py

Helper functions for pulling course schedule and instructor data out of Populi's
API (v2) for a given student and term.

Main entry points:
    getTermID(termName)                       -> Populi's numeric ID for a term name
    GetTermScheduleFromPopuliID(pid, term_id) -> list of meeting dicts for a student
    get_faculty(offering_id, term_id)         -> primary instructor info for a course section

Terminology:
    populi_id    - Populi's internal person ID (not the school's student ID number)
    term_id      - Populi's internal ID for an academic term (e.g. Fall 2026)
    offering_id  - ID of one section of a course in one term
"""

import requests
import os
import time
import json
from dotenv import load_dotenv
from functools import lru_cache
from pathlib import Path

load_dotenv()  # Load the API key from the .env file into environment variables

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#lukeStudentID= "000020650"
#lukePersonID = 0
#termID = 0
#schedule = ''
#enrollments = ''

url = "https://corban.populiweb.com/api2/"  # Corban's Populi API base URL
api_key = os.getenv("API_KEY")              # Read from .env so the key never lives in the code

# Cache of course meeting times, saved next to this script so it persists between runs
CACHE_FILE = Path(__file__).with_name("course_meetings.json")

# In-memory lookup of instructor info, filled while the term's course list is downloaded.
# Shape: {offering_id: {"name": ..., "last_name": ..., "email": ..., "other_emails": ...}}
_faculty = {}

# Every request sends the API key in this header
headers = {
    "Authorization": f'Bearer {api_key}'
}

# ---------------------------------------------------------------------------
# Meeting-time cache (saved to disk)
# ---------------------------------------------------------------------------

def _load_cache():
    """Load the saved meeting-times cache from disk.

    Returns an empty dict if the file doesn't exist yet or is unreadable,
    so the script just rebuilds the cache from the API.
    """
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            print("course_meetings.json was corrupt, starting fresh")
    return {}

# Loaded once when the module is imported.
# Shape: {"<offering_id as string>": [meeting dict, ...]}
_meetings_cache = _load_cache()

def _save_cache():
    """Write the meeting-times cache to disk.

    Writes to a temp file first and then swaps it in, so pressing Ctrl+C
    in the middle of a save can't leave a half-written (corrupt) cache file.
    """
    tmp = CACHE_FILE.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_meetings_cache, f)
    tmp.replace(CACHE_FILE)


# ---------------------------------------------------------------------------
# API helper
# ---------------------------------------------------------------------------

def _get(endpoint, **kwargs):
    """Make a GET request to Populi and return the parsed JSON.

    All API calls go through here so they share the same timeout,
    error handling, and delay between requests.

    endpoint: path after the base URL, e.g. "people/123/enrollments"
    kwargs:   passed to requests.get (Populi takes parameters as a JSON body,
              so callers use json={...})
    Raises requests.HTTPError on a bad status code.
    """
    response = requests.get(url + endpoint, headers=headers, timeout=30, **kwargs)
    response.raise_for_status()
    time.sleep(1.3)  # small pause between calls to stay under Populi's rate limit
    return response.json() if response.content else {}


# ---------------------------------------------------------------------------
# Term course list and instructors
# ---------------------------------------------------------------------------

@lru_cache(maxsize = None)
def _get_term_offering_ids(term_id):
    """Return the IDs of every course section offered in a term.

    Populi returns course offerings in pages of up to 200, so this loops until
    there are no more pages. While reading each course, it also saves that
    course's primary instructor info into the _faculty dict, so instructor
    emails come at no extra API cost.

    Cached with lru_cache, so the download happens once per run per term.
    Returns a frozenset (immutable, and fast to check membership in).
    """
    ids = set()
    page = 1

    while True:
        result = _get(f"academicterms/{term_id}/courseofferings", json={"page": page})

        for c in result.get("data", []):
            ids.add(c["id"])
            # report_data holds summary fields, including the faculty details
            rd = c.get("report_data") or {}
            _faculty[c["id"]] = {
                "name": rd.get("primary_faculty_display_name"),
                "last_name": rd.get("primary_faculty_last_name"),
                "email": rd.get("primary_faculty_email_address"),
                "other_emails": rd.get("other_faculty_email_addresses")
            }
        if not result.get("has_more", False):  # last page reached
            return frozenset(ids)
        page += 1

def get_faculty(offering_id, term_id):
    """Return the primary instructor info for one course section.

    Returns a dict with name, last_name, email, and other_emails.
    Any of those values can be None, and the dict is empty if the
    offering isn't found, so callers should check before using them.
    """
    _get_term_offering_ids(term_id)  # makes sure the _faculty dict has been filled
    return _faculty.get(offering_id, {})


# ---------------------------------------------------------------------------
# Course meeting times
# ---------------------------------------------------------------------------

def _get_course_meetings(offering_id):
    """Return the meeting times for one course section.

    The result is a list with one dict per meeting time:
        {"offering_id", "name", "days", "start", "end"}
    Courses with no scheduled meetings (internships, advising, etc.) return a
    single entry where days/start/end are None.

    Checks the on-disk cache first, and only calls the API for courses it
    hasn't seen. The cache is saved after each new course, so an interrupted
    run keeps its progress.
    """
    key = str(offering_id)  # JSON file keys are always strings
    if key not in _meetings_cache:
        # Two requests per new course: one for the name, one for the meeting times
        name = _get(f"courseofferings/{offering_id}")["catalog_courses"][0]["name"]
        meetings = _get(f"courseofferings/{offering_id}/meetingtimes").get("data", [])
        if not meetings:
            entries = [{"offering_id": offering_id, "name": name,
                        "days": None, "start": None, "end": None}]
        else:
            entries = [
                {"offering_id": offering_id, "name": name, "days": m["weekdays"],
                 "start": m["start_time"], "end": m["end_time"]}
                for m in meetings
            ]
        _meetings_cache[key] = entries
        _save_cache()  # saved as you go, so an interrupted run keeps its progress
    return _meetings_cache[key]


# ---------------------------------------------------------------------------
# Student schedule
# ---------------------------------------------------------------------------

def _get_student_offering_ids(populi_id, term_id):
    """Return the set of course-section IDs a student is enrolled in for one term.

    Populi's enrollments endpoint returns the student's entire history, so each
    enrollment is checked against the term's course list to keep only this term.
    """
    enrollments = _get(f"people/{populi_id}/enrollments").get("data", [])
    term_ids = _get_term_offering_ids(term_id)
    return {e["course_offering_id"] for e in enrollments if e["course_offering_id"] in term_ids}

def GetTermScheduleFromPopuliID(populi_id, term_id):
    """Return a student's full schedule for a term as a list of meeting dicts.

    This is the main function other scripts call. Each dict has:
        offering_id, name, days (e.g. ['mo','we','fr']), start, end (e.g. '10:30AM')
    A course meeting at two different times appears as two entries.
    """
    schedule = []
    for offering_id in _get_student_offering_ids(populi_id, term_id):
        schedule.extend(_get_course_meetings(offering_id))
    return schedule


# ---------------------------------------------------------------------------
# Term lookup
# ---------------------------------------------------------------------------

#Returns the term ID given the term name
def getTermID(termName):
    """Look up Populi's numeric term ID from a term name like "Fall 2026".

    Prints the ID when found. Returns None if no term has that exact name
    (the match is case-sensitive), and returns the exception object if the
    request fails.
    """

    if(termName == "Current Term"):
        response = requests.get(f"{url}academicterms/current", headers=headers)
        time.sleep(1.3)
        response.raise_for_status()
        data = response.json()
        return(data["id"])
    else:
        response = requests.get(f"{url}academicTerms/", headers=headers)
        time.sleep(1.3)
        response.raise_for_status()
        data = response.json()

    print("\n=======================\nGetting the Term ID...")

    #Loops through the response under the data key to find the term that matches the 
    # termName parameter and returns the corresponding term ID.
    for year in range(len(data["data"])):
        if(data["data"][year]["name"] == termName):
            print(f"Found {termName} term ID:", data["data"][year]["id"])
            print("=======================")
            return data["data"][year]["id"]
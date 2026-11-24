"""
RosterDownloader.py

Downloads each athletic team's roster from Populi and replaces the matching
file in RosterCSVs/ (Baseball -> RosterCSVs/Baseball.csv).

Run it from the project folder:

    uv run RosterDownloader.py              # every team in TEAMS
    uv run RosterDownloader.py Baseball     # just the teams you name

Teams are found with Populi's people filter on the "athletic team" custom
field. To add a team, build the same search in Populi for that team, copy
the "choice" number from its filter, and add a line to TEAMS below.

A team's old CSV is only replaced if the download finishes and finds at least
one person, so a failed or empty search never wipes out a good roster.
"""

import csv
import sys
from pathlib import Path

import PopuliPuller

ROSTER_DIR = Path("RosterCSVs")

# The Populi custom field that holds a person's athletic team
TEAM_FIELD_ID = "200018"

# Team name (becomes the file name) -> that team's choice ID in the custom field
TEAMS = {
    "Baseball":                 "434648",
    "Basketball - Men's":       "434654",
    "Basketball - Women's":     "434660",
    "Beach Volleyball":         "434666",
    "Cross Country - Men's":    "434672",
    "Cross Country - Women's":  "434678",
    "Golf - Men's":             "434684",
    "Golf - Women's":           "434690",
    "Soccer - Men's":           "434708",
    "Soccer - Women's":         "434714",
    "Softball":                 "434720",
    "Track - Men's":            "434726",
    "Track - Women's":          "434732",
    "Volleyball - Women's":     "438518",
    "Wrestling - Men's":        "434744",
}

# The app reads these exact column names (see build_players in ConflictEmailGenerator.py)
CSV_COLUMNS = ["Person Populi ID", "First Name", "Last Name"]


def team_filter(choice_id):
    """The people filter for one team, in the same shape as Populi's own search."""
    return {
        "0": {
            "logic": "ALL",
            "fields": [
                {
                    "name": "custom_field",
                    "value": {"custom_info_field_id": TEAM_FIELD_ID, "choice": choice_id},
                    "positive": 1,
                }
            ],
        }
    }


def download_team(choice_id):
    """Return [{"id", "first", "last"}, ...] for everyone on one team.

    Populi returns people in pages, so this keeps asking until has_more is false.
    """
    people, page = [], 1
    while True:
        result = PopuliPuller._get("people", json={"filter": team_filter(choice_id), "page": page})
        for p in result.get("data", []):
            people.append({
                "id": str(p["id"]),
                "first": p.get("first_name") or "",
                "last": p.get("last_name") or "",
            })
        if not result.get("has_more", False):
            break
        page += 1
    people.sort(key=lambda p: (p["last"].lower(), p["first"].lower()))
    return people


def save_roster(team, people):
    """Write the team's CSV, swapping it in only once it's fully written."""
    ROSTER_DIR.mkdir(exist_ok=True)
    final = ROSTER_DIR / f"{team}.csv"
    tmp = final.with_suffix(".tmp")
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)
        for p in people:
            writer.writerow([p["id"], p["first"], p["last"]])
    tmp.replace(final)
    return final


def update_rosters(team_names=None):
    """Download and replace rosters. Returns {team: count or error message}."""
    names = team_names or list(TEAMS)
    unknown = [t for t in names if t not in TEAMS]
    if unknown:
        raise ValueError(f"Unknown team(s): {', '.join(unknown)}. Known: {', '.join(TEAMS)}")

    results = {}
    for team in names:
        try:
            people = download_team(TEAMS[team])
        except Exception as e:
            results[team] = f"FAILED, kept the old file ({e})"
            continue
        if not people:
            results[team] = "no one found, kept the old file"
            continue
        try:
            save_roster(team, people)
        except PermissionError:
            # Windows won't replace a file another program has open (usually Excel)
            ROSTER_DIR.joinpath(f"{team}.tmp").unlink(missing_ok=True)
            results[team] = (f"could not replace {ROSTER_DIR / (team + '.csv')} - "
                             f"close it in Excel (or anything else that has it open) and run again")
            continue
        results[team] = len(people)
    return results


if __name__ == "__main__":
    for team, outcome in update_rosters(sys.argv[1:]).items():
        if isinstance(outcome, int):
            print(f"{team}: saved {outcome} players to {ROSTER_DIR / (team + '.csv')}")
        else:
            print(f"{team}: {outcome}")
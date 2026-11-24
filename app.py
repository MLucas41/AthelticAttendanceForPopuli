"""
app.py

Window for generating professor emails about athletes' trip absences.

Two ways to say who is traveling and when:
  1. Load a Microsoft Forms export (.xlsx) by dropping it on the box or
     clicking Browse. The players and dates come from the file.
  2. Leave the file empty, pick a team, and type the dates by hand.

Run with:  uv run app.py
"""

import os
import threading
import tkinter as tk
import webbrowser
import json
import subprocess
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote

import ConflictEmailGenerator as core
import FormImport

_Base = tk.Tk

FMT = "%Y-%m-%d %H:%M"   # manual date entry, 24-hour, e.g. 2026-10-23 13:00


class App(_Base):
    def __init__(self):
        super().__init__()
        ensure_outlook_running()
        self.title("Athletic Absence Emails")
        self.drafts = []
        self.trips = []      # every response found in the loaded file
        self.trip = None     # the response currently selected (None = manual mode)
        self.players = []    # Player objects (with schedules) from the last Generate
        self.window = None   # (start, end) that the last Generate used
        self._players_win = None
        self.sendtoProfeesors = tk.BooleanVar(master=self, value=False)
        self.sendToLuke = tk.BooleanVar(master=self, value=False)

        self._build_file_box()        # row 0
        self._build_manual_form()     # row 1

        self.status = ttk.Label(self, text="Load a trip file, or pick a team and dates.", padding=(10, 0))
        self.status.grid(row=2, column=0, sticky="w")

        body = ttk.Frame(self, padding=10)
        body.grid(row=3, column=0, sticky="nsew")
        self.listbox = tk.Listbox(body, width=38, height=18, exportselection=False)
        self.listbox.grid(row=0, column=0, sticky="ns")
        self.listbox.bind("<<ListboxSelect>>", self.show_draft)
        self.preview = tk.Text(body, width=70, height=18, wrap="word")
        self.preview.grid(row=0, column=1, padx=(10, 0))

        self._build_action_bar()      # row 4

    # --- layout ------------------------------------------------------------

    def _build_file_box(self):
        box = ttk.LabelFrame(self, text="Trip request (Microsoft Forms export)", padding=8)
        box.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        box.columnconfigure(0, weight=1)

        hint = "Click Browse to pick the .xlsx file"
        self.drop = tk.Label(box, text=hint, relief="groove", height=3)
        self.drop.grid(row=0, column=0, columnspan=3, sticky="ew")

        ttk.Button(box, text="Browse...", command=self.browse).grid(row=1, column=1, padx=5, pady=(6, 0))
        ttk.Button(box, text="Clear", command=self.clear_trip).grid(row=1, column=2, pady=(6, 0))

        self.trip_pick = ttk.Combobox(box, state="disabled", width=60)
        self.trip_pick.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.trip_pick.bind("<<ComboboxSelected>>", self._trip_selected)

        self.trip_info = ttk.Label(box, text="", wraplength=520, justify="left")
        self.trip_info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def _build_manual_form(self):
        form = ttk.LabelFrame(self, text="Or enter a trip by hand", padding=8)
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        ttk.Label(form, text="Team").grid(row=0, column=0, sticky="w")
        self.team = ttk.Combobox(form, values=core.list_teams(), state="readonly", width=22)
        self.team.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(form, text="Leaves (YYYY-MM-DD HH:MM)").grid(row=1, column=0, sticky="w")
        self.start = ttk.Entry(form, width=24)
        self.start.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(form, text="Returns (YYYY-MM-DD HH:MM)").grid(row=2, column=0, sticky="w")
        self.end = ttk.Entry(form, width=24)
        self.end.grid(row=2, column=1, padx=5, pady=2)

        self.go = ttk.Button(form, text="Generate drafts", command=self.generate)
        self.go.grid(row=3, column=1, sticky="e", pady=5)

    def _build_action_bar(self):
        bar = ttk.Frame(self, padding=(10, 0, 10, 10))
        bar.grid(row=4, column=0, sticky="ew")
        bar.columnconfigure(0, weight=1)

        # The two send options. Both start unchecked (drafts only).
        checks = ttk.Frame(bar)
        checks.grid(row=0, column=0, sticky="w")
        self.luke_check = ttk.Checkbutton(checks, text="Send Emails to Luke?", variable=self.sendToLuke)
        self.luke_check.grid(row=0, column=0, padx=(0, 15))
        self.prof_check = ttk.Checkbutton(checks, text="Send Emails to Profs?", variable=self.sendtoProfeesors)
        self.prof_check.grid(row=0, column=1)
        for cb in (self.luke_check, self.prof_check):
            cb.state(["!selected", "!alternate"])

        buttons = ttk.Frame(bar)
        buttons.grid(row=0, column=1, sticky="e")
        self.players_btn = ttk.Button(buttons, text="View players", command=self.show_players)
        self.players_btn.grid(row=0, column=0, padx=(0, 5))
        self.open_btn = ttk.Button(buttons, text="Open selected in mail app", command=self.open_mail)
        self.open_btn.grid(row=0, column=1, padx=(0, 5))
        # Each button gets its own name so create_drafts disables the right one
        self.draft_btn = ttk.Button(
            buttons, text="Send Emails / Create all drafts in Outlook", command=self.create_drafts
        )
        self.draft_btn.grid(row=0, column=2)

        tests = ttk.Frame(bar)
        tests.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.test_save_btn = ttk.Button(tests, text="Test draft saving", command=self._test_draft_saving)
        self.test_save_btn.grid(row=0, column=0, padx=(0, 5))
        self.test_count_btn = ttk.Button(tests, text="Test draft count", command=core._testDraftCount)
        self.test_count_btn.grid(row=0, column=1)

    # --- loading the trip file --------------------------------------------

    def _on_drop(self, event):
        # event.data wraps paths containing spaces in {braces}; splitlist unwraps them
        paths = self.tk.splitlist(event.data)
        if paths:
            self.load_file(paths[0])

    def browse(self):
        path = filedialog.askopenfilename(
            title="Select the Forms export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.load_file(path)

    def load_file(self, path):
        if not path.lower().endswith(".xlsx"):
            messagebox.showerror("Wrong file type", "Please use the .xlsx file exported from Microsoft Forms.")
            return
        try:
            trips = FormImport.load_trips(path)
        except Exception as e:
            messagebox.showerror("Could not read the file", str(e))
            return
        if not trips:
            messagebox.showwarning("Empty file", "That file has no responses in it.")
            return

        self.trips = trips
        self.trip_pick.config(state="readonly", values=[t.label for t in trips])
        self.trip_pick.current(len(trips) - 1)    # newest response is last
        self.drop.config(text=os.path.basename(path))
        self._set_manual_state(False)
        self._trip_selected()

    def _trip_selected(self, _event=None):
        self.trip = self.trips[self.trip_pick.current()]
        self.players, self.window = [], None      # schedules loaded earlier belong to a different trip
        t = self.trip
        info = (
            f"{len(t.players)} players, {t.start.strftime('%a %b %d')} through "
            f"{t.end.strftime('%a %b %d')} (whole days). Click Generate drafts."
        )
        if t.warnings:
            info += "\nWarnings: " + "; ".join(t.warnings)
        self.trip_info.config(text=info)
        self.status.config(text="Trip loaded.")

    def clear_trip(self):
        self.trips, self.trip = [], None
        self.players, self.window = [], None
        self.trip_pick.config(state="disabled", values=[])
        self.trip_pick.set("")
        self.trip_info.config(text="")
        self.drop.config(text="Click Browse to pick the .xlsx file")
        self._set_manual_state(True)

    def _set_manual_state(self, enabled):
        self.team.config(state="readonly" if enabled else "disabled")
        self.start.config(state="normal" if enabled else "disabled")
        self.end.config(state="normal" if enabled else "disabled")

    # --- actions -----------------------------------------------------------

    def _open_classic_outlook(self):

        core._openApp("Outlook.Application")

        return

    def _test_draft_saving(self):
        if messagebox.askyesno(
            "Test Drafts",
            "Do you want to test whether drafts can save?"
        ):
            core.subsequentAccessTest()
            core._test_draft_saving()
            return
        return

    def create_drafts(self):

        with_email = sum(1 for d in self.drafts if d["to"])

        if(self.sendToLuke.get()):
            s=f"This will send {with_email} email to Luke\nDo you wish to proceed?"
            msg = f"Sent {with_email} emails to Luke."
        else:
            if(self.sendtoProfeesors.get()):
                s=f"This will send {with_email} emails to various professors\nAre you aboslutely sure you wish to proceed?"
                msg = f"Sent {with_email} emails to Professors."
            else:
                s=f"This will create {with_email} drafts in outlook"

        if not self.drafts:
            return

        if not messagebox.askyesno(
            "Send / Create",
            s + f"\n{len(self.drafts) - with_email} have no professor email and will be skipped."
        ):
            return
        try:
            created, skipped = core.create_outlook_drafts(self.drafts, send=self.sendtoProfeesors.get(), sendtoLuke=self.sendToLuke.get())
        except Exception as e:
            messagebox.showerror("Outlook error", str(e))
            return
        self.draft_btn.config(state="disabled")   # prevents accidental duplicates; re-enabled on next Generate
        if(self.sendToLuke.get() == False and self.sendtoProfeesors.get() == False):
            msg = f"Created {created} drafts in Outlook."
        if skipped:
            msg += "\n\nNo email found for:\n" + "\n".join(skipped)
        messagebox.showinfo("Done", msg)

    # --- generating drafts ------------------------------------------------

    def generate(self):
        if self.trip:
            trip = self.trip

            def job():
                drafts = core.generate_drafts_for_trip(trip)
                # Schedules are cached by now, so this second call is instant
                return drafts, core.load_players_from_trip(trip), (trip.start, trip.end)

            msg = f"Loading schedules for {len(trip.players)} players from Populi..."
        else:
            try:
                start = datetime.strptime(self.start.get().strip(), FMT)
                end = datetime.strptime(self.end.get().strip(), FMT)
            except ValueError:
                messagebox.showerror("Bad date", f"Use the format {FMT}, e.g. 2026-10-23 13:00")
                return
            if end <= start:
                messagebox.showerror("Bad range", "The return time must be after the departure time.")
                return
            team = self.team.get()
            if not team:
                messagebox.showerror("No team", "Pick a team, or load a trip file.")
                return
            def job():
                drafts = core.generate_drafts_for_players(team, start, end)
                return drafts, core.load_players(team), (start, end)

            msg = "Loading schedules from Populi (the first run for a team is slow)..."

        self.players, self.window = [], None      # don't show old results if this run fails
        self.go.config(state="disabled")
        self.status.config(text=msg)
        # Run the slow Populi calls off the UI thread so the window doesn't freeze
        threading.Thread(target=self._work, args=(job,), daemon=True).start()

    def _work(self, job):
        try:
            self.after(0, self._done, job(), None)
        except Exception as e:
            self.after(0, self._done, None, e)

    def _done(self, result, error):
        self.go.config(state="normal")
        if error:
            self.status.config(text="Failed.")
            messagebox.showerror("Error", str(error))
            return
        drafts, self.players, self.window = result
        self.drafts = drafts
        self.draft_btn.config(state="normal")     # a fresh set of drafts can be created again
        self.listbox.delete(0, tk.END)
        for d in drafts:
            flag = "" if d["to"] else "  [NO EMAIL]"
            self.listbox.insert(tk.END, d["course"] + flag)
        self.preview.delete("1.0", tk.END)
        self.status.config(text=f"{len(drafts)} email(s) to review.")

    # --- reviewing ---------------------------------------------------------

    def show_players(self):
        if self.players:        # drafts were generated, so schedules are loaded
            entries = [{"id": p.populi_id, "first": p.first, "last": p.last, "player": p}
                       for p in self.players]
            window = self.window
        elif self.trip:         # file loaded but not generated yet: names and IDs only
            entries = [{**e, "player": None} for e in self.trip.players]
            window = None
        else:
            messagebox.showinfo("No players yet", "Load a trip file, or generate drafts first.")
            return
        if self._players_win is not None and self._players_win.winfo_exists():
            self._players_win.destroy()               # one players window at a time
        self._players_win = PlayersWindow(self, entries, window)

    def show_draft(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        d = self.drafts[sel[0]]
        self.preview.delete("1.0", tk.END)
        self.preview.insert(
            tk.END, f"To: {d['to'] or 'NO EMAIL FOUND'}\nSubject: {d['subject']}\n\n{d['body']}"
        )

    def open_mail(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        d = self.drafts[sel[0]]
        if not d["to"]:
            messagebox.showwarning("No email", "Populi has no email for this professor. Send it manually.")
            return
        core.openSelectedInMailApp(d['to'], d['subject'], d['body'])
        #webbrowser.open(f"mailto:{d['to']}?subject={quote(d['subject'])}&body={quote(d['body'])}")


def missed_sessions(player, window):
    """(start, course name) for every class the player misses in the (start, end) window.

    Skips the same courses as the emails do (Chapel), so the counts here match them.
    """
    if not window:
        return []
    return [(s, name) for s, name, _end in player.missed_classes(*window)
            if name.lower() not in core.SKIP_COURSES]


def unscheduled_courses(player):
    """Courses with no meeting time. They can never appear as a conflict, so flag them for a manual check."""
    return sorted({c.name for c in player.courses if not c.has_meetings})


class PlayersWindow(tk.Toplevel):
    """The traveling players. Once drafts are generated, also each player's missed classes."""

    def __init__(self, parent, entries, window):
        # entries: [{"id", "first", "last", "player"}], where "player" is a Player or None
        super().__init__(parent)
        self.title("Players")
        self.entries = sorted(entries, key=lambda e: (e["last"].lower(), e["first"].lower()))
        self.window = window

        if window:
            head = f"{len(entries)} players - classes missed {window[0]:%a %b %d} to {window[1]:%a %b %d}"
        else:
            head = f"{len(entries)} players - generate drafts to see their missed classes"
        ttk.Label(self, text=head, padding=10).grid(row=0, column=0, sticky="w")

        self.tree = ttk.Treeview(self, columns=("name", "id", "missed"), show="headings", height=12)
        for col, text, width in (("name", "Name", 220), ("id", "Populi ID", 90), ("missed", "Classes missed", 110)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=10)
        self.tree.bind("<<TreeviewSelect>>", self._show_detail)

        self.detail = tk.Text(self, width=60, height=10, wrap="word")
        self.detail.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        for i, e in enumerate(self.entries):
            p = e["player"]
            missed = len(missed_sessions(p, window)) if p and window else "-"
            self.tree.insert("", "end", iid=str(i), values=(f'{e["last"]}, {e["first"]}', e["id"], missed))

    def _show_detail(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        e = self.entries[int(sel[0])]
        p = e["player"]
        self.detail.delete("1.0", tk.END)
        if not p or not self.window:
            self.detail.insert(tk.END, "Generate drafts to load this player's schedule.")
            return
        lines = [f"{s:%a %b %d  %I:%M%p}  {name}" for s, name in missed_sessions(p, self.window)]
        lines = lines or ["No classes missed."]
        extra = unscheduled_courses(p)
        if extra:
            lines += ["", "No meeting time listed (check manually):"] + [f"  {n}" for n in extra]
        self.detail.insert(tk.END, "\n".join(lines))


def ensure_outlook_running():
    """Start classic Outlook if it isn't already open."""
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"],
        capture_output=True, text=True,
    )
    if "OUTLOOK.EXE" not in result.stdout.upper():
        # "start" finds outlook.exe through Windows' registered app paths
        subprocess.Popen(["cmd", "/c", "start", "", "outlook.exe"])


if __name__ == "__main__":
    App().mainloop()
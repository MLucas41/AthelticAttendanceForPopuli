import threading
import tkinter as tk
import webbrowser
import json
import subprocess
from datetime import datetime
from tkinter import ttk, messagebox
from urllib.parse import quote

import ConflictEmailGenerator as core

FMT = "%Y-%m-%d %H:%M"   # 24-hour, e.g. 2026-10-23 13:00

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        ensure_outlook_running()
        self.title("Athletic Absence Emails")
        self.drafts = []
        self.sendtoProfeesors = tk.BooleanVar(master=self, value=False)
        self.sendToLuke = tk.BooleanVar(master=self, value=False)

        form = ttk.Frame(self, padding=10)
        form.grid(row=0, column=0, sticky="ew")

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

        self.checkbox = ttk.Checkbutton(self, text="Send Emails to Profs?", variable=self.sendtoProfeesors)
        self.checkbox.place(x=440, y=499)
        self.checkbox.state(["!selected", "!alternate"])
        self.sendtoProfeesors.set(False)

        self.checkbox = ttk.Checkbutton(self, text="Send Emails to Luke?", variable=self.sendToLuke)
        self.checkbox.place(x=300, y=499)
        self.checkbox.state(["!selected", "!alternate"])
        self.sendtoProfeesors.set(False)

        self.draft_btn = ttk.Button(self, text="Send Emails / Create all drafts in Outlook", command=self.create_drafts)
        self.draft_btn.grid(row=4, column=0, sticky="e", padx=10, pady=(0, 10))

        self.draft_btn = ttk.Button(self, text="Test draft saving", command=self._test_draft_saving)
        self.draft_btn.grid(row=5, column=0, sticky="e", padx=10, pady=(0, 10))

        self.draft_btn = ttk.Button(self, text="Test draft count", command=core._testDraftCount)
        self.draft_btn.grid(row=6, column=0, sticky="e", padx=10, pady=(0, 10))

        self.status = ttk.Label(self, text="Pick a team and dates.", padding=(10, 0))
        self.status.grid(row=1, column=0, sticky="w")

        body = ttk.Frame(self, padding=10)
        body.grid(row=2, column=0, sticky="nsew")
        self.listbox = tk.Listbox(body, width=38, height=18, exportselection=False)
        self.listbox.grid(row=0, column=0, sticky="ns")
        self.listbox.bind("<<ListboxSelect>>", self.show_draft)
        self.preview = tk.Text(body, width=70, height=18, wrap="word")
        self.preview.grid(row=0, column=1, padx=(10, 0))

        self.open_btn = ttk.Button(self, text="Open selected in mail app", command=self.open_mail)
        self.open_btn.grid(row=3, column=0, sticky="e", padx=10, pady=(0, 10))

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

    def generate(self):
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
            messagebox.showerror("No team", "Pick a team first.")
            return

        self.go.config(state="disabled")
        self.status.config(text="Loading schedules from Populi (the first run for a team is slow)...")
        # Run the slow Populi calls off the UI thread so the window doesn't freeze
        threading.Thread(target=self._work, args=(team, start, end), daemon=True).start()

    def _work(self, team, start, end):
        try:
            drafts = core.generate_drafts(team, start, end)
            self.after(0, self._done, drafts, None)
        except Exception as e:
            self.after(0, self._done, [], e)

    def _done(self, drafts, error):
        self.go.config(state="normal")
        if error:
            self.status.config(text="Failed.")
            messagebox.showerror("Error", str(error))
            return
        self.drafts = drafts
        self.listbox.delete(0, tk.END)
        for d in drafts:
            flag = "" if d["to"] else "  [NO EMAIL]"
            self.listbox.insert(tk.END, d["course"] + flag)
        self.preview.delete("1.0", tk.END)
        self.status.config(text=f"{len(drafts)} course(s) with conflicts.")

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
        webbrowser.open(f"mailto:{d['to']}?subject={quote(d['subject'])}&body={quote(d['body'])}")

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
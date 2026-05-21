"""
Optihome Payroll Processing — Tkinter GUI

Provides file pickers for Notion, Turno, timeclock, and output Excel files,
then runs the export-timesheet process and displays results.
"""

import calendar
import importlib
import json
import os
import re
import sys
import threading
import subprocess
import tkinter as tk
from datetime import date, timedelta
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Import the processing function from the hyphenated module name
# ---------------------------------------------------------------------------

def _get_script_dir():
    """Return the directory containing the running code.

    Non-frozen: directory of this .py file.
    Frozen (.app): the _MEIPASS temp dir (for bundled data files).
    """
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _get_project_dir():
    """Best-effort guess at the project root (contains Raw/, Timesheets/, timesheet-rates.csv).

    Walks upward from the script location (or .app bundle) looking for timesheet-rates.csv.
    """
    if not getattr(sys, "frozen", False):
        start = os.path.dirname(os.path.abspath(__file__))
    else:
        # sys.executable is …/Foo.app/Contents/MacOS/Foo
        app_dir = os.path.dirname(sys.executable)      # Contents/MacOS
        app_dir = os.path.dirname(app_dir)             # Contents
        app_dir = os.path.dirname(app_dir)             # Foo.app
        start = os.path.dirname(app_dir)               # dir containing .app

    candidate = start
    for _ in range(4):
        if os.path.isfile(os.path.join(candidate, "timesheet-rates.csv")):
            return candidate
        parent = os.path.dirname(candidate)
        if parent == candidate:
            break
        candidate = parent

    return start

def _import_export_module():
    """Import export-timesheet.py (hyphenated name requires importlib)."""
    # When frozen, the file is in _MEIPASS; otherwise next to this script
    script_dir = _get_script_dir()
    module_path = os.path.join(script_dir, "export-timesheet.py")
    if not os.path.exists(module_path):
        # Fallback: check project dir
        module_path = os.path.join(_get_project_dir(), "export-timesheet.py")
    spec = importlib.util.spec_from_file_location("export_timesheet", module_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

export_mod = _import_export_module()
process_timesheet = export_mod.process_timesheet

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

DAYS_OF_WEEK = list(calendar.day_name)  # ['Monday', ..., 'Sunday']
_CONFIG_FILE = os.path.expanduser("~/.optihome_payroll_config.json")
REPORT_VISIBILITY_DEFAULTS = {
    "notion": True,
    "turno": True,
    "expenses": True,
    "time": False,
}
RUN_BUTTON_BG = "#00897b"
RUN_BUTTON_ACTIVE_BG = "#00695c"
RUN_BUTTON_DISABLED_BG = "#607d8b"
RUN_BUTTON_FG = "#0b2f2a"
RUN_BUTTON_DISABLED_FG = "#263238"


def _load_config():
    try:
        with open(_CONFIG_FILE) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def _save_config(data):
    try:
        with open(_CONFIG_FILE, "w") as f:
            json.dump(data, f)
    except OSError:
        pass


def _get_period_end_date(day_name):
    """Most recent date (including today) that falls on day_name."""
    target_wd = DAYS_OF_WEEK.index(day_name)
    today = date.today()
    days_back = (today.weekday() - target_wd) % 7
    return today - timedelta(days=days_back)


def _shorten_path(path, segments=3):
    """Return the last *segments* components of a path, prefixed with /."""
    if not path:
        return ""
    parts = path.replace("\\", "/").rstrip("/").split("/")
    parts = [p for p in parts if p]
    if len(parts) <= segments:
        return path
    return "/" + "/".join(parts[-segments:])


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#fffde7",
            foreground="#222222",
            relief="solid",
            borderwidth=1,
            wraplength=300,
            padx=8,
            pady=5,
        )
        label.pack()

    def hide(self, _event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


def label_with_tip(parent, text, tip, bg=None, fg=None):
    frame_opts = {}
    if bg is not None:
        frame_opts["bg"] = bg
    frame = tk.Frame(parent, **frame_opts)
    label_opts = {"text": text, "anchor": "w"}
    if bg is not None:
        label_opts["bg"] = bg
    if fg is not None:
        label_opts["fg"] = fg
    tk.Label(frame, **label_opts).pack(side="left")
    create_help_icon(frame, tip, bg=bg).pack(side="left", padx=(7, 0))
    return frame


def create_help_icon(parent, tip, bg=None):
    icon_opts = {
        "text": "i",
        "fg": "white",
        "bg": "#1976d2",
        "activeforeground": "white",
        "activebackground": "#1565c0",
        "cursor": "question_arrow",
        "font": ("Helvetica", 9, "bold"),
        "width": 2,
        "height": 1,
        "relief": "solid",
        "borderwidth": 1,
        "highlightthickness": 1,
        "highlightbackground": "#90caf9",
        "highlightcolor": "#90caf9",
    }
    icon = tk.Label(parent, **icon_opts)
    Tooltip(icon, tip)
    return icon

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class PayrollApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Optihome Payroll Processing")
        self.resizable(False, False)

        self._project_dir = _get_project_dir()
        self._raw_dir = os.path.join(self._project_dir, "Raw")
        self._timesheets_dir = os.path.join(self._project_dir, "Timesheets")

        # Ensure default directories exist for the file dialogs
        os.makedirs(self._raw_dir, exist_ok=True)
        os.makedirs(self._timesheets_dir, exist_ok=True)

        # Full paths stored separately from the display variables
        self._time_path = ""
        self._turno_path = ""
        self._notion_path = ""
        self._expenses_path = ""
        self._output_path = ""
        self._rates_path = os.path.join(self._project_dir, "timesheet-rates.csv")

        config = _load_config()
        self._end_day_var = tk.StringVar(value=config.get("end_day", "Wednesday"))
        self._report_visible_vars = {
            key: tk.BooleanVar(value=bool(config.get(f"show_{key}", default)))
            for key, default in REPORT_VISIBILITY_DEFAULTS.items()
        }

        self._build_ui()
        self._auto_load_paths()
        self._apply_report_visibility(center=False)
        self._center_window()

    # ---- layout -----------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}
        muted_fg = "#888888"
        muted_font = ("Helvetica", 10)
        self.columnconfigure(0, weight=1)
        self._report_widgets = {}

        row = 0

        # --- Header banner ---
        header = tk.Frame(self, bg="#2e7d32", padx=16, pady=12)
        header.grid(row=row, column=0, columnspan=2, sticky="we")
        tk.Label(
            header, text="Optihome Payroll", font=("Helvetica", 18, "bold"),
            fg="white", bg="#2e7d32",
        ).pack(anchor="w")
        steps = tk.Frame(header, bg="#2e7d32")
        steps.pack(anchor="w", pady=(6, 0))
        for i, step in enumerate([
            "Verify the period end date",
            "Select all applicable input reports to process; adjust visible report options in Advanced Settings",
            "Choose the output file",
            "Click Run Export",
        ], 1):
            tk.Label(
                steps, text=f"{i}. {step}", font=("Helvetica", 11),
                fg="#c8e6c9", bg="#2e7d32", anchor="w",
            ).pack(anchor="w")
        row += 1

        # --- Period End Day ---
        day_frame = tk.Frame(self)
        day_frame.grid(row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        label_with_tip(
            day_frame,
            "Period end day:",
            "The app uses the most recent selected weekday to auto-suggest report and output filenames.",
        ).pack(side="left")
        day_combo = ttk.Combobox(
            day_frame, textvariable=self._end_day_var,
            values=DAYS_OF_WEEK, width=12, state="readonly",
        )
        day_combo.pack(side="left", padx=(6, 16))
        self._end_date_label = tk.Label(day_frame, text="", fg="#2e7d32")
        self._end_date_label.pack(side="left")
        day_combo.bind("<<ComboboxSelected>>", lambda e: self._on_day_changed())
        self._refresh_end_date()
        row += 1

        tk.Frame(self, height=4).grid(row=row, column=0, columnspan=2)
        row += 1

        # --- Notion CSV ---
        notion_label = label_with_tip(
            self,
            "Notion Report:",
            "Bi-weekly contractor time report exported from Notion. Use files named like 04-22-2026_notion.csv.",
        )
        notion_label.grid(row=row, column=0, columnspan=2, sticky="w", **pad)
        row += 1
        self._notion_display = tk.StringVar()
        self._notion_entry = tk.Entry(
            self, textvariable=self._notion_display, width=52,
            state="readonly", readonlybackground="white", fg="black"
        )
        self._notion_entry.grid(row=row, column=0, sticky="we", padx=(12, 4), pady=2)
        notion_btn_frame = tk.Frame(self)
        notion_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(notion_btn_frame, text="Browse\u2026", width=10, command=self._browse_notion).pack(
            side="left", padx=(0, 2)
        )
        self._notion_clear_btn = tk.Button(
            notion_btn_frame, text="\u2715", width=2, command=self._clear_notion, state="disabled"
        )
        self._notion_clear_btn.pack(side="left")
        row += 1
        self._notion_full_label = tk.Label(
            self, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=420, justify="left"
        )
        self._notion_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))
        self._report_widgets["notion"] = [
            notion_label, self._notion_entry, notion_btn_frame, self._notion_full_label
        ]
        row += 1

        # --- Turno CSV (primary input) ---
        turno_label = label_with_tip(
            self,
            "Turno Report:",
            "Cleaning job report exported from Turno. This feeds the Mango Villas, Casa Damisela, and Other sections.",
        )
        turno_label.grid(row=row, column=0, columnspan=2, sticky="w", **pad)
        row += 1
        self._turno_display = tk.StringVar()
        self._turno_entry = tk.Entry(
            self, textvariable=self._turno_display, width=52,
            state="readonly", readonlybackground="white", fg="black"
        )
        self._turno_entry.grid(row=row, column=0, sticky="we", padx=(12, 4), pady=2)
        turno_btn_frame = tk.Frame(self)
        turno_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(turno_btn_frame, text="Browse\u2026", width=10, command=self._browse_turno).pack(
            side="left", padx=(0, 2)
        )
        self._turno_clear_btn = tk.Button(turno_btn_frame, text="\u2715", width=2, command=self._clear_turno, state="disabled")
        self._turno_clear_btn.pack(side="left")
        row += 1
        self._turno_full_label = tk.Label(
            self, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=420, justify="left"
        )
        self._turno_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))
        self._report_widgets["turno"] = [
            turno_label, self._turno_entry, turno_btn_frame, self._turno_full_label
        ]
        row += 1

        # --- Expenses CSV ---
        expenses_label = label_with_tip(
            self,
            "Expenses Report:",
            "Expense export from Notion. Rows are listed on each worker sheet; only Reimbursable = Yes adds to payroll.",
        )
        expenses_label.grid(row=row, column=0, columnspan=2, sticky="w", **pad)
        row += 1
        self._expenses_display = tk.StringVar()
        self._expenses_entry = tk.Entry(
            self, textvariable=self._expenses_display, width=52,
            state="readonly", readonlybackground="white", fg="black"
        )
        self._expenses_entry.grid(row=row, column=0, sticky="we", padx=(12, 4), pady=2)
        expenses_btn_frame = tk.Frame(self)
        expenses_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(expenses_btn_frame, text="Browse\u2026", width=10, command=self._browse_expenses).pack(
            side="left", padx=(0, 2)
        )
        self._expenses_clear_btn = tk.Button(
            expenses_btn_frame, text="\u2715", width=2, command=self._clear_expenses, state="disabled"
        )
        self._expenses_clear_btn.pack(side="left")
        row += 1
        self._expenses_full_label = tk.Label(
            self, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=420, justify="left"
        )
        self._expenses_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))
        self._report_widgets["expenses"] = [
            expenses_label, self._expenses_entry, expenses_btn_frame, self._expenses_full_label
        ]
        row += 1

        # --- Output Excel ---
        output_label = label_with_tip(
            self,
            "Output File:",
            "Excel workbook to create. The app auto-suggests a file in Timesheets based on the period end date.",
        )
        output_label.grid(row=row, column=0, columnspan=2, sticky="w", **pad)
        row += 1
        self._output_display = tk.StringVar()
        tk.Entry(self, textvariable=self._output_display, width=52,
                 state="readonly", readonlybackground="white", fg="black").grid(
            row=row, column=0, sticky="we", padx=(12, 4), pady=2
        )
        output_btn_frame = tk.Frame(self)
        output_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(output_btn_frame, text="Save As\u2026", width=10, command=self._browse_output).pack(
            side="left", padx=(0, 2)
        )
        self._output_clear_btn = tk.Button(output_btn_frame, text="\u2715", width=2, command=self._clear_output, state="disabled")
        self._output_clear_btn.pack(side="left")
        row += 1
        self._output_full_label = tk.Label(
            self, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=420, justify="left"
        )
        self._output_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))
        row += 1

        # --- Run button ---
        self._run_btn = tk.Button(
            self,
            text="Run Export",
            command=self._run_export,
            width=18,
            height=2,
            bg=RUN_BUTTON_BG,
            fg=RUN_BUTTON_FG,
            activebackground=RUN_BUTTON_ACTIVE_BG,
            activeforeground=RUN_BUTTON_FG,
            disabledforeground=RUN_BUTTON_DISABLED_FG,
            font=("Helvetica", 15, "bold"),
            relief="raised",
            borderwidth=2,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground="#4db6ac",
        )
        self._run_btn.grid(row=row, column=0, columnspan=2, pady=(18, 10))
        row += 1

        # --- Advanced Settings (collapsed) ---
        self._advanced_visible = False
        advanced_header = tk.Frame(self)
        advanced_header.grid(row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 0))
        self._advanced_btn = tk.Button(
            advanced_header, text="\u25b6  Advanced Settings", command=self._toggle_advanced,
            relief="flat", anchor="w", fg="#555555", font=("Helvetica", 11),
            activeforeground="#333333",
        )
        self._advanced_btn.pack(side="left")
        create_help_icon(
            advanced_header,
            "Choose which report pickers are visible and select optional support files.",
        ).pack(side="left", padx=(7, 0))
        row += 1

        self._advanced_frame = tk.Frame(self, padx=8, pady=8)
        self._advanced_row = row
        self._advanced_frame.grid(row=row, column=0, columnspan=2, sticky="we", padx=12, pady=(0, 4))
        self._build_advanced_section()
        self._advanced_frame.grid_remove()
        row += 1

        # --- Status area ---
        output_log_label = label_with_tip(
            self,
            "Output Log:",
            "Shows export progress, warnings, and errors from the processing script.",
        )
        output_log_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 0))
        row += 1
        frame = tk.Frame(self)
        frame.grid(row=row, column=0, columnspan=2, sticky="nswe", padx=12, pady=(2, 12))
        self._status = tk.Text(frame, height=16, width=68, state="disabled", wrap="word")
        scrollbar = tk.Scrollbar(frame, command=self._status.yview)
        self._status.configure(yscrollcommand=scrollbar.set)
        self._status.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        row += 1

        tk.Label(
            self,
            text="© Optihome Services LLC - 2026",
            anchor="center",
            fg="#888888",
            font=("Helvetica", 10),
        ).grid(row=row, column=0, columnspan=2, sticky="we", padx=12, pady=(0, 12))

        # Tag styles for coloured status messages
        self._status.tag_configure("success", foreground="#2e7d32")
        self._status.tag_configure("warning", foreground="#e65100")
        self._status.tag_configure("error", foreground="#c62828")

    def _build_advanced_section(self):
        """Build report visibility, Timeclock, and Employee Rates fields."""
        f = self._advanced_frame
        muted_fg = "#888888"
        muted_font = ("Helvetica", 10)

        arow = 0

        visibility_label = label_with_tip(
            f,
            "Visible Input Reports:",
            "Checked reports appear in the main app and are included when Run Export is clicked.",
        )
        visibility_label.grid(row=arow, column=0, columnspan=2, sticky="w", pady=(0, 2))
        arow += 1

        visibility_frame = tk.Frame(f)
        visibility_frame.grid(row=arow, column=0, columnspan=2, sticky="w", pady=(0, 8))
        for key, label in [
            ("notion", "Notion Report"),
            ("turno", "Turno Report"),
            ("expenses", "Expenses Report"),
            ("time", "Timeclock File"),
        ]:
            tk.Checkbutton(
                visibility_frame,
                text=label,
                variable=self._report_visible_vars[key],
                command=self._on_report_visibility_changed,
            ).pack(side="left", padx=(0, 12))
        arow += 1

        # --- Timeclock CSV ---
        time_label = label_with_tip(
            f,
            "Timeclock File (optional):",
            "NGTecoTime punch report. The app groups each person's first and last punch per day.",
        )
        time_label.grid(row=arow, column=0, columnspan=2, sticky="w", pady=(0, 2))
        arow += 1
        self._time_display = tk.StringVar()
        self._time_entry = tk.Entry(
            f, textvariable=self._time_display, width=48,
            state="readonly", readonlybackground="white", fg="black"
        )
        self._time_entry.grid(row=arow, column=0, sticky="we", padx=(0, 4), pady=2)
        time_btn_frame = tk.Frame(f)
        time_btn_frame.grid(row=arow, column=1, padx=(0, 0), pady=2)
        tk.Button(time_btn_frame, text="Browse\u2026", width=10, command=self._browse_time).pack(
            side="left", padx=(0, 2)
        )
        self._time_clear_btn = tk.Button(time_btn_frame, text="\u2715", width=2, command=self._clear_time, state="disabled")
        self._time_clear_btn.pack(side="left")
        arow += 1
        self._time_full_label = tk.Label(
            f, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=400, justify="left"
        )
        self._time_full_label.grid(row=arow, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 6))
        self._report_widgets["time"] = [
            time_label, self._time_entry, time_btn_frame, self._time_full_label
        ]
        arow += 1

        # --- Employee Rates CSV ---
        rates_label = label_with_tip(
            f,
            "Employee Rates:",
            "CSV lookup for employee IDs, names, hourly rates, start dates, recurring extras, and notes.",
        )
        rates_label.grid(row=arow, column=0, columnspan=2, sticky="w", pady=(0, 2))
        arow += 1
        self._rates_display = tk.StringVar(value=_shorten_path(self._rates_path))
        tk.Entry(f, textvariable=self._rates_display, width=48,
                 state="readonly", readonlybackground="white", fg="black").grid(
            row=arow, column=0, sticky="we", padx=(0, 4), pady=2
        )
        rates_btn_frame = tk.Frame(f)
        rates_btn_frame.grid(row=arow, column=1, padx=(0, 0), pady=2)
        tk.Button(rates_btn_frame, text="Browse\u2026", width=10, command=self._browse_rates).pack(
            side="top", pady=(0, 2)
        )
        tk.Button(rates_btn_frame, text="Open", width=10, command=self._open_rates_csv).pack(
            side="top"
        )
        arow += 1
        self._rates_full_label = tk.Label(
            f, text=self._rates_path, anchor="w", fg=muted_fg, font=muted_font,
            wraplength=400, justify="left"
        )
        self._rates_full_label.grid(row=arow, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 2))

    def _toggle_advanced(self):
        self._advanced_visible = not self._advanced_visible
        if self._advanced_visible:
            self._advanced_frame.grid()
            self._advanced_btn.config(text="\u25bc  Advanced Settings")
        else:
            self._advanced_frame.grid_remove()
            self._advanced_btn.config(text="\u25b6  Advanced Settings")
        self.geometry("")
        self._center_window()

    def _on_report_visibility_changed(self):
        config = _load_config()
        for key, var in self._report_visible_vars.items():
            config[f"show_{key}"] = bool(var.get())
        _save_config(config)
        self._apply_report_visibility()

    def _is_report_visible(self, key):
        var = self._report_visible_vars.get(key)
        return bool(var.get()) if var is not None else True

    def _apply_report_visibility(self, center=True):
        for key, widgets in getattr(self, "_report_widgets", {}).items():
            visible = self._is_report_visible(key)
            for widget in widgets:
                if visible:
                    widget.grid()
                else:
                    widget.grid_remove()
        self.geometry("")
        if center:
            self._center_window()

    def _refresh_end_date(self):
        """Update the end date label from the current day selection."""
        day = self._end_day_var.get()
        end_date = _get_period_end_date(day)
        self._end_date_label.config(
            text=f"→  {end_date.strftime('%A, %B %d, %Y')}  ({end_date.strftime('%m-%d-%Y')})"
        )

    def _on_day_changed(self):
        self._refresh_end_date()
        _save_config({**_load_config(), "end_day": self._end_day_var.get()})
        self._auto_load_paths()

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{x}+{y}")

    # ---- path helpers -----------------------------------------------------

    def _set_time_path(self, path):
        self._time_path = path
        self._time_display.set(_shorten_path(path) if path else "")
        self._time_full_label.config(text=path)
        self._time_clear_btn.config(state="normal" if path else "disabled")

    def _set_notion_path(self, path):
        self._notion_path = path
        self._notion_display.set(_shorten_path(path) if path else "")
        self._notion_full_label.config(text=path)
        self._notion_clear_btn.config(state="normal" if path else "disabled")

    def _set_turno_path(self, path):
        self._turno_path = path
        self._turno_display.set(_shorten_path(path) if path else "")
        self._turno_full_label.config(text=path)
        self._turno_clear_btn.config(state="normal" if path else "disabled")

    def _set_expenses_path(self, path):
        self._expenses_path = path
        self._expenses_display.set(_shorten_path(path) if path else "")
        self._expenses_full_label.config(text=path)
        self._expenses_clear_btn.config(state="normal" if path else "disabled")

    def _set_output_path(self, path):
        self._output_path = path
        self._output_display.set(_shorten_path(path) if path else "")
        self._output_full_label.config(text=path)
        self._output_clear_btn.config(state="normal" if path else "disabled")

    def _auto_load_paths(self):
        """Auto-select input/output files matching the current period end date."""
        config = _load_config()
        end_date = _get_period_end_date(self._end_day_var.get())
        date_str = end_date.strftime("%m-%d-%Y")
        year_str = end_date.strftime("%Y")

        time_dir = config.get("last_time_dir", self._raw_dir)
        turno_dir = config.get("last_turno_dir", self._raw_dir)
        notion_dir = config.get("last_notion_dir", self._raw_dir)
        expenses_dir = config.get("last_expenses_dir", self._raw_dir)
        output_dir = config.get("last_output_dir", self._timesheets_dir)

        time_candidate = self._find_report_candidate(time_dir, year_str, f"{date_str}_time.csv")
        self._set_time_path(time_candidate if os.path.isfile(time_candidate) else "")

        turno_candidate = self._find_report_candidate(turno_dir, year_str, f"{date_str}_turno.csv")
        self._set_turno_path(turno_candidate if os.path.isfile(turno_candidate) else "")

        notion_candidate = self._find_report_candidate(notion_dir, year_str, f"{date_str}_notion.csv")
        self._set_notion_path(notion_candidate if os.path.isfile(notion_candidate) else "")

        expenses_candidate = self._find_report_candidate(expenses_dir, year_str, f"{date_str}_expenses.csv")
        self._set_expenses_path(expenses_candidate if os.path.isfile(expenses_candidate) else "")

        output_candidate = os.path.join(output_dir, f"{date_str}.xlsx")
        self._set_output_path(output_candidate)

    def _find_report_candidate(self, preferred_dir, year_str, filename):
        search_dirs = [
            preferred_dir,
            os.path.join(preferred_dir, year_str),
            self._raw_dir,
            os.path.join(self._raw_dir, year_str),
        ]
        seen = set()
        for folder in search_dirs:
            if not folder or folder in seen:
                continue
            seen.add(folder)
            candidate = os.path.join(folder, filename)
            if os.path.isfile(candidate):
                return candidate
        return os.path.join(preferred_dir, filename)

    # ---- file dialogs -----------------------------------------------------

    def _browse_time(self):
        config = _load_config()
        initial_dir = config.get("last_time_dir", self._raw_dir)
        path = filedialog.askopenfilename(
            title="Select Timeclock CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            _save_config({**config, "last_time_dir": os.path.dirname(path)})
            self._set_time_path(path)
            self._suggest_output(path)

    def _clear_time(self):
        self._set_time_path("")

    def _browse_notion(self):
        config = _load_config()
        initial_dir = config.get("last_notion_dir", self._raw_dir)
        path = filedialog.askopenfilename(
            title="Select Notion CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            _save_config({**config, "last_notion_dir": os.path.dirname(path)})
            self._set_notion_path(path)
            if not self._output_path:
                self._suggest_output(path)

    def _clear_notion(self):
        self._set_notion_path("")

    def _browse_turno(self):
        config = _load_config()
        initial_dir = config.get("last_turno_dir", self._raw_dir)
        path = filedialog.askopenfilename(
            title="Select Turno CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            _save_config({**config, "last_turno_dir": os.path.dirname(path)})
            self._set_turno_path(path)
            if not self._output_path:
                self._suggest_output(path)

    def _clear_turno(self):
        self._set_turno_path("")

    def _browse_expenses(self):
        config = _load_config()
        initial_dir = config.get("last_expenses_dir", self._raw_dir)
        path = filedialog.askopenfilename(
            title="Select Expenses CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            _save_config({**config, "last_expenses_dir": os.path.dirname(path)})
            self._set_expenses_path(path)
            if not self._output_path:
                self._suggest_output(path)

    def _clear_expenses(self):
        self._set_expenses_path("")

    def _browse_output(self):
        config = _load_config()
        initial_dir = config.get("last_output_dir", self._timesheets_dir)
        initial_name = self._suggested_output_name() or "output.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Excel Output As",
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            _save_config({**config, "last_output_dir": os.path.dirname(path)})
            self._set_output_path(path)

    def _clear_output(self):
        self._set_output_path("")

    def _suggested_output_name(self):
        """Derive an output filename from whichever input CSV is set."""
        for source in (self._notion_path, self._turno_path, self._expenses_path, self._time_path):
            if source:
                basename = os.path.basename(source)
                match = re.search(r"(\d{2}-\d{2}-\d{4})", basename)
                if match:
                    return f"{match.group(1)}.xlsx"
        return None

    def _suggest_output(self, source_path):
        """Auto-fill the output path when an input file is selected."""
        if self._output_path:
            return  # don't overwrite an already-chosen path
        basename = os.path.basename(source_path)
        match = re.search(r"(\d{2}-\d{2}-\d{4})", basename)
        if match:
            output_dir = _load_config().get("last_output_dir", self._timesheets_dir)
            path = os.path.join(output_dir, f"{match.group(1)}.xlsx")
            self._set_output_path(path)

    # ---- rates CSV --------------------------------------------------------

    def _browse_rates(self):
        path = filedialog.askopenfilename(
            title="Select Employee Rates CSV",
            initialdir=os.path.dirname(self._rates_path) or _get_project_dir(),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self._rates_path = path
            self._rates_display.set(_shorten_path(path))
            self._rates_full_label.config(text=path)

    def _open_rates_csv(self):
        """Open the selected rates CSV in the default editor."""
        rates_path = self._rates_path.strip()
        if not rates_path or not os.path.exists(rates_path):
            messagebox.showerror(
                "File not found",
                f"Cannot find rates file at:\n{rates_path}\n\n"
                "Use Browse to locate timesheet-rates.csv.",
            )
            return
        subprocess.Popen(["open", rates_path])

    # ---- run export -------------------------------------------------------

    def _append_status(self, text, tag=None):
        self._status.configure(state="normal")
        if tag:
            self._status.insert("end", text + "\n", tag)
        else:
            self._status.insert("end", text + "\n")
        self._status.see("end")
        self._status.configure(state="disabled")

    def _clear_status(self):
        self._status.configure(state="normal")
        self._status.delete("1.0", "end")
        self._status.configure(state="disabled")

    def _run_export(self):
        notion_csv = self._notion_path.strip() if self._is_report_visible("notion") else ""
        turno_csv = self._turno_path.strip() if self._is_report_visible("turno") else ""
        expenses_csv = self._expenses_path.strip() if self._is_report_visible("expenses") else ""
        time_csv = self._time_path.strip() if self._is_report_visible("time") else ""
        output_xlsx = self._output_path.strip()

        if not notion_csv and not turno_csv and not expenses_csv and not time_csv:
            messagebox.showwarning(
                "Missing file",
                "Please select at least one visible input CSV file (Notion, Turno, Expenses, or Timeclock).",
            )
            return
        if not output_xlsx:
            messagebox.showwarning("Missing file", "Please choose an output Excel file location.")
            return

        rates_csv = self._rates_path.strip()

        self._clear_status()
        self._append_status("Running export...")
        self._run_btn.configure(
            text="Running...",
            state="disabled",
            bg=RUN_BUTTON_DISABLED_BG,
            fg=RUN_BUTTON_DISABLED_FG,
        )

        # Run in a background thread so the UI stays responsive
        thread = threading.Thread(
            target=self._export_thread,
            args=(
                time_csv or None,
                turno_csv or None,
                notion_csv or None,
                expenses_csv or None,
                output_xlsx,
                rates_csv,
            ),
            daemon=True,
        )
        thread.start()

    def _export_thread(self, time_csv, turno_csv, notion_csv, expenses_csv, output_xlsx, rates_csv):
        try:
            message, warnings = process_timesheet(
                time_csv,
                output_xlsx,
                turno_csv,
                rates_csv=rates_csv,
                notion_csv=notion_csv,
                expenses_csv=expenses_csv,
            )
            self.after(0, self._on_export_done, message, warnings, None, output_xlsx)
        except Exception as exc:
            self.after(0, self._on_export_done, None, [], exc, None)

    def _on_export_done(self, message, warnings, error, output_path):
        self._run_btn.configure(text="Run Export", state="normal", bg=RUN_BUTTON_BG, fg=RUN_BUTTON_FG)
        if error:
            self._append_status(f"ERROR: {error}", "error")
            messagebox.showerror("Export failed", str(error))
        else:
            for w in warnings:
                self._append_status(f"Warning: {w}", "warning")
            self._append_status(message, "success")
            messagebox.showinfo("Done", message)
            if output_path and os.path.exists(output_path):
                subprocess.Popen(["open", output_path])


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = PayrollApp()
    app.mainloop()

"""
Optihome Payroll Processing — Tkinter GUI

Provides file pickers for the timeclock CSV, turno CSV, and output Excel file,
then runs the export-timesheet process and displays results.
"""

import importlib
import os
import re
import sys
import threading
import subprocess
import tkinter as tk
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

    Non-frozen: directory of this .py file.
    Frozen (.app): walk upward from the .app bundle looking for timesheet-rates.csv.
    Falls back to the directory that contains the .app bundle.
    """
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(__file__))

    # sys.executable is …/Foo.app/Contents/MacOS/Foo
    # Go up to the directory containing the .app bundle
    app_dir = os.path.dirname(sys.executable)          # Contents/MacOS
    app_dir = os.path.dirname(app_dir)                 # Contents
    app_dir = os.path.dirname(app_dir)                 # Foo.app
    bundle_parent = os.path.dirname(app_dir)           # dir containing .app

    # Search upward (up to 3 levels) for timesheet-rates.csv
    candidate = bundle_parent
    for _ in range(4):
        if os.path.isfile(os.path.join(candidate, "timesheet-rates.csv")):
            return candidate
        parent = os.path.dirname(candidate)
        if parent == candidate:
            break
        candidate = parent

    return bundle_parent

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

def _shorten_path(path, segments=3):
    """Return the last *segments* components of a path, prefixed with /."""
    if not path:
        return ""
    parts = path.replace("\\", "/").rstrip("/").split("/")
    parts = [p for p in parts if p]
    if len(parts) <= segments:
        return path
    return "/" + "/".join(parts[-segments:])

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
        self._output_path = ""
        self._rates_path = os.path.join(self._project_dir, "timesheet-rates.csv")

        self._build_ui()
        self._center_window()

    # ---- layout -----------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}
        muted_fg = "#888888"
        muted_font = ("Helvetica", 10)

        row = 0

        # --- Timeclock CSV ---
        tk.Label(self, text="Timeclock CSV (_time.csv)  \u2014  optional:", anchor="w").grid(
            row=row, column=0, columnspan=2, sticky="w", **pad
        )
        row += 1
        self._time_display = tk.StringVar()
        tk.Entry(self, textvariable=self._time_display, width=52,
                 state="readonly", readonlybackground="white", fg="black").grid(
            row=row, column=0, sticky="we", padx=(12, 4), pady=2
        )
        time_btn_frame = tk.Frame(self)
        time_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(time_btn_frame, text="Browse\u2026", width=10, command=self._browse_time).pack(
            side="left", padx=(0, 2)
        )
        self._time_clear_btn = tk.Button(time_btn_frame, text="\u2715", width=2, command=self._clear_time, state="disabled")
        self._time_clear_btn.pack(side="left")
        row += 1
        self._time_full_label = tk.Label(
            self, text="", anchor="w", fg=muted_fg, font=muted_font, wraplength=420, justify="left"
        )
        self._time_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))

        row += 1

        # --- Turno CSV ---
        tk.Label(self, text="Turno CSV (_turno.csv)  \u2014  optional:", anchor="w").grid(
            row=row, column=0, columnspan=2, sticky="w", **pad
        )
        row += 1
        self._turno_display = tk.StringVar()
        tk.Entry(self, textvariable=self._turno_display, width=52,
                 state="readonly", readonlybackground="white", fg="black").grid(
            row=row, column=0, sticky="we", padx=(12, 4), pady=2
        )
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

        row += 1

        # --- Output Excel ---
        tk.Label(self, text="Output Excel file (.xlsx):", anchor="w").grid(
            row=row, column=0, columnspan=2, sticky="w", **pad
        )
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

        # --- Employee Rates CSV ---
        tk.Label(self, text="Employee Rates CSV (timesheet-rates.csv):", anchor="w").grid(
            row=row, column=0, columnspan=2, sticky="w", **pad
        )
        row += 1
        self._rates_display = tk.StringVar(value=_shorten_path(self._rates_path))
        tk.Entry(self, textvariable=self._rates_display, width=52,
                 state="readonly", readonlybackground="white", fg="black").grid(
            row=row, column=0, sticky="we", padx=(12, 4), pady=2
        )
        rates_btn_frame = tk.Frame(self)
        rates_btn_frame.grid(row=row, column=1, padx=(0, 12), pady=2)
        tk.Button(rates_btn_frame, text="Browse\u2026", width=10, command=self._browse_rates).pack(
            side="top", pady=(0, 2)
        )
        tk.Button(rates_btn_frame, text="Open", width=10, command=self._open_rates_csv).pack(
            side="top"
        )
        row += 1
        self._rates_full_label = tk.Label(
            self, text=self._rates_path, anchor="w", fg=muted_fg, font=muted_font,
            wraplength=420, justify="left"
        )
        self._rates_full_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))

        row += 1

        # --- Run button ---
        self._run_btn = ttk.Button(
            self, text="Run Export", command=self._run_export,
        )
        self._run_btn.grid(row=row, column=0, columnspan=2, pady=(12, 4), ipady=4)

        row += 1

        # --- Status area ---
        tk.Label(self, text="Status:", anchor="w").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 0)
        )
        row += 1
        frame = tk.Frame(self)
        frame.grid(row=row, column=0, columnspan=2, sticky="nswe", padx=12, pady=(2, 12))
        self._status = tk.Text(frame, height=10, width=62, state="disabled", wrap="word")
        scrollbar = tk.Scrollbar(frame, command=self._status.yview)
        self._status.configure(yscrollcommand=scrollbar.set)
        self._status.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Tag styles for coloured status messages
        self._status.tag_configure("success", foreground="#2e7d32")
        self._status.tag_configure("warning", foreground="#e65100")
        self._status.tag_configure("error", foreground="#c62828")

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{x}+{y}")

    # ---- file dialogs -----------------------------------------------------

    def _browse_time(self):
        path = filedialog.askopenfilename(
            title="Select Timeclock CSV",
            initialdir=self._raw_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self._time_path = path
            self._time_display.set(_shorten_path(path))
            self._time_full_label.config(text=path)
            self._time_clear_btn.config(state="normal")
            self._suggest_output(path)

    def _clear_time(self):
        self._time_path = ""
        self._time_display.set("")
        self._time_full_label.config(text="")
        self._time_clear_btn.config(state="disabled")

    def _browse_turno(self):
        path = filedialog.askopenfilename(
            title="Select Turno CSV",
            initialdir=self._raw_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self._turno_path = path
            self._turno_display.set(_shorten_path(path))
            self._turno_full_label.config(text=path)
            self._turno_clear_btn.config(state="normal")
            # Suggest output if not already set
            if not self._output_path:
                self._suggest_output(path)

    def _clear_turno(self):
        self._turno_path = ""
        self._turno_display.set("")
        self._turno_full_label.config(text="")
        self._turno_clear_btn.config(state="disabled")

    def _browse_output(self):
        initial_name = self._suggested_output_name() or "output.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Excel Output As",
            initialdir=self._timesheets_dir,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._output_path = path
            self._output_display.set(_shorten_path(path))
            self._output_full_label.config(text=path)
            self._output_clear_btn.config(state="normal")

    def _clear_output(self):
        self._output_path = ""
        self._output_display.set("")
        self._output_full_label.config(text="")
        self._output_clear_btn.config(state="disabled")

    def _suggested_output_name(self):
        """Derive an output filename from whichever input CSV is set."""
        for source in (self._time_path, self._turno_path):
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
            path = os.path.join(self._timesheets_dir, f"{match.group(1)}.xlsx")
            self._output_path = path
            self._output_display.set(_shorten_path(path))
            self._output_full_label.config(text=path)

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
        time_csv = self._time_path.strip()
        turno_csv = self._turno_path.strip()
        output_xlsx = self._output_path.strip()

        if not time_csv and not turno_csv:
            messagebox.showwarning("Missing file", "Please select at least one input CSV file (Timeclock or Turno).")
            return
        if not output_xlsx:
            messagebox.showwarning("Missing file", "Please choose an output Excel file location.")
            return

        rates_csv = self._rates_path.strip()

        self._clear_status()
        self._append_status("Running export...")
        self._run_btn.configure(text="Running...")
        self._run_btn.state(["disabled"])

        # Run in a background thread so the UI stays responsive
        thread = threading.Thread(
            target=self._export_thread,
            args=(time_csv or None, turno_csv or None, output_xlsx, rates_csv),
            daemon=True,
        )
        thread.start()

    def _export_thread(self, time_csv, turno_csv, output_xlsx, rates_csv):
        try:
            message, warnings = process_timesheet(
                time_csv, output_xlsx, turno_csv, rates_csv=rates_csv
            )
            self.after(0, self._on_export_done, message, warnings, None, output_xlsx)
        except Exception as exc:
            self.after(0, self._on_export_done, None, [], exc, None)

    def _on_export_done(self, message, warnings, error, output_path):
        self._run_btn.state(["!disabled"])
        self._run_btn.configure(text="Run Export")
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

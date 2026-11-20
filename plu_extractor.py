"""PLU Extractor

Reads an .xlsx or .csv file, keeps only the values of the 'PLU' column,
adds a column filled with 1 to its right, and writes the result as a CSV
with no header row.

Usage:
  - Double-click the exe and pick a file, or
  - Drag a file onto the exe (output is written next to the input as <name>_PLU.csv).
"""

import csv
import datetime
import os
import sys

TARGET_HEADER = "PLU"
APP_NAME = "PLU Extractor"


class ExtractError(Exception):
    pass


def _cell_to_str(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, float) and value.is_integer():
        # Excel stores numeric PLUs as floats; avoid "12345.0"
        return str(int(value))
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()
    return str(value)


def _read_csv(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines(keepends=True)
    sample = "".join(lines[:50])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [list(row) for row in csv.reader(lines, dialect)]


def _read_xlsx(path):
    import openpyxl

    # Not read_only: read-only mode doesn't expose which rows are hidden.
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.active
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True), start=ws.min_row):
            # Skip rows hidden by an active filter (or manually hidden/collapsed)
            dim = ws.row_dimensions.get(i)
            if dim is not None and dim.hidden:
                continue
            rows.append([_cell_to_str(v) for v in row])
        return rows
    finally:
        wb.close()


def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return _read_csv(path)
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(path)
    raise ExtractError(f"Unsupported file type '{ext}'. Please use .xlsx or .csv.")


def extract(path, out_path):
    rows = read_rows(path)

    # Drop rows that are entirely empty (e.g. formatted-but-blank Excel rows)
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise ExtractError("The file is empty.")

    # The header may not be in the first row (e.g. title/notes rows above it):
    # use the first row that contains a cell matching TARGET_HEADER.
    header_row = idx = None
    for ri, row in enumerate(rows):
        for ci, h in enumerate(row):
            if str(h).strip().lower() == TARGET_HEADER.lower():
                header_row, idx = ri, ci
                break
        if idx is not None:
            break
    if idx is None:
        raise ExtractError(f"No '{TARGET_HEADER}' column header found anywhere in the file.")

    count = 0
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in rows[header_row + 1:]:
            plu = r[idx].strip() if idx < len(r) else ""
            w.writerow([plu, 1])
            count += 1
    return count


def default_output_path(path):
    base, _ = os.path.splitext(path)
    return base + "_PLU.csv"


# ---------------------------------------------------------------- GUI
def enable_high_dpi():
    """Tell Windows we render at native DPI so it doesn't bitmap-stretch (blur) the UI."""
    if sys.platform != "win32":
        return
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor aware (Win 8.1+)
    except Exception:  # noqa: BLE001
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # system aware (Vista+)
        except Exception:  # noqa: BLE001
            pass


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.title(APP_NAME)
    root.resizable(False, False)

    frm = tk.Frame(root, padx=16, pady=16)
    frm.pack()

    tk.Label(
        frm,
        text=f"Keeps the '{TARGET_HEADER}' column values, adds a column of 1s,\nand saves the result as a CSV (no header row).",
        justify="left",
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

    status = tk.StringVar(value="")

    def choose():
        path = filedialog.askopenfilename(
            title="Select an Excel or CSV file",
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        suggested = default_output_path(path)
        out = filedialog.asksaveasfilename(
            title="Save output CSV",
            initialdir=os.path.dirname(suggested),
            initialfile=os.path.basename(suggested),
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not out:
            return
        try:
            n = extract(path, out)
        except ExtractError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except PermissionError:
            messagebox.showerror(APP_NAME, "Could not write the output file.\nIs it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Unexpected error:\n{e}")
            return
        status.set(f"Done: {n} rows written.")
        messagebox.showinfo(APP_NAME, f"{n} rows written to:\n{out}")

    tk.Button(frm, text="Select file...", width=24, command=choose).grid(
        row=2, column=0, columnspan=2, pady=(14, 4)
    )
    tk.Label(frm, textvariable=status, fg="green").grid(row=3, column=0, columnspan=2)

    root.mainloop()


def run_dropped(paths):
    """Files dragged onto the exe: process each without dialogs, then report."""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    msgs, ok = [], True
    for p in paths:
        out = default_output_path(p)
        try:
            n = extract(p, out)
            msgs.append(f"{os.path.basename(p)} -> {os.path.basename(out)} ({n} rows)")
        except Exception as e:  # noqa: BLE001
            ok = False
            msgs.append(f"{os.path.basename(p)}: ERROR - {e}")
    (messagebox.showinfo if ok else messagebox.showerror)(APP_NAME, "\n\n".join(msgs))
    root.destroy()


if __name__ == "__main__":
    enable_high_dpi()
    if len(sys.argv) > 1:
        run_dropped(sys.argv[1:])
    else:
        run_gui()

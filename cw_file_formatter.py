"""CW File Formatter

Reads an .xlsx or .csv file, or cells pasted from Excel, and writes a
two-column CSV (PLU, QTY) with no header row.

  - No header row (default): column 1 is the PLU and column 2 is the QTY.
    A missing or blank QTY becomes the QTY entered in the window (default 1).
  - Header row: keeps only the values of the 'PLU' column and sets every QTY
    to the QTY entered in the window (default 1).

Usage:
  - Double-click the exe, pick a file (or paste cells copied from Excel), set
    the header option and QTY, and click Convert (for a file, output defaults
    to <name>_PLU.csv next to the input), or
  - Drag a file onto the exe (it is read as having no header row; output is
    written next to the input as <name>_PLU.csv).
"""

import csv
import datetime
import os
import sys

TARGET_HEADER = "PLU"
APP_NAME = "CW File Formatter"
ICON_FILE = "icon.ico"


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


def _cell(row, i):
    return str(row[i]).strip() if i < len(row) else ""


def parse_pasted(text):
    """Cells copied from Excel arrive as tab-separated lines (filtered-out rows excluded)."""
    import io

    return [list(row) for row in csv.reader(io.StringIO(text), delimiter="\t")]


def extract(path, out_path, has_header=False, qty="1"):
    """Convert an .xlsx/.csv file. See extract_rows."""
    return extract_rows(read_rows(path), out_path, has_header, qty, source="file")


def extract_rows(rows, out_path, has_header=False, qty="1", source="input"):
    """Write (PLU, QTY) rows to out_path and return the row count.

    qty fills every row in header mode, and blank QTY cells in no-header mode.
    """
    # Drop rows that are entirely empty (e.g. formatted-but-blank Excel rows)
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise ExtractError(f"The {source} is empty.")

    if not has_header:
        # No header row: first column is PLU, second is QTY.
        count = 0
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            for r in rows:
                w.writerow([_cell(r, 0), _cell(r, 1) or qty])
                count += 1
        return count

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
        raise ExtractError(f"No '{TARGET_HEADER}' column header found anywhere in the {source}.")

    count = 0
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in rows[header_row + 1:]:
            w.writerow([_cell(r, idx), qty])
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


def resource_path(name):
    """Path to a bundled file, both when run from source and from the PyInstaller exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def set_window_icon(root):
    try:
        root.iconbitmap(default=resource_path(ICON_FILE))
    except Exception:  # noqa: BLE001 - a missing icon must never stop the app
        pass


def run_gui():
    import subprocess
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(APP_NAME)
    root.resizable(False, False)
    set_window_icon(root)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Ok.TLabel", foreground="#1a7f37")

    frm = ttk.Frame(root, padding=16)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    mode = tk.StringVar(value="file")  # "file" or "paste"
    in_var = tk.StringVar()
    out_var = tk.StringVar()
    has_header = tk.BooleanVar(value=False)
    qty_var = tk.StringVar(value="1")
    status = tk.StringVar()
    paste_info = tk.StringVar()
    last_out = {"path": None}

    def auto_output():
        """The output path suggested for the current input ('' if there isn't one)."""
        if mode.get() == "file" and in_var.get().strip():
            return os.path.normpath(default_output_path(in_var.get().strip()))
        return ""

    # --- input source switch
    ttk.Label(frm, text="Input:").grid(row=0, column=0, sticky="w")
    src_row = ttk.Frame(frm)
    src_row.grid(row=0, column=1, columnspan=2, sticky="w", padx=6)

    # --- file input
    file_box = ttk.Frame(frm)
    file_box.columnconfigure(1, weight=1)

    def browse_input():
        path = filedialog.askopenfilename(
            title="Select an Excel or CSV file",
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
        )
        if path:
            in_var.set(os.path.normpath(path))
            out_var.set(auto_output())
            clear_status()

    ttk.Label(file_box, text="Input file:").grid(row=0, column=0, sticky="w")
    in_entry = ttk.Entry(file_box, textvariable=in_var, width=48)
    in_entry.grid(row=0, column=1, sticky="ew", padx=6)
    ttk.Button(file_box, text="Browse…", command=browse_input).grid(row=0, column=2)

    # --- paste input
    paste_box = ttk.Frame(frm)
    paste_box.columnconfigure(0, weight=1)
    ttk.Label(
        paste_box, text="Copy the cells in Excel, then click in the box below and press Ctrl+V."
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
    paste_txt = tk.Text(
        paste_box, height=10, width=60, wrap="none", undo=True, relief="solid", borderwidth=1
    )
    paste_txt.grid(row=1, column=0, sticky="nsew")
    paste_scroll = ttk.Scrollbar(paste_box, orient="vertical", command=paste_txt.yview)
    paste_scroll.grid(row=1, column=1, sticky="ns")
    paste_txt.configure(yscrollcommand=paste_scroll.set)

    paste_foot = ttk.Frame(paste_box)
    paste_foot.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
    paste_foot.columnconfigure(0, weight=1)
    ttk.Label(paste_foot, textvariable=paste_info).grid(row=0, column=0, sticky="w")

    def paste_clipboard():
        try:
            data = root.clipboard_get()
        except tk.TclError:
            messagebox.showerror(APP_NAME, "The clipboard is empty.\nCopy the cells in Excel first.")
            return
        paste_txt.delete("1.0", "end")
        paste_txt.insert("1.0", data)

    ttk.Button(paste_foot, text="Paste", command=paste_clipboard).grid(row=0, column=1)
    ttk.Button(paste_foot, text="Clear", command=lambda: paste_txt.delete("1.0", "end")).grid(
        row=0, column=2, padx=(6, 0)
    )

    def pasted_text():
        return paste_txt.get("1.0", "end-1c")

    def on_paste_change(_=None):
        n = sum(1 for r in parse_pasted(pasted_text()) if any(c.strip() for c in r))
        paste_info.set(f"{n} row{'s' if n != 1 else ''} pasted" if n else "")
        paste_txt.edit_modified(False)
        clear_status()

    paste_txt.bind("<<Modified>>", on_paste_change)

    def switch_mode(*_):
        if mode.get() == "file":
            paste_box.grid_remove()
            file_box.grid()
            if not out_var.get().strip():
                out_var.set(auto_output())
        else:
            file_box.grid_remove()
            paste_box.grid()
            # A path suggested from the input file doesn't apply to pasted data
            if in_var.get().strip() and out_var.get().strip() == os.path.normpath(
                default_output_path(in_var.get().strip())
            ):
                out_var.set("")
            paste_txt.focus_set()
        clear_status()

    for value, label in (("file", "Select file"), ("paste", "Paste from Excel")):
        ttk.Radiobutton(
            src_row, text=label, value=value, variable=mode, command=switch_mode
        ).pack(side="left", padx=(0, 12))

    file_box.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))
    paste_box.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))
    paste_box.grid_remove()

    # --- output file
    def browse_output():
        current = out_var.get().strip() or auto_output() or "pasted_PLU.csv"
        path = filedialog.asksaveasfilename(
            title="Save output CSV as",
            initialdir=os.path.dirname(current) or None,
            initialfile=os.path.basename(current),
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if path:
            out_var.set(os.path.normpath(path))
        return bool(path)

    out_row = ttk.Frame(frm)
    out_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
    out_row.columnconfigure(1, weight=1)
    ttk.Label(out_row, text="Output file:").grid(row=0, column=0, sticky="w")
    out_entry = ttk.Entry(out_row, textvariable=out_var, width=48)
    out_entry.grid(row=0, column=1, sticky="ew", padx=6)
    ttk.Button(out_row, text="Save as…", command=browse_output).grid(row=0, column=2)

    # When a path is filled in for the user, keep the file name (end of a long path) visible
    for var, entry in ((in_var, in_entry), (out_var, out_entry)):
        var.trace_add(
            "write",
            lambda *_, e=entry: None if e.focus_get() is e else e.after_idle(e.xview_moveto, 1.0),
        )

    ttk.Separator(frm).grid(row=3, column=0, columnspan=3, sticky="ew", pady=12)

    # --- options
    ttk.Checkbutton(frm, text="Data has a header row", variable=has_header).grid(
        row=4, column=0, columnspan=3, sticky="w"
    )
    ttk.Label(frm, text="QTY:").grid(row=5, column=0, sticky="w", pady=(8, 0))
    ttk.Spinbox(frm, textvariable=qty_var, from_=1, to=99999, increment=1, width=8).grid(
        row=5, column=1, sticky="w", padx=6, pady=(8, 0)
    )

    # --- action + result
    bottom = ttk.Frame(frm)
    bottom.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(16, 0))
    bottom.columnconfigure(0, weight=1)

    status_lbl = ttk.Label(bottom, textvariable=status, style="Ok.TLabel")
    status_lbl.grid(row=0, column=0, sticky="w")

    def open_folder():
        if last_out["path"] and os.path.exists(last_out["path"]):
            subprocess.Popen(["explorer", "/select,", last_out["path"]])

    open_btn = ttk.Button(bottom, text="Open folder", command=open_folder)

    def clear_status():
        status.set("")
        open_btn.grid_remove()

    def convert(event=None):
        if event is not None and event.widget is paste_txt:
            return None  # Enter in the paste box is a new line, not Convert
        qty = qty_var.get().strip()
        try:
            float(qty)
        except ValueError:
            messagebox.showerror(APP_NAME, "QTY must be a number.")
            return

        pasting = mode.get() == "paste"
        if pasting:
            text = pasted_text()
            if not text.strip():
                messagebox.showerror(APP_NAME, "Paste the cells from Excel first.")
                return
        else:
            src = in_var.get().strip()
            if not src:
                messagebox.showerror(APP_NAME, "Choose an input file first.")
                return
            if not os.path.isfile(src):
                messagebox.showerror(APP_NAME, f"Input file not found:\n{src}")
                return

        out = out_var.get().strip()
        if not out:
            if pasting:
                # No input file to put the output next to: ask where to save.
                # (The Save dialog asks about replacing an existing file itself.)
                if not browse_output():
                    return
                out = out_var.get().strip()
            else:
                out = auto_output()
                out_var.set(out)
            confirmed = pasting
        else:
            confirmed = False

        if not pasting and os.path.abspath(out) == os.path.abspath(src):
            messagebox.showerror(APP_NAME, "The output file can't be the same as the input file.")
            return
        if not confirmed and os.path.exists(out) and not messagebox.askyesno(
            APP_NAME, f"{os.path.basename(out)} already exists.\nReplace it?"
        ):
            return

        clear_status()
        try:
            if pasting:
                n = extract_rows(
                    parse_pasted(text), out, has_header=has_header.get(), qty=qty,
                    source="pasted data",
                )
            else:
                n = extract(src, out, has_header=has_header.get(), qty=qty)
        except ExtractError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except PermissionError:
            messagebox.showerror(APP_NAME, "Could not write the output file.\nIs it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Unexpected error:\n{e}")
            return
        last_out["path"] = out
        status.set(f"Done: {n} rows written to {os.path.basename(out)}")
        open_btn.grid(row=0, column=1, padx=(8, 0))

    ttk.Button(bottom, text="Convert", command=convert, default="active").grid(
        row=0, column=2, padx=(8, 0)
    )
    root.bind("<Return>", convert)

    # Centre on screen
    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 3
    root.geometry(f"+{x}+{y}")

    root.mainloop()


def run_dropped(paths):
    """Files dragged onto the exe: process each without dialogs, then report."""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    set_window_icon(root)
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

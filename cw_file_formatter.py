"""CW File Formatter

The home page offers two tasks:

Generate PLU file
  Reads an .xlsx or .csv file, or cells pasted from Excel, and writes a
  two-column CSV (PLU, QTY) with no header row. The PLU column can be chosen in
  the window; every other column is ignored and every QTY is the QTY entered in
  the window (default 1). Rows without a PLU are left out.

  - No header row (default): the first column is assumed to be the PLU.
  - Header row: the column headed 'PLU' is assumed, and rows up to the header
    row are left out.

Process B&W Flight
  Reads a B&W flight workbook (.xlsx or .csv) and lists its sections, e.g.
  "HEALTHY BREAK 24A BANANA BOAT", so the user can pick which to process (all
  are selected to begin with).
  A section starts at a row whose first filled cell is text (the section
  title); a 'MYCHEM ID' row gives the column headers (a section without its
  own header row reuses the previous one); rows starting with a number are
  the section's items.

  The next page takes a stock report (.xlsx/.csv file or pasted cells) with PLU
  and SOH columns plus a planogram column the user picks. Items are matched on
  PLU and written to a two-column CSV (PLU, QTY), no header row, when
      QTY = MIN QTY 770MM + Planogram - SOH  is at least the threshold.
  A review page lists the items, where they can be left out and their quantities
  changed. Each selected section gets its own CSV, named after the section.

Usage:
  - Double-click the exe and pick a task on the home page (files can be dropped
    onto the window if tkinterdnd2 is installed), or
  - Drag a file onto the exe (its first column becomes a PLU file with a QTY of 1,
    read as having no header row; output is written next to the input as
    <name>_PLU.csv).
"""

import csv
import datetime
import math
import os
import re
import sys
from dataclasses import dataclass, field

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
    if ext == ".xls":
        raise ExtractError("Old .xls files aren't supported.\nOpen it in Excel and save it as .xlsx.")
    raise ExtractError(f"Unsupported file type '{ext}'. Please use .xlsx or .csv.")


def _cell(row, i):
    return str(row[i]).strip() if i < len(row) else ""


def parse_pasted(text):
    """Cells copied from Excel arrive as tab-separated lines (filtered-out rows excluded)."""
    import io

    return [list(row) for row in csv.reader(io.StringIO(text), delimiter="\t")]


def plu_table(rows, has_header=False):
    """Split rows into ([(column index, label)], body rows, index of the assumed PLU column).

    Columns with nothing in them are left out. With a header row (the first row
    holding a 'PLU' cell, else the first row) the assumed PLU column is the one headed
    'PLU', or None if there isn't one. Without, it is the first column.
    """
    # Drop rows that are entirely empty (e.g. formatted-but-blank Excel rows)
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    header, default = [], None
    if has_header and rows:
        # The header may not be in the first row (e.g. title/notes rows above it)
        start = 0
        for ri, row in enumerate(rows):
            names = [str(c).strip().lower() for c in row]
            if TARGET_HEADER.lower() in names:
                start, default = ri, names.index(TARGET_HEADER.lower())
                break
        header, rows = rows[start], rows[start + 1:]

    columns, used = [], set()
    for i in range(max((len(r) for r in [header, *rows]), default=0)):
        label = re.sub(r"\s+", " ", _cell(header, i))
        if not label and not any(_cell(r, i) for r in rows):
            continue
        if not label:
            label = f"Column {i + 1}"
        elif label in used:
            label = f"{label} (column {i + 1})"
        used.add(label)
        columns.append((i, label))
    if not has_header and columns:
        default = columns[0][0]
    return columns, rows, default


def extract(path, out_path, has_header=False, qty="1", plu_col=None):
    """Convert an .xlsx/.csv file. See extract_rows."""
    return extract_rows(read_rows(path), out_path, has_header, qty, plu_col, source="file")


def extract_rows(rows, out_path, has_header=False, qty="1", plu_col=None, source="input"):
    """Write (PLU, QTY) rows to out_path, with the same QTY on every row.

    plu_col is a column index; None uses the assumed PLU column (see plu_table).
    Returns (rows written, rows left out because they have no PLU).
    """
    columns, body, default = plu_table(rows, has_header)
    if not columns:
        raise ExtractError(f"The {source} is empty.")
    if plu_col is None:
        plu_col = default
    if plu_col is None:
        raise ExtractError(f"No '{TARGET_HEADER}' column header found anywhere in the {source}.")
    plus = [plu for plu in (_cell(r, plu_col) for r in body) if plu]
    if not plus:
        raise ExtractError(f"The PLU column has nothing in it in the {source}.")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows((plu, qty) for plu in plus)
    return len(plus), len(body) - len(plus)


def default_output_path(path):
    base, _ = os.path.splitext(path)
    return base + "_PLU.csv"


# ---------------------------------------------------------------- B&W flight
BW_ID_HEADER = "MYCHEM ID"
# Used when items appear before any 'MYCHEM ID' header row in the file.
BW_DEFAULT_HEADERS = [
    "MYCHEM ID", "PLU", "DESCRIPTION", "SUPPLIER", "RE-ORDER",
    "STORE TO PLACE ORDER", "MIN QTY 600MM", "MIN QTY 770MM",
]


@dataclass
class BWSection:
    title: str
    items: list = field(default_factory=list)  # one {header: value} dict per item row


def _norm_header(text):
    """'PLU:' -> 'PLU', ' mychem id ' -> 'MYCHEM ID'."""
    return re.sub(r"\s+", " ", str(text)).strip().rstrip(":").strip().upper()


def _is_number(text):
    return re.fullmatch(r"\d+(\.0+)?", text) is not None


def parse_bw_sections(rows):
    """Split B&W flight rows into sections. Sections with no items are left out."""
    sections = []
    current = None
    headers = None  # [(column index, header name)] from the latest 'MYCHEM ID' row

    for row in rows:
        cells = [str(c).strip() for c in row]
        first = next((i for i, c in enumerate(cells) if c), None)
        if first is None:
            continue  # blank (or colour-only) row
        lead = cells[first]

        if _norm_header(lead) == BW_ID_HEADER:
            headers = [(i, _norm_header(c)) for i, c in enumerate(cells) if c]
            continue

        if not _is_number(lead):
            current = BWSection(re.sub(r"\s+", " ", lead).strip())
            sections.append(current)
            continue

        if current is None:
            continue  # items before any section title: nothing to file them under
        cols = headers or list(enumerate(BW_DEFAULT_HEADERS, start=first))
        current.items.append({name: _cell(cells, i) for i, name in cols})

    return [s for s in sections if s.items]


def load_bw_sections(path):
    sections = parse_bw_sections(read_rows(path))
    if not sections:
        raise ExtractError(
            "No sections found in this file.\n\nExpected section title rows (e.g. "
            "'HEALTHY BREAK 24A BANANA BOAT') followed by rows starting with a MYCHEM ID."
        )
    return sections


def section_supplier(section):
    """The supplier shared by all of the section's items, or 'Mix' if there are several."""
    names = {}
    for item in section.items:
        name = re.sub(r"\s+", " ", item.get("SUPPLIER", "")).strip()
        if name:
            names.setdefault(name.upper(), name)
    if len(names) > 1:
        return "Mix"
    return next(iter(names.values()), "")


# ---------------------------------------------------------------- B&W order
SOH_HEADER = "SOH"
MIN_QTY_PREFIX = "MIN QTY 770"  # matches 'MIN QTY 770M' and 'MIN QTY 770MM'


def _plu_key(text):
    """PLU in the form used for matching: '0654051' and '654051.0' -> '654051'."""
    text = str(text).strip()
    if _is_number(text):
        return text.split(".")[0].lstrip("0") or "0"
    return text.upper()


def _to_number(text):
    """Number in a cell; blank counts as 0. None if it isn't a number."""
    text = str(text).strip().replace(",", "")
    if not text:
        return 0
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def _format_qty(value):
    return str(int(value)) if value == int(value) else str(value)


@dataclass
class StockTable:
    plu_col: int
    soh_col: int
    columns: dict  # {label: column index}, the columns that could be the planogram
    rows: dict  # {PLU key: row}
    conflicts: set = field(default_factory=set)  # PLU keys on several rows that differ


def parse_stock_rows(rows, source="file"):
    """Find the header row (it holds both PLU and SOH) and index the rows below by PLU."""
    rows = [[str(c).strip() for c in r] for r in rows]
    if not any(any(r) for r in rows):
        raise ExtractError(f"The {source} is empty.")

    for ri, row in enumerate(rows):
        names = [_norm_header(c) for c in row]
        if TARGET_HEADER in names and SOH_HEADER in names:
            break
    else:
        raise ExtractError(
            f"No header row with both '{TARGET_HEADER}' and '{SOH_HEADER}' columns "
            f"was found in the {source}."
        )

    plu_col, soh_col = names.index(TARGET_HEADER), names.index(SOH_HEADER)
    columns = {}
    for ci, cell in enumerate(row):
        label = re.sub(r"\s+", " ", cell)
        if not label or ci in (plu_col, soh_col):
            continue
        if label in columns:
            label = f"{label} (column {ci + 1})"
        columns[label] = ci

    table = StockTable(plu_col, soh_col, columns, {})
    for row in rows[ri + 1:]:
        key = _plu_key(_cell(row, plu_col))
        if not key or key == TARGET_HEADER:
            continue  # footer, or the header row repeated
        first = table.rows.setdefault(key, row)
        if first != row:
            table.conflicts.add(key)
    if not table.rows:
        raise ExtractError(f"No items were found under the '{TARGET_HEADER}' header in the {source}.")
    return table


# What calculate_order decided about an item
ORDER, NOT_NEEDED, FAILED = "order", "not needed", "failed"


@dataclass
class OrderItem:
    section: int  # index into the list of sections
    key: str  # PLU used for matching, "" if the flight file has none
    plu: str
    description: str
    min_qty: str = ""  # these three are display text
    planogram: str = ""
    soh: str = ""
    status: str = FAILED
    calculated: float = None  # MIN QTY 770MM + planogram - SOH
    qty: float = None  # starts as calculated; the user may change it
    include: bool = False  # goes into the order file
    notes: list = field(default_factory=list)

    @property
    def changed(self):
        return self.qty != self.calculated or self.include != (self.status == ORDER)


def calculate_order(section, table, planogram_col, threshold, index=0):
    """OrderItems for a section: QTY = MIN QTY 770MM + planogram - SOH, ordered when
    QTY >= threshold.

    A PLU listed more than once in the section is ordered once: its MIN QTYs are
    added together so that the planogram and SOH are only counted once.
    """
    items = []
    wanted = {}  # PLU key -> [OrderItem, total MIN QTY, times listed]
    for row in section.items:
        plu = row.get("PLU", "").strip()
        key = _plu_key(plu)
        item = OrderItem(
            index, key, key if _is_number(plu) else plu or "(blank)", row.get("DESCRIPTION", "")
        )
        min_header = next((h for h in row if h.startswith(MIN_QTY_PREFIX)), None)
        min_qty = _to_number(row[min_header]) if min_header else None
        if min_header:
            item.min_qty = row[min_header]
        if not key:
            item.notes.append("no PLU in the flight file")
        elif min_header is None:
            item.notes.append(f"section has no {MIN_QTY_PREFIX}MM column")
        elif min_qty is None:
            item.notes.append(f"{min_header} is not a number")
        elif key not in table.rows:
            item.notes.append("PLU not found in the stock report")
        elif key in wanted:
            wanted[key][1] += min_qty
            wanted[key][2] += 1
            continue
        else:
            wanted[key] = [item, min_qty, 1]
        items.append(item)

    for key, (item, min_qty, times) in wanted.items():
        row = table.rows[key]
        item.min_qty = _format_qty(min_qty)
        item.soh, item.planogram = _cell(row, table.soh_col), _cell(row, planogram_col)
        soh, planogram = _to_number(item.soh), _to_number(item.planogram)
        if soh is None or planogram is None:
            item.notes.append(f"{'SOH' if soh is None else 'planogram'} is not a number")
            continue
        if times > 1:
            item.notes.append(f"listed {times} times in the section; MIN QTYs added together")
        if key in table.conflicts:
            item.notes.append("on several differing stock report rows; the first was used")
        # Rounding hides float noise such as 0.30000000000000004
        item.calculated = item.qty = round(min_qty + planogram - soh, 6)
        item.status = ORDER if item.qty >= threshold else NOT_NEEDED
        item.include = item.status == ORDER
    return items


def calculate_orders(sections, table, planogram_col, threshold):
    """OrderItems for all sections, each section calculated on its own."""
    items = []
    for i, section in enumerate(sections):
        items += calculate_order(section, table, planogram_col, threshold, i)
    # The same planogram and SOH are used for every section a PLU is in: point that out.
    found_in = {}  # PLU key -> its items in the different sections
    for item in items:
        if item.status != FAILED:
            found_in.setdefault(item.key, []).append(item)
    for same in found_in.values():
        for item in same:
            others = ", ".join(sections[o.section].title for o in same if o is not item)
            if others:
                item.notes.append(f"also in: {others}")
    return items


def order_lines(items, index):
    """The (PLU, QTY) rows to write for a section."""
    return [(i.plu, _format_qty(i.qty)) for i in items if i.section == index and i.include]


def section_file_names(sections):
    """A distinct CSV file name for each section, e.g. 'HEALTHY BREAK 24A BANANA BOAT.csv'."""
    names, used = [], set()
    for section in sections:
        # Characters Windows doesn't allow in file names
        base = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", section.title)[:100].strip(" .") or "Section"
        name, n = base, 1
        while name.casefold() in used:
            n += 1
            name = f"{base} ({n})"
        used.add(name.casefold())
        names.append(name + ".csv")
    return names


def write_order(lines, out_path):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(lines)


def default_order_folder(flight_path):
    return os.path.splitext(flight_path)[0] + "_ORDER"


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


BW_STEPS = ("Sections", "Stock report", "Review")
FILE_TYPES = [("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")]
UNCHECKED, CHECKED = "☐", "☑"


@dataclass
class Page:
    frame: object
    action: object = None  # Enter key
    back: object = None  # Esc key
    drop: object = None  # called with the path of a file dropped on the window
    start: object = None  # fills the page before it is shown


def _plural(n, word):
    return f"{n:,} {word}{'' if n == 1 else 's'}"


def _one_line(text):
    return re.sub(r"\s+", " ", str(text)).strip()


def _page(parent, title, go_home, step=None):
    """A task page: header, body that takes the spare space. Returns (frame, body)."""
    from tkinter import ttk

    frm = ttk.Frame(parent, padding=16)
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(1, weight=1)
    bar = ttk.Frame(frm)
    bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
    ttk.Button(bar, text="⌂ Home", command=go_home).pack(side="left")
    ttk.Label(bar, text=title, style="PageTitle.TLabel").pack(side="left", padx=12)
    if step is not None:
        for i in reversed(range(len(BW_STEPS))):
            ttk.Label(
                bar, text=f"{i + 1}  {BW_STEPS[i]}",
                style="Step.TLabel" if i == step else "Hint.TLabel",
            ).pack(side="right")
            if i:
                ttk.Label(bar, text="→", style="Hint.TLabel").pack(side="right", padx=8)
    body = ttk.Frame(frm)
    body.grid(row=1, column=0, sticky="nsew")
    return frm, body


def _footer(frm, text, command, go_back=None):
    """Back button, and a main button that can be greyed out with the reason beside it.

    Returns (run, set_reason): run presses the main button if it is enabled and
    set_reason("") enables it.
    """
    from tkinter import ttk

    bar = ttk.Frame(frm)
    bar.grid(row=2, column=0, sticky="ew", pady=(14, 0))
    bar.columnconfigure(1, weight=1)
    if go_back:
        ttk.Button(bar, text="← Back", command=go_back).grid(row=0, column=0)
    reason_lbl = ttk.Label(bar, style="Hint.TLabel")
    reason_lbl.grid(row=0, column=1, sticky="e", padx=10)
    button = ttk.Button(bar, text=text, command=command, default="active", width=14)
    button.grid(row=0, column=2)

    def set_reason(reason):
        reason_lbl.configure(text=reason)
        button.state(["disabled" if reason else "!disabled"])

    def run():
        if not button.instate(["disabled"]):
            command()

    return run, set_reason


def _show_end(entry):
    """Scroll an entry to the end, so a long path shows its file name."""
    entry.after_idle(entry.xview_moveto, 1.0)


def _path_entry(parent, var):
    """Entry showing a chosen path, which can be copied but not typed over.

    A read-only entry would do, but Windows draws those grey.
    """
    from tkinter import ttk

    entry = ttk.Entry(parent, textvariable=var, width=56, validate="key")
    entry.configure(validatecommand=(entry.register(lambda: False),))
    return entry


def _result_dialog(root, title, message, buttons=(), warn=False):
    """Popup with extra buttons, given as (label, command). Each button closes it."""
    import tkinter as tk
    from tkinter import ttk

    win = tk.Toplevel(root)
    win.withdraw()
    win.title(APP_NAME)
    win.transient(root)
    win.resizable(False, False)
    frm = ttk.Frame(win, padding=20)
    frm.pack(fill="both", expand=True)
    ttk.Label(frm, text=title, style="DialogWarn.TLabel" if warn else "DialogOk.TLabel").pack(anchor="w")
    ttk.Label(frm, text=message, wraplength=520, justify="left").pack(anchor="w", pady=(10, 16))
    bar = ttk.Frame(frm)
    bar.pack(anchor="e")

    def press(command=None):
        win.grab_release()
        win.destroy()
        if command:
            command()

    for label, command in buttons:
        ttk.Button(bar, text=label, command=lambda c=command: press(c)).pack(side="left", padx=(0, 8))
    close = ttk.Button(bar, text="Close", command=press, default="active")
    close.pack(side="left")

    def on_return(event):
        (event.widget if event.widget.winfo_class() == "TButton" else close).invoke()

    win.bind("<Return>", on_return)
    win.bind("<Escape>", lambda e: press())
    win.protocol("WM_DELETE_WINDOW", press)
    win.update_idletasks()
    x = root.winfo_rootx() + (root.winfo_width() - win.winfo_reqwidth()) // 2
    y = root.winfo_rooty() + (root.winfo_height() - win.winfo_reqheight()) // 3
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    win.deiconify()
    win.grab_set()
    close.focus_set()
    return win


class RowsTable:
    """Read-only table of rows of text. Place .frame.

    on_pick(column number) is called when a column is clicked.
    """

    MAX_ROWS = 1000

    def __init__(self, parent, height=6, on_pick=None):
        from tkinter import font, ttk

        self.frame = ttk.Frame(parent)
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(self.frame, show="headings", height=height, selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        ys.grid(row=0, column=1, sticky="ns")
        xs = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        xs.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.font = font.nametofont("TkDefaultFont")
        self.headings = []
        self.on_pick = on_pick
        if on_pick:
            self.tree.bind("<Button-1>", self._clicked, add="+")

    def _clicked(self, event):
        if self.tree.identify_region(event.x, event.y) in ("cell", "heading"):
            column = int(self.tree.identify_column(event.x)[1:]) - 1
            if 0 <= column < len(self.headings):
                self.on_pick(column)

    def set(self, headings, rows, marked=None):
        tree = self.tree
        tree.delete(*tree.get_children())
        self.headings = headings = [_one_line(h) for h in headings]
        shown = [[_one_line(_cell(r, i)) for i in range(len(headings))] for r in rows[: self.MAX_ROWS]]
        tree.configure(columns=[f"c{i}" for i in range(len(headings))])
        for i, text in enumerate(headings):
            widest = max([r[i] for r in shown[:100]], key=len, default="")
            heading = f"✔ PLU: {text}" if self.on_pick else text
            width = max(self.font.measure(heading), self.font.measure(widest)) + 24
            tree.column(f"c{i}", width=min(max(width, 50), 300), stretch=False, anchor="w")
        self.mark(marked)
        for r in shown:
            tree.insert("", "end", values=r)
        if headings and len(rows) > len(shown):
            tree.insert("", "end", values=[f"… and {len(rows) - len(shown):,} more rows"])

    def mark(self, column):
        """Show which column is the PLU column."""
        for i, text in enumerate(self.headings):
            if i == column:
                text = "✔ PLU" if text.upper() == TARGET_HEADER else f"✔ PLU: {text}"
            self.tree.heading(f"c{i}", text=text, anchor="w")


def _raw_table(rows):
    """Headings and rows to show for cells without a recognised header row."""
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    return [f"Column {i + 1}" for i in range(max((len(r) for r in rows), default=0))], rows


def _paste_box(parent, root, on_change, to_table, height=6, on_pick=None):
    """Box for cells pasted from Excel: a text box that turns into a table after pasting.

    to_table(rows) gives the arguments for RowsTable.set. Returns
    (frame, get_text, refresh table, focus, table).
    """
    import tkinter as tk
    from tkinter import messagebox, ttk

    info = tk.StringVar()
    box = ttk.Frame(parent)
    box.columnconfigure(0, weight=1)
    box.rowconfigure(1, weight=1)
    ttk.Label(
        box, text="Copy the cells in Excel, then click in the box below and press Ctrl+V."
    ).grid(row=0, column=0, sticky="w", pady=(0, 4))

    text_frame = ttk.Frame(box)
    text_frame.columnconfigure(0, weight=1)
    text_frame.rowconfigure(0, weight=1)
    txt = tk.Text(
        text_frame, height=height, width=60, wrap="none", undo=True, relief="solid", borderwidth=1
    )
    txt.grid(row=0, column=0, sticky="nsew")
    scroll = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
    scroll.grid(row=0, column=1, sticky="ns")
    txt.configure(yscrollcommand=scroll.set)
    table = RowsTable(box, height, on_pick)
    for w in (text_frame, table.frame):
        w.grid(row=1, column=0, sticky="nsew")
    table.frame.grid_remove()
    state = {"table": False}  # whether the table is showing

    foot = ttk.Frame(box)
    foot.grid(row=2, column=0, sticky="ew", pady=(4, 0))
    foot.columnconfigure(0, weight=1)
    ttk.Label(foot, textvariable=info, style="Hint.TLabel").grid(row=0, column=0, sticky="w")

    def get_text():
        return txt.get("1.0", "end-1c")

    def parsed():
        try:
            return parse_pasted(get_text())
        except csv.Error:
            return []

    def refresh():
        if state["table"]:
            table.set(*to_table(parsed()))

    def focus():
        (table.tree if state["table"] else txt).focus_set()

    def show_table(on):
        on = on and bool(get_text().strip())
        state["table"] = on
        (text_frame if on else table.frame).grid_remove()
        (table.frame if on else text_frame).grid()
        toggle_btn.configure(text="Edit as text" if on else "Show as table")
        refresh()
        focus()

    def paste_clipboard(_=None):
        try:
            data = root.clipboard_get()
        except tk.TclError:
            messagebox.showerror(APP_NAME, "The clipboard is empty.\nCopy the cells in Excel first.")
            return
        txt.delete("1.0", "end")
        txt.insert("1.0", data)
        show_table(True)

    def clear():
        txt.delete("1.0", "end")
        show_table(False)

    toggle_btn = ttk.Button(
        foot, text="Show as table", command=lambda: show_table(not state["table"])
    )
    toggle_btn.grid(row=0, column=1)
    toggle_btn.state(["disabled"])
    ttk.Button(foot, text="Paste", command=paste_clipboard).grid(row=0, column=2, padx=(6, 0))
    ttk.Button(foot, text="Clear", command=clear).grid(row=0, column=3, padx=(6, 0))

    def modified(_=None):
        if not txt.edit_modified():
            return  # fired by the reset below
        txt.edit_modified(False)
        n = sum(1 for r in parsed() if any(c.strip() for c in r))
        info.set(f"{_plural(n, 'row')} pasted" if n else "")
        toggle_btn.state(["!disabled" if get_text().strip() else "disabled"])
        on_change()
        refresh()

    txt.bind("<<Modified>>", modified)
    txt.bind("<<Paste>>", lambda e: txt.after_idle(show_table, True))
    table.tree.bind("<<Paste>>", paste_clipboard)
    return box, get_text, refresh, focus, table


def build_home_page(parent, open_page):
    from tkinter import ttk

    page_frm = ttk.Frame(parent, padding=24)
    frm = ttk.Frame(page_frm)
    frm.pack(expand=True)  # centred
    ttk.Label(frm, text=APP_NAME, style="AppTitle.TLabel").pack(anchor="w")
    ttk.Label(frm, text="What would you like to do?").pack(anchor="w", pady=(2, 16))

    for page, label, blurb in (
        ("plu", "Generate PLU file",
         "Turn an Excel/CSV file or pasted cells into a PLU, QTY CSV."),
        ("bw", "Process B&W Flight",
         "Work out what to order for a B&W flight from a stock report."),
    ):
        ttk.Button(frm, text=label, style="Big.TButton", command=lambda p=page: open_page(p)).pack(
            fill="x", pady=(8, 0)
        )
        ttk.Label(frm, text=blurb, style="Hint.TLabel").pack(anchor="w", pady=(2, 4))
    return Page(page_frm)


def build_plu_page(parent, root, go_home):
    import subprocess
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    frm, body = _page(parent, "Generate PLU file", go_home)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(2, weight=1)  # the preview takes the spare height

    mode = tk.StringVar(value="file")  # "file" or "paste"
    in_var = tk.StringVar()
    out_var = tk.StringVar()
    has_header = tk.BooleanVar(value=False)
    plu_var = tk.StringVar()
    qty_var = tk.StringVar(value="1")
    # For each input: the rows read, the columns [(index, label)] and rows below the header
    # found in them, and the index of the PLU column (None: not chosen)
    sources = {m: {"rows": [], "columns": [], "body": [], "plu": None} for m in ("file", "paste")}
    # path: the file the preview shows; error: why it couldn't be read
    preview = {"path": "", "error": "", "timer": None}

    def auto_output():
        """The output path suggested for the current input ('' if there isn't one)."""
        if mode.get() == "file" and in_var.get().strip():
            return os.path.normpath(default_output_path(in_var.get().strip()))
        return ""

    # --- input source switch
    ttk.Label(body, text="Input:").grid(row=1, column=0, sticky="w")
    src_row = ttk.Frame(body)
    src_row.grid(row=1, column=1, columnspan=2, sticky="w", padx=6)

    # --- PLU column
    def set_rows(which, rows):
        """Take new rows for an input, going back to the assumed PLU column."""
        src = sources[which]
        src["rows"] = rows
        src["columns"], src["body"], src["plu"] = plu_table(rows, has_header.get())

    def to_table(which):
        """The preview of an input."""
        src = sources[which]
        return (
            [label for _, label in src["columns"]],
            [[_cell(r, i) for i, _ in src["columns"]] for r in src["body"]],
            plu_position(which),
        )

    def plu_position(which):
        """Where the PLU column is among the columns shown."""
        src = sources[which]
        return next((n for n, (i, _) in enumerate(src["columns"]) if i == src["plu"]), None)

    def pick(position):
        """Make a column of the current preview the PLU column."""
        src = sources[mode.get()]
        src["plu"] = src["columns"][position][0]
        show_plu()

    def show_plu():
        """Update the dropdown, the mark in the preview and the Convert button."""
        which = mode.get()
        labels = [label for _, label in sources[which]["columns"]]
        position = plu_position(which)
        plu_box.configure(values=labels, state="readonly" if labels else "disabled")
        plu_var.set("" if position is None else labels[position])
        (file_rows if which == "file" else paste_rows).mark(position)
        update_ready()

    # --- file input
    file_box = ttk.Frame(body)
    file_box.columnconfigure(1, weight=1)
    file_box.rowconfigure(1, weight=1)

    def use_file(path):
        mode.set("file")
        in_var.set(os.path.normpath(path))
        out_var.set(auto_output())
        load_preview()
        switch_mode()

    def browse_input():
        path = filedialog.askopenfilename(title="Select an Excel or CSV file", filetypes=FILE_TYPES)
        if path:
            use_file(path)

    ttk.Label(file_box, text="Input file:").grid(row=0, column=0, sticky="w")
    in_entry = ttk.Entry(file_box, textvariable=in_var, width=48)
    in_entry.grid(row=0, column=1, sticky="ew", padx=6)
    ttk.Button(file_box, text="Browse…", command=browse_input).grid(row=0, column=2)
    file_rows = RowsTable(file_box, on_pick=pick)
    file_rows.frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
    file_info = ttk.Label(file_box)
    file_info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def load_preview():
        """Read the chosen file and show its rows."""
        if preview["timer"]:
            root.after_cancel(preview["timer"])
        path = in_var.get().strip()
        rows, error, text = [], "", ""
        if os.path.isfile(path):
            try:
                rows = [r for r in read_rows(path) if any(str(c).strip() for c in r)]
                text = f"{_plural(len(rows), 'row')} found in the file."
            except PermissionError:
                error = "Could not open the file. Is it open in Excel?"
            except Exception as e:  # noqa: BLE001
                error = f"Could not read the file: {_one_line(e)}"
        elif root.can_drop:
            text = "You can also drop a file onto this window."
        preview.update(path=path, error=error, timer=None)
        file_info.configure(text=error or text, style="Warn.TLabel" if error else "Hint.TLabel")
        set_rows("file", rows)
        file_rows.set(*to_table("file"))
        show_plu()

    def input_changed(*_):
        # Wait for a pause in typing before reading the file
        if preview["timer"]:
            root.after_cancel(preview["timer"])
        preview["timer"] = root.after(250, load_preview)
        update_ready()

    # --- paste input
    def on_paste_change():
        try:
            rows = parse_pasted(pasted_text())
        except csv.Error:
            rows = []
        set_rows("paste", rows)
        show_plu()

    paste_box, pasted_text, refresh_paste, focus_paste, paste_rows = _paste_box(
        body, root, on_paste_change, lambda rows: to_table("paste"), on_pick=pick
    )

    def switch_mode():
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
            focus_paste()
        show_plu()

    for value, label in (("file", "Select file"), ("paste", "Paste from Excel")):
        ttk.Radiobutton(
            src_row, text=label, value=value, variable=mode, command=switch_mode
        ).pack(side="left", padx=(0, 12))

    for box in (file_box, paste_box):
        box.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(10, 0))

    # --- options
    def header_changed():
        for which, src in sources.items():
            set_rows(which, src["rows"])
        file_rows.set(*to_table("file"))
        refresh_paste()
        show_plu()

    ttk.Checkbutton(
        body, text="Data has a header row", variable=has_header, command=header_changed
    ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    ttk.Label(body, text="PLU column:").grid(row=4, column=0, sticky="w", pady=(8, 0))
    plu_row = ttk.Frame(body)
    plu_row.grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=(8, 0))
    plu_box = ttk.Combobox(plu_row, textvariable=plu_var, state="disabled", width=30, height=15)
    plu_box.pack(side="left")
    plu_box.bind("<<ComboboxSelected>>", lambda e: pick(plu_box.current()))
    ttk.Label(plu_row, text="or click a column in the preview", style="Hint.TLabel").pack(
        side="left", padx=8
    )
    ttk.Label(body, text="QTY:").grid(row=5, column=0, sticky="w", pady=(8, 0))
    qty_row = ttk.Frame(body)
    qty_row.grid(row=5, column=1, columnspan=2, sticky="w", padx=6, pady=(8, 0))
    ttk.Spinbox(qty_row, textvariable=qty_var, from_=1, to=99999, increment=1, width=8).pack(
        side="left"
    )
    ttk.Label(qty_row, text="used for every PLU", style="Hint.TLabel").pack(side="left", padx=8)

    ttk.Separator(body).grid(row=6, column=0, columnspan=3, sticky="ew", pady=12)

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

    out_row = ttk.Frame(body)
    out_row.grid(row=7, column=0, columnspan=3, sticky="ew")
    out_row.columnconfigure(1, weight=1)
    ttk.Label(out_row, text="Output file:").grid(row=0, column=0, sticky="w")
    out_entry = ttk.Entry(out_row, textvariable=out_var, width=48)
    out_entry.grid(row=0, column=1, sticky="ew", padx=6)
    ttk.Button(out_row, text="Save as…", command=browse_output).grid(row=0, column=2)

    # When a path is filled in for the user, keep the file name (end of a long path) visible
    for var, entry in ((in_var, in_entry), (out_var, out_entry)):
        var.trace_add(
            "write", lambda *_, e=entry: None if e.focus_get() is e else _show_end(e)
        )

    # --- action + result
    def convert():
        qty = qty_var.get().strip()
        pasting = mode.get() == "paste"
        text, src = pasted_text(), in_var.get().strip()
        plu_col = sources[mode.get()]["plu"]

        out = out_var.get().strip()
        confirmed = False
        if not out:
            if pasting:
                # No input file to put the output next to: ask where to save.
                # (The Save dialog asks about replacing an existing file itself.)
                if not browse_output():
                    return
                out = out_var.get().strip()
                confirmed = True
            else:
                out = auto_output()
                out_var.set(out)

        if not pasting and os.path.abspath(out) == os.path.abspath(src):
            messagebox.showerror(APP_NAME, "The output file can't be the same as the input file.")
            return
        if not confirmed and os.path.exists(out) and not messagebox.askyesno(
            APP_NAME, f"{os.path.basename(out)} already exists.\nReplace it?"
        ):
            return

        try:
            if pasting:
                n, skipped = extract_rows(
                    parse_pasted(text), out, has_header.get(), qty, plu_col, source="pasted data"
                )
            else:
                n, skipped = extract(src, out, has_header.get(), qty, plu_col)
        except ExtractError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except PermissionError:
            messagebox.showerror(APP_NAME, "Could not write the output file.\nIs it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Unexpected error:\n{e}")
            return
        message = out
        if skipped:
            message += f"\n\n{_plural(skipped, 'row')} without a PLU left out."
        _result_dialog(
            root, f"File generated with {_plural(n, 'PLU')}", message,
            [("Open folder", lambda: subprocess.Popen(["explorer", "/select,", out]))],
        )

    run, set_reason = _footer(frm, "Convert", convert)

    def update_ready(*_):
        src = in_var.get().strip()
        reason = ""
        if mode.get() == "paste":
            if not pasted_text().strip():
                reason = "Paste the cells from Excel"
        elif not src:
            reason = "Choose an input file"
        elif not os.path.isfile(src):
            reason = "Input file not found"
        elif preview["path"] != src:
            reason = "Reading the file…"
        elif preview["error"]:
            reason = "The file could not be read"
        elif not sources["file"]["columns"]:
            reason = "The file is empty"
        if not reason and sources[mode.get()]["plu"] is None:
            reason = "Choose the PLU column"
        if not reason:
            try:
                if not math.isfinite(float(qty_var.get().strip())):
                    raise ValueError
            except ValueError:
                reason = "QTY must be a number"
        set_reason(reason)

    in_var.trace_add("write", input_changed)
    qty_var.trace_add("write", update_ready)
    load_preview()
    switch_mode()
    return Page(frm, run, go_home, use_file)


def build_bw_page(parent, root, go_home, go_next):
    """go_next(flight path, chosen sections) opens the stock report page."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    frm, body = _page(parent, "Process B&W Flight", go_home, step=0)
    body.columnconfigure(0, weight=1)
    body.rowconfigure(3, weight=1)

    ALL_SUPPLIERS, NO_SUPPLIER = "All suppliers", "(no supplier)"
    in_var = tk.StringVar()
    found_var = tk.StringVar(
        value="Choose a B&W flight file to list its sections"
        + (", or drop it onto this window." if root.can_drop else ".")
    )
    sel_info = tk.StringVar()
    search_var = tk.StringVar()
    supplier_var = tk.StringVar(value=ALL_SUPPLIERS)
    # selected holds indexes into sections; sort is (column or None, descending);
    # shown lists the indexes that pass the filter, in display order
    state = {"sections": [], "selected": set(), "sort": (None, False), "shown": []}

    # --- input file
    file_row = ttk.Frame(body)
    file_row.grid(row=0, column=0, sticky="ew")
    file_row.columnconfigure(1, weight=1)
    ttk.Label(file_row, text="Flight file:").grid(row=0, column=0, sticky="w")
    in_entry = _path_entry(file_row, in_var)
    in_entry.grid(row=0, column=1, sticky="ew", padx=6)

    # --- sections list
    ttk.Label(body, textvariable=found_var).grid(row=1, column=0, sticky="w", pady=(12, 4))

    filter_row = ttk.Frame(body)
    filter_row.grid(row=2, column=0, sticky="ew", pady=(0, 6))
    filter_row.columnconfigure(1, weight=1)
    ttk.Label(filter_row, text="Search:").grid(row=0, column=0)
    search_entry = ttk.Entry(filter_row, textvariable=search_var)
    search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 12))
    search_entry.bind("<Return>", lambda e: "break")  # filter only, don't go to the next page
    ttk.Label(filter_row, text="Supplier:").grid(row=0, column=2)
    supplier_box = ttk.Combobox(
        filter_row, textvariable=supplier_var, state="readonly", width=24, values=(ALL_SUPPLIERS,)
    )
    supplier_box.grid(row=0, column=3, padx=(6, 0))

    list_box = ttk.Frame(body)
    list_box.grid(row=3, column=0, sticky="nsew")
    list_box.columnconfigure(0, weight=1)
    list_box.rowconfigure(0, weight=1)
    tree = ttk.Treeview(
        list_box, columns=("check", "section", "supplier", "items"), show="headings", height=8,
        selectmode="none",
    )
    tree.heading("check", text="")
    headings = {"section": ("Section", "w"), "supplier": ("Supplier", "w"), "items": ("Items", "center")}
    tree.column("check", width=36, minwidth=36, stretch=False, anchor="center")
    tree.column("section", width=420, anchor="w")
    tree.column("supplier", width=170, stretch=False, anchor="w")
    tree.column("items", width=80, stretch=False, anchor="center")
    tree.grid(row=0, column=0, sticky="nsew")
    tree_scroll = ttk.Scrollbar(list_box, orient="vertical", command=tree.yview)
    tree_scroll.grid(row=0, column=1, sticky="ns")
    tree.configure(yscrollcommand=tree_scroll.set)
    tree.tag_configure("on", background="#e6f0fb")

    def refresh_row(idx):
        on = idx in state["selected"]
        tree.set(str(idx), "check", CHECKED if on else UNCHECKED)
        tree.item(str(idx), tags=("on",) if on else ())

    def update_sel_info():
        secs, sel = state["sections"], state["selected"]
        if not secs:
            sel_info.set("")
            set_reason("Choose a flight file")
            return
        items = sum(len(secs[i].items) for i in sel)
        hidden = len(sel - set(state["shown"]))
        sel_info.set(
            f"{len(sel)} of {_plural(len(secs), 'section')} selected"
            + (f" ({_plural(items, 'item')})" if sel else "")
            + (f", including {hidden} hidden by the filter" if hidden else "")
        )
        set_reason("" if sel else "Select at least one section")

    def toggle(idx):
        state["selected"].symmetric_difference_update({idx})
        refresh_row(idx)
        update_sel_info()

    def sort_key(col, idx):
        section = state["sections"][idx]
        if col == "items":
            return len(section.items)
        text = section.title if col == "section" else section_supplier(section)
        # Numbers in the text sort by value, so 'BREAK 9' comes before 'BREAK 24A'
        return [
            (0, int(part), "") if part.isdecimal() else (1, 0, part)
            for part in re.split(r"(\d+)", text.upper()) if part
        ]

    def passes_filter(section):
        supplier = supplier_var.get()
        if supplier != ALL_SUPPLIERS and supplier.upper() != (
            section_supplier(section) or NO_SUPPLIER
        ).upper():
            return False
        return all(word in section.title.upper() for word in search_var.get().upper().split())

    def refresh_view(*_):
        """List the sections that pass the filter, in the chosen order (default: as in the file)."""
        col, descending = state["sort"]
        order = [i for i, s in enumerate(state["sections"]) if passes_filter(s)]
        if col:
            order.sort(key=lambda i: sort_key(col, i), reverse=descending)
        state["shown"] = order
        tree.set_children("", *map(str, order))
        for name, (label, anchor) in headings.items():
            arrow = (" ▼" if descending else " ▲") if name == col else ""
            tree.heading(name, text=label + arrow, anchor=anchor)
        update_sel_info()

    def sort_by(col):
        prev, descending = state["sort"]
        state["sort"] = (col, col == prev and not descending)
        refresh_view()

    for name in headings:
        tree.heading(name, command=lambda c=name: sort_by(c))

    def on_click(event):
        if tree.identify_region(event.x, event.y) == "heading":
            return None  # let the heading handle it (sorts the list)
        row = tree.identify_row(event.y)
        if row and tree.identify_region(event.x, event.y) == "cell":
            tree.focus(row)
            tree.focus_set()
            toggle(int(row))
        return "break"

    def on_space(_):
        row = tree.focus()
        if row and int(row) in state["shown"]:
            toggle(int(row))
        return "break"

    tree.bind("<Button-1>", on_click)
    tree.bind("<space>", on_space)

    def set_all(on):
        """Select or clear the sections that are showing."""
        (state["selected"].update if on else state["selected"].difference_update)(state["shown"])
        for i in state["shown"]:
            refresh_row(i)
        update_sel_info()

    sel_row = ttk.Frame(body)
    sel_row.grid(row=4, column=0, sticky="ew", pady=(6, 0))
    sel_row.columnconfigure(2, weight=1)
    ttk.Button(sel_row, text="Select all", command=lambda: set_all(True)).grid(row=0, column=0)
    ttk.Button(sel_row, text="Select none", command=lambda: set_all(False)).grid(
        row=0, column=1, padx=(6, 0)
    )
    ttk.Label(sel_row, textvariable=sel_info).grid(row=0, column=2, sticky="e")

    # --- loading
    def load(path):
        try:
            sections = load_bw_sections(path)
        except ExtractError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except PermissionError:
            messagebox.showerror(APP_NAME, "Could not open the file.\nIs it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Could not read the file:\n{e}")
            return
        in_var.set(os.path.normpath(path))
        _show_end(in_entry)
        # Rows hidden by the filter are still items of the tree, so clear by id
        tree.delete(*(str(i) for i in range(len(state["sections"]))))
        state["sections"], state["sort"] = sections, (None, False)
        state["selected"] = set(range(len(sections)))
        suppliers = {}
        for i, s in enumerate(sections):
            supplier = section_supplier(s)
            tree.insert("", "end", iid=str(i), values=("", s.title, supplier, len(s.items)))
            refresh_row(i)
            suppliers.setdefault((supplier or NO_SUPPLIER).upper(), supplier or NO_SUPPLIER)
        supplier_box.configure(values=(ALL_SUPPLIERS, *(suppliers[k] for k in sorted(suppliers))))
        search_var.set("")
        supplier_var.set(ALL_SUPPLIERS)
        found_var.set(
            f"{_plural(len(sections), 'section')} found. "
            "Click a section to include or leave it out, or a heading to sort:"
        )
        refresh_view()
        tree.focus_set()

    def browse_input():
        path = filedialog.askopenfilename(title="Select the B&W flight file", filetypes=FILE_TYPES)
        if path:
            load(path)

    ttk.Button(file_row, text="Browse…", command=browse_input).grid(row=0, column=2)

    run, set_reason = _footer(
        frm, "Next →",
        lambda: go_next(in_var.get(), [state["sections"][i] for i in sorted(state["selected"])]),
    )
    for var in (search_var, supplier_var):
        var.trace_add("write", refresh_view)
    refresh_view()
    return Page(frm, run, go_home, load)


def build_bw_stock_page(parent, root, go_home, go_back, go_next):
    """go_next(stock table, planogram column label, threshold, stock file path) opens the
    review page.

    Page.start takes (flight path, sections).
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    frm, body = _page(parent, "Process B&W Flight", go_home, step=1)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(3, weight=1)

    mode = tk.StringVar(value="file")  # "file" or "paste"
    chosen_info = tk.StringVar()
    in_var = tk.StringVar()
    stock_info = tk.StringVar()
    planogram_var = tk.StringVar()
    threshold_var = tk.StringVar(value="0")
    state = {
        "file_table": None,  # StockTable from the chosen file
        "paste_table": None,  # StockTable from the pasted cells
        "paste_error": "",
    }

    ttk.Label(body, textvariable=chosen_info).grid(row=1, column=0, columnspan=3, sticky="w")

    # --- stock report source
    ttk.Label(body, text="Stock report:").grid(row=2, column=0, sticky="w", pady=(10, 0))
    src_row = ttk.Frame(body)
    src_row.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def table_rows(table):
        """What to show for a StockTable: the labelled columns, in report order."""
        cols = {table.plu_col: TARGET_HEADER, table.soh_col: SOH_HEADER}
        cols.update((ci, label) for label, ci in table.columns.items())
        order = sorted(cols)
        return (
            [cols[c] for c in order],
            [[_cell(row, c) for c in order] for row in table.rows.values()],
        )

    file_box = ttk.Frame(body)
    file_box.columnconfigure(1, weight=1)
    file_box.rowconfigure(1, weight=1)
    ttk.Label(file_box, text="Stock file:").grid(row=0, column=0, sticky="w")
    in_entry = _path_entry(file_box, in_var)
    in_entry.grid(row=0, column=1, sticky="ew", padx=6)
    file_rows = RowsTable(file_box, height=5)
    file_rows.frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(8, 0))

    def browse_input():
        path = filedialog.askopenfilename(title="Select the stock report", filetypes=FILE_TYPES)
        if path:
            load(path)

    def load(path):
        try:
            table = parse_stock_rows(read_rows(path))
        except ExtractError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except PermissionError:
            messagebox.showerror(APP_NAME, "Could not open the file.\nIs it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Could not read the file:\n{e}")
            return
        in_var.set(os.path.normpath(path))
        _show_end(in_entry)
        state["file_table"] = table
        file_rows.set(*table_rows(table))
        mode.set("file")
        switch_mode()

    ttk.Button(file_box, text="Browse…", command=browse_input).grid(row=0, column=2)

    def on_paste_change():
        state["paste_table"], state["paste_error"] = None, ""
        text = pasted_text()
        if text.strip():
            try:
                state["paste_table"] = parse_stock_rows(parse_pasted(text), "pasted data")
            except (ExtractError, csv.Error) as e:
                state["paste_error"] = str(e)
        refresh_table()

    paste_box, pasted_text, _, focus_paste, _ = _paste_box(
        body, root, on_paste_change,
        lambda rows: table_rows(state["paste_table"]) if state["paste_table"] else _raw_table(rows),
        height=5,
    )

    for box in (file_box, paste_box):
        box.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
    paste_box.grid_remove()

    def current_table():
        return state["file_table"] if mode.get() == "file" else state["paste_table"]

    def switch_mode():
        if mode.get() == "file":
            paste_box.grid_remove()
            file_box.grid()
        else:
            file_box.grid_remove()
            paste_box.grid()
            focus_paste()
        refresh_table()

    for value, label in (("file", "Select file"), ("paste", "Paste from Excel")):
        ttk.Radiobutton(
            src_row, text=label, value=value, variable=mode, command=switch_mode
        ).pack(side="left", padx=(0, 12))

    ttk.Label(body, textvariable=stock_info, style="Hint.TLabel").grid(
        row=4, column=0, columnspan=3, sticky="w", pady=(6, 0)
    )

    ttk.Separator(body).grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

    # --- calculation
    ttk.Label(body, text="Planogram column:").grid(row=6, column=0, sticky="w")
    planogram_box = ttk.Combobox(
        body, textvariable=planogram_var, state="disabled", width=36, height=15
    )
    planogram_box.grid(row=6, column=1, sticky="w", padx=6)
    ttk.Label(body, text="Threshold:").grid(row=8, column=0, sticky="w", pady=(8, 0))
    ttk.Spinbox(body, textvariable=threshold_var, from_=-99999, to=99999, increment=1, width=8).grid(
        row=8, column=1, sticky="w", padx=6, pady=(8, 0)
    )
    ttk.Label(
        body, text="An item is ordered when  MIN QTY 770MM + Planogram − SOH  ≥  threshold.",
        style="Hint.TLabel",
    ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def refresh_table():
        """Show the current stock report's item count and its columns to pick from."""
        table = current_table()
        if table is None:
            if mode.get() == "paste":
                stock_info.set(state["paste_error"])
            else:
                stock_info.set(
                    "You can also drop the file onto this window." if root.can_drop else ""
                )
            planogram_box.configure(values=(), state="disabled")
            planogram_var.set("")
        else:
            stock_info.set(f"{_plural(len(table.rows), 'item')} found in the stock report.")
            labels = list(table.columns)
            planogram_box.configure(values=labels, state="readonly")
            if planogram_var.get() not in table.columns:
                likely = [c for c in labels if "PLANOGRAM" in c.upper()]
                planogram_var.set(likely[0] if len(likely) == 1 else "")
        update_ready()

    def update_ready(*_):
        """Enable Next, or say what is missing."""
        table, reason = current_table(), ""
        if table is None:
            reason = "Choose a stock report file" if mode.get() == "file" else "Paste the stock report"
        elif planogram_var.get() not in table.columns:
            reason = "Select the planogram column"
        threshold = threshold_var.get().strip()
        if not reason and (not threshold or _to_number(threshold) is None):
            reason = "Threshold must be a number"
        set_reason(reason)

    run, set_reason = _footer(
        frm, "Next →",
        lambda: go_next(
            current_table(), planogram_var.get(), _to_number(threshold_var.get()),
            in_var.get() if mode.get() == "file" else "",
        ),
        go_back,
    )
    for var in (planogram_var, threshold_var):
        var.trace_add("write", update_ready)
    refresh_table()

    def start(flight, sections):
        chosen_info.set(
            f"{_plural(len(sections), 'section')} selected from {os.path.basename(flight)} "
            f"({_plural(sum(len(s.items) for s in sections), 'item')})"
        )

    return Page(frm, run, go_back, load, start)


def build_bw_review_page(parent, root, go_home, go_back):
    """Page.start takes (flight path, sections, stock table, planogram label, threshold,
    stock file path) and returns False if the user chose to stay on the previous page."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    frm, body = _page(parent, "Process B&W Flight", go_home, step=2)
    body.columnconfigure(0, weight=1)
    body.rowconfigure(1, weight=1)

    FILTERS = {
        "order": ("To order", lambda i: i.status == ORDER or i.include),
        "not needed": ("Not needed", lambda i: i.status == NOT_NEEDED),
        "failed": ("Couldn't calculate", lambda i: i.status == FAILED),
        "check": ("To check", lambda i: i.status != FAILED and i.notes),
        "all": ("All", lambda i: True),
    }
    filter_var = tk.StringVar(value="order")
    out_var = tk.StringVar()
    total_var = tk.StringVar()
    hint_var = tk.StringVar()
    HINT = "Click a box to include or leave out an item. Click an order quantity to change it."
    state = {
        "flight": "", "sections": [], "inputs": None, "items": [],
        "sources": [],  # input files, which must not be overwritten
        "editor": None,  # (entry widget, item index) while a quantity is being changed
    }

    # --- filter
    filter_row = ttk.Frame(body)
    filter_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
    ttk.Label(filter_row, text="Show:").pack(side="left", padx=(0, 8))
    filter_btns = {}
    for name in FILTERS:
        filter_btns[name] = ttk.Radiobutton(
            filter_row, value=name, variable=filter_var, command=lambda: refresh()
        )
        filter_btns[name].pack(side="left", padx=(0, 12))

    # --- items list
    list_box = ttk.Frame(body)
    list_box.grid(row=1, column=0, sticky="nsew")
    list_box.columnconfigure(0, weight=1)
    list_box.rowconfigure(0, weight=1)
    columns = (
        ("check", "", 36, "center"), ("section", "Section", 150, "w"), ("plu", "PLU", 65, "w"),
        ("desc", "Description", 200, "w"), ("min", "MIN QTY", 65, "center"),
        ("plano", "Planogram", 75, "center"), ("soh", "SOH", 50, "center"),
        ("qty", "Order", 60, "center"), ("note", "Notes", 235, "w"),
    )
    tree = ttk.Treeview(
        list_box, columns=[c[0] for c in columns], show="headings", height=8, selectmode="browse"
    )
    for name, label, width, anchor in columns:
        tree.heading(name, text=label, anchor=anchor)
        tree.column(name, width=width, minwidth=36, stretch=name == "note", anchor=anchor)
    tree.grid(row=0, column=0, sticky="nsew")
    ys = ttk.Scrollbar(list_box, orient="vertical", command=tree.yview)
    ys.grid(row=0, column=1, sticky="ns")
    xs = ttk.Scrollbar(list_box, orient="horizontal", command=tree.xview)
    xs.grid(row=1, column=0, sticky="ew")

    def scrolled(bar, *pos):
        bar.set(*pos)
        finish_edit()  # the entry doesn't scroll with the list

    tree.configure(
        yscrollcommand=lambda *p: scrolled(ys, *p), xscrollcommand=lambda *p: scrolled(xs, *p)
    )
    tree.tag_configure("on", background="#e6f0fb")
    tree.tag_configure("off", foreground="#80868b")

    info_row = ttk.Frame(body)
    info_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
    info_row.columnconfigure(0, weight=1)
    hint_lbl = ttk.Label(info_row, textvariable=hint_var, style="Hint.TLabel")
    hint_lbl.grid(row=0, column=0, sticky="w")
    ttk.Label(info_row, textvariable=total_var).grid(row=0, column=1, sticky="e")

    def hint(text="", warn=False):
        hint_var.set(text or HINT)
        hint_lbl.configure(style="Warn.TLabel" if warn else "Hint.TLabel")

    def row_values(item):
        notes = list(item.notes)
        if item.qty != item.calculated:
            was = "blank" if item.calculated is None else _format_qty(item.calculated)
            notes.insert(0, f"changed from {was}")
        return (
            CHECKED if item.include else UNCHECKED,
            state["sections"][item.section].title, item.plu, _one_line(item.description),
            item.min_qty, item.planogram, item.soh,
            "" if item.qty is None else _format_qty(item.qty), "; ".join(notes),
        )

    def refresh_row(idx):
        item = state["items"][idx]
        tree.item(str(idx), values=row_values(item), tags=("on" if item.include else "off",))

    def update_totals():
        items = state["items"]
        for name, (label, test) in FILTERS.items():
            filter_btns[name].configure(text=f"{label} ({sum(1 for i in items if test(i))})")
        chosen = [i for i in items if i.include]
        files = len({i.section for i in chosen})
        total_var.set(f"{_plural(len(chosen), 'item')} to order in {_plural(files, 'file')}")
        if not chosen:
            set_reason("Tick at least one item")
        elif not out_var.get().strip():
            set_reason("Choose an output folder")
        else:
            set_reason("")

    def refresh():
        """List the items of the chosen filter."""
        finish_edit()
        tree.delete(*tree.get_children())
        test = FILTERS[filter_var.get()][1]
        for idx, item in enumerate(state["items"]):
            if test(item):
                tree.insert("", "end", iid=str(idx))
                refresh_row(idx)
        tree.yview_moveto(0)
        update_totals()

    # --- ticking items and changing quantities
    def toggle(idx):
        item = state["items"][idx]
        if item.qty is None:
            start_edit(idx)  # nothing to order until there is a quantity
            return
        hint()
        item.include = not item.include
        refresh_row(idx)
        update_totals()

    def start_edit(idx):
        finish_edit()
        item = state["items"][idx]
        if not item.key:
            hint("This item has no PLU in the flight file, so it can't be ordered.", True)
            return
        tree.see(str(idx))
        tree.update_idletasks()
        bbox = tree.bbox(str(idx), "qty")
        if not bbox:
            return  # scrolled out of sight sideways
        x, y, w, h = bbox
        entry = ttk.Entry(tree, justify="center")
        entry.insert(0, "" if item.qty is None else _format_qty(item.qty))
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=w, height=h)
        state["editor"] = (entry, idx)
        hint("Type the quantity and press Enter. Clear it to leave the item out.")

        def done(apply):
            finish_edit(apply)
            tree.focus_set()
            return "break"

        def left():
            # Only if the focus really moved to something else in this window
            if state["editor"] and state["editor"][0] is entry and root.focus_get() not in (None, entry):
                finish_edit()

        entry.bind("<Return>", lambda e: done(True))
        entry.bind("<Tab>", lambda e: done(True))
        entry.bind("<Escape>", lambda e: done(False))
        entry.bind("<FocusOut>", lambda e: root.after(50, left))
        entry.focus_set()

    def finish_edit(apply=True):
        """Close the quantity entry if it is open. False if it held something unusable."""
        if state["editor"] is None:
            return True
        entry, idx = state["editor"]
        state["editor"] = None  # first: destroying the entry fires FocusOut
        text = entry.get().strip()
        entry.destroy()
        hint()
        item = state["items"][idx]
        if not apply:
            return True
        value = _to_number(text)
        if value is None:
            hint(f"'{text}' isn't a number, so the quantity wasn't changed.", True)
            return False
        if text:
            value = round(value, 6)
            if value != item.qty or not item.include:
                item.qty, item.include = value, True
        else:
            item.include = False
        if tree.exists(str(idx)):
            refresh_row(idx)
        update_totals()
        return True

    def on_click(event):
        region = tree.identify_region(event.x, event.y)
        row = tree.identify_row(event.y)
        if region != "cell" or not row:
            return None
        finish_edit()
        col = columns[int(tree.identify_column(event.x)[1:]) - 1][0]
        tree.selection_set(row)
        tree.focus(row)
        tree.focus_set()
        if col == "check":
            toggle(int(row))
        elif col == "qty":
            start_edit(int(row))
        return "break"

    def on_key(action):
        def handler(_):
            row = tree.focus()
            if row and tree.exists(row):
                action(int(row))
            return "break"
        return handler

    tree.bind("<Button-1>", on_click)
    tree.bind("<space>", on_key(toggle))
    tree.bind("<Return>", on_key(start_edit))
    tree.bind("<F2>", on_key(start_edit))
    tree.bind("<Configure>", lambda e: finish_edit())

    # --- output folder
    def browse_output():
        current = out_var.get().strip() or default_order_folder(state["flight"])
        path = filedialog.askdirectory(
            title="Choose the folder for the order files",
            initialdir=current if os.path.isdir(current) else os.path.dirname(current) or None,
        )
        if path:
            out_var.set(os.path.normpath(path))
            _show_end(out_entry)

    out_row = ttk.Frame(body)
    out_row.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    out_row.columnconfigure(1, weight=1)
    ttk.Label(out_row, text="Output folder:").grid(row=0, column=0, sticky="w")
    out_entry = ttk.Entry(out_row, textvariable=out_var, width=56)
    out_entry.grid(row=0, column=1, sticky="ew", padx=6)
    # Keep the folder name (end of a long path) visible once the page is laid out
    out_entry.bind("<Configure>", lambda e: _show_end(out_entry))
    ttk.Button(out_row, text="Browse…", command=browse_output).grid(row=0, column=2)
    ttk.Label(
        out_row, text="Each section is saved as its own CSV in this folder.", style="Hint.TLabel"
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

    # --- action + result
    def show_details():
        items = state["items"]
        for name in ("failed", "check", "not needed"):
            if any(FILTERS[name][1](i) for i in items):
                break
        else:
            name = "all"
        filter_var.set(name)
        refresh()

    def generate():
        if not finish_edit():
            return
        sections, items = state["sections"], state["items"]
        if not any(i.include for i in items):
            return  # the edit just finished left nothing to order
        folder = out_var.get().strip()
        lines = [order_lines(items, i) for i in range(len(sections))]
        paths = [os.path.join(folder, name) for name in section_file_names(sections)]
        todo = [(path, rows) for path, rows in zip(paths, lines) if rows]
        # Files from an earlier run for sections that now have nothing to order
        stale = [p for p, rows in zip(paths, lines) if not rows and os.path.exists(p)]

        def full(path):
            return os.path.normcase(os.path.abspath(path))

        clash = {full(p) for p, _ in todo} & {full(p) for p in state["sources"]}
        if clash:
            messagebox.showerror(
                APP_NAME,
                "An order file would replace an input file. Choose another output folder.\n\n"
                + "\n".join(sorted(clash)),
            )
            return
        existing = sum(os.path.exists(p) for p, _ in todo)
        if existing and not messagebox.askyesno(
            APP_NAME,
            f"{_plural(existing, 'order file')} already in this folder will be replaced.\n"
            f"Continue?\n\n{folder}",
        ):
            return

        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Could not create the output folder:\n{e}")
            return
        written, failed = [], []  # written: rows of each file; failed: messages
        for path, rows in todo:
            try:
                write_order(rows, path)
                written.append(rows)
            except PermissionError:
                failed.append(f"{os.path.basename(path)} (is it open in Excel?)")
            except OSError as e:
                failed.append(f"{os.path.basename(path)} ({e.strerror or e})")

        def count(test):
            return sum(1 for i in items if test(i))

        parts = [
            f"{n} {label.lower()}" for n, label in (
                (count(lambda i: i.status == NOT_NEEDED and not i.include), "not needed"),
                (count(lambda i: i.status == FAILED and not i.include), "couldn't be calculated"),
                (count(lambda i: i.status == ORDER and not i.include), "left out by you"),
                (count(lambda i: i.include and i.changed), "changed or added by you"),
                (count(FILTERS["check"][1]), "to check"),
            ) if n
        ]
        message = folder
        if parts:
            message += f"\n\nItems: {', '.join(parts)}."
        empty = sum(not rows for rows in lines)
        if empty:
            message += f"\n\n{_plural(empty, 'section')} had nothing to order, so no file was made."

        def listed(names):
            return "\n".join(names[:10]) + (f"\n…and {len(names) - 10} more" if len(names) > 10 else "")

        if stale:
            message += (
                "\n\nOlder files for sections with nothing to order are still in the folder:\n"
                + listed([os.path.basename(p) for p in stale])
            )
        if failed:
            message += "\n\nCould not be written:\n" + listed(failed)
        _result_dialog(
            root,
            f"{_plural(len(written), 'file')} generated with "
            f"{_plural(sum(map(len, written)), 'item')} to order",
            message,
            [("Open folder", lambda: os.startfile(folder)), ("View details", show_details)],
            warn=bool(failed),
        )

    run, set_reason = _footer(frm, "Generate", generate, go_back)
    out_var.trace_add("write", lambda *_: update_totals())
    update_totals()

    def start(flight, sections, table, planogram, threshold, stock_path):
        old, new = state["inputs"], (flight, sections, table, planogram, threshold)
        state["sources"] = [p for p in (flight, stock_path) if p]
        if (
            old and old[0] == flight and old[2] is table and old[3:] == new[3:]
            and len(old[1]) == len(sections) and all(a is b for a, b in zip(old[1], sections))
        ):
            refresh()
            return True  # nothing changed: keep the user's ticks and quantities
        if any(i.changed for i in state["items"]) and not messagebox.askyesno(
            APP_NAME,
            "The order will be worked out again, so the changes you made on the Review page "
            "will be lost.\nContinue?",
        ):
            return False
        if flight != state["flight"]:
            out_var.set(os.path.normpath(default_order_folder(flight)))
        state.update(
            flight=flight, sections=sections, inputs=new,
            items=calculate_orders(sections, table, table.columns[planogram], threshold),
        )
        filter_var.set("order" if any(i.include for i in state["items"]) else "all")
        hint()
        refresh()
        return True

    frm.bind("<Unmap>", lambda e: finish_edit() if e.widget is frm else None)
    return Page(frm, run, go_back, None, start)


def enable_file_drop(root, widget, on_drop):
    """Call on_drop(paths) when files are dropped on widget. False if not available."""
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD

        TkinterDnD.require(root)

        def dropped(event):
            # Let the drop finish first: a message box here would freeze Explorer
            root.after(10, on_drop, root.tk.splitlist(event.data))
            return event.action

        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", dropped)
        return True
    except Exception:  # noqa: BLE001 - dropping files is a convenience; run without it
        return False


def run_gui():
    import tkinter as tk
    from tkinter import font, messagebox, ttk

    root = tk.Tk()
    root.withdraw()  # until it is laid out
    root.title(APP_NAME)
    set_window_icon(root)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Ok.TLabel", foreground="#1a7f37")
    style.configure("Warn.TLabel", foreground="#b3261e")
    style.configure("Hint.TLabel", foreground="#5f6368")
    style.configure("Step.TLabel", font=("Segoe UI", 9, "bold"), foreground="#0b57d0")
    style.configure("AppTitle.TLabel", font=("Segoe UI", 16, "bold"))
    style.configure("PageTitle.TLabel", font=("Segoe UI", 12, "bold"))
    style.configure("DialogOk.TLabel", font=("Segoe UI", 12, "bold"), foreground="#1a7f37")
    style.configure("DialogWarn.TLabel", font=("Segoe UI", 12, "bold"), foreground="#b3261e")
    style.configure("Big.TButton", font=("Segoe UI", 11), padding=(16, 10))
    style.configure(
        "Treeview", rowheight=font.nametofont("TkDefaultFont").metrics("linespace") + 6
    )

    container = ttk.Frame(root)
    container.pack(fill="both", expand=True)

    pages = {}  # name -> Page
    current = {"name": None}

    def show(name):
        if current["name"]:
            pages[current["name"]].frame.pack_forget()
        pages[name].frame.pack(fill="both", expand=True)
        current["name"] = name

    def on_drop(paths):
        page = pages[current["name"]]
        if page.drop is None or root.grab_current() is not None:
            return
        if len(paths) != 1:
            messagebox.showerror(APP_NAME, "Drop one file at a time.")
            return
        page.drop(paths[0])

    root.can_drop = enable_file_drop(root, container, on_drop)

    flight = {}  # path and sections chosen on the first B&W page

    def open_stock(path, sections):
        flight.update(path=path, sections=sections)
        pages["bw_stock"].start(path, sections)
        show("bw_stock")

    def open_review(table, planogram, threshold, stock_path):
        if pages["bw_review"].start(
            flight["path"], flight["sections"], table, planogram, threshold, stock_path
        ):
            show("bw_review")

    go_home = lambda: show("home")  # noqa: E731
    pages["home"] = build_home_page(container, show)
    pages["plu"] = build_plu_page(container, root, go_home)
    pages["bw"] = build_bw_page(container, root, go_home, open_stock)
    pages["bw_stock"] = build_bw_stock_page(
        container, root, go_home, lambda: show("bw"), open_review
    )
    pages["bw_review"] = build_bw_review_page(container, root, go_home, lambda: show("bw_stock"))

    def on_return(event):
        kind = event.widget.winfo_class()
        if kind == "TButton":
            event.widget.invoke()
        elif kind != "Text":  # Enter in the paste box is a new line
            action = pages[current["name"]].action
            if action:
                action()

    def on_escape(_):
        back = pages[current["name"]].back
        if back:
            back()

    root.bind("<Return>", on_return)
    root.bind("<Escape>", on_escape)
    show("home")

    # One size for all pages, so the buttons stay put when moving between them
    root.update_idletasks()
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    min_w = min(max(900, pages["plu"].frame.winfo_reqwidth()), screen_w)
    min_h = min(max(p.frame.winfo_reqheight() for p in pages.values()), screen_h - 80)
    w, h = min(max(1000, min_w), screen_w), min(max(640, min_h), screen_h - 80)
    root.minsize(min_w, min_h)
    root.geometry(f"{w}x{h}+{(screen_w - w) // 2}+{max((screen_h - h) // 3, 0)}")
    root.deiconify()

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
            n, _ = extract(p, out)
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

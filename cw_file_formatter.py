"""CW File Formatter

The home page offers two tasks:

Generate PLU file
  Reads an .xlsx or .csv file, or cells pasted from Excel, and writes a
  two-column CSV (PLU, QTY) with no header row. The PLU column can be chosen in
  the window; every other column is ignored and every QTY is the QTY entered in
  the window (default 1). Rows without a PLU are left out. Values can be changed and
  rows deleted in the preview; the input file itself is never changed.

  - No header row: the first column is assumed to be the PLU.
  - Header row (assumed if a cell says 'PLU'): the column headed 'PLU' is assumed,
    and rows up to the header row are left out.

Process B&W Flight
  Reads a B&W flight workbook (.xlsx or .csv) and lists its sections, e.g.
  "HEALTHY BREAK 24A BANANA BOAT", so the user can pick which to process (all
  are selected to begin with).
  A section starts at the text rows above its items: the last one that isn't
  a note ('*WA stores need to ...') is the section title, notes are kept, and a
  short heading before it ('GONDOLA ENDS') is the category for it and the
  sections after it; a
  row of column headings, in any order and starting in any column (e.g.
  'MYCHEM ID | PLU | ...' or 'PAGE NO | PLU | ...'), gives the column headers
  (a section without its own header row reuses the previous one); rows with a
  number in the PLU (or MYCHEM ID) column are the section's items.

  The next page takes a stock report (.xlsx/.csv file or pasted cells) with PLU
  and SOH columns plus a planogram column the user picks. Items are matched on
  PLU and written to a two-column CSV (PLU, QTY), no header row, when
      QTY = MIN QTY 770MM + Planogram - SOH  is at least the threshold.
  A review page lists the items, where they can be left out and their quantities
  changed. Each selected section gets its own CSV, named after the section.

Usage:
  - Double-click the exe and pick a task on the home page (files can be dropped
    onto the window if tkinterdnd2 is installed), or
  - Drag a file onto the exe (its assumed PLU column becomes a PLU file with a QTY
    of 1; output is written next to the input as <name>_PLU.csv).
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


def has_plu_header(rows):
    """Whether any cell is a 'PLU' heading."""
    return any(str(c).strip().lower() == TARGET_HEADER.lower() for r in rows for c in r)


def plu_table(rows, has_header=False):
    """Split rows into ([(column index, label)], body, index of the assumed PLU column).

    body is [(index into rows, row)]. Rows and columns with nothing in them are left out. With a header row (the first row
    holding a 'PLU' cell, else the first row) the assumed PLU column is the one headed
    'PLU', or None if there isn't one. Without, it is the first column.
    """
    # Drop rows that are entirely empty (e.g. formatted-but-blank Excel rows)
    body = [(i, r) for i, r in enumerate(rows) if any(str(c).strip() for c in r)]
    header, default = [], None
    if has_header and body:
        # The header may not be in the first row (e.g. title/notes rows above it)
        start = 0
        for n, (_, row) in enumerate(body):
            names = [str(c).strip().lower() for c in row]
            if TARGET_HEADER.lower() in names:
                start, default = n, names.index(TARGET_HEADER.lower())
                break
        header, body = body[start][1], body[start + 1:]

    columns, used = [], set()
    for i in range(max((len(r) for r in [header, *(r for _, r in body)]), default=0)):
        label = re.sub(r"\s+", " ", _cell(header, i))
        if not label and not any(_cell(r, i) for _, r in body):
            continue
        if not label:
            label = f"Column {i + 1}"
        elif label in used:
            label = f"{label} (column {i + 1})"
        used.add(label)
        columns.append((i, label))
    if not has_header and columns:
        default = columns[0][0]
    return columns, body, default


def extract(path, out_path, qty="1"):
    """Convert an .xlsx/.csv file using the assumed PLU column. See extract_rows."""
    rows = read_rows(path)
    return extract_rows(rows, out_path, has_plu_header(rows), qty, source="file")


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
    plus = [plu for plu in (_cell(r, plu_col) for _, r in body) if plu]
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
    notes: list = field(default_factory=list)  # note rows above the items, e.g. '*WA stores ...'
    category: str = ""  # heading the section is under, e.g. 'GONDOLA ENDS'


# Other spellings of headers, by their normalised form. DESCRIPTION is what the
# review page shows as the item name.
BW_HEADER_ALIASES = {
    "MYCHEMID": BW_ID_HEADER, "MY CHEM ID": BW_ID_HEADER,
    "PRODUCT DESCRIPTION": "DESCRIPTION", "ITEM DESCRIPTION": "DESCRIPTION",
    "PRODUCT NAME": "DESCRIPTION", "ITEM NAME": "DESCRIPTION",
}
# Words that make a cell look like a column heading rather than a title or a value
BW_HEADER_WORDS = re.compile(
    r"\b(PLU|MY ?CHEM ?ID|SKU|APN|BARCODE|DESCRIPTION|PRODUCT|ITEM|NAME|BRAND|SUPPLIER|"
    r"SUPPLIED|VENDOR|RE-?ORDER|ORDERS?|QTY|QUANTITY|MIN|MAX|PAGE|NO|CODE|PRICE|RRP|"
    r"COST|SAVING|COMMENTS?|NOTES?|SOH|STOCK|PLANOGRAM|CATEGORY|SIZE)\b"
)


def _norm_header(text):
    """'PLU:' -> 'PLU', ' mychem id ' -> 'MYCHEM ID', 'Product Description' -> 'DESCRIPTION'."""
    name = re.sub(r"\s+", " ", str(text)).strip().rstrip(":").strip().upper()
    return BW_HEADER_ALIASES.get(name, name)


def _looks_like_header(name):
    """Whether a normalised cell reads like a column heading: short, no figures."""
    return (
        len(name) <= 40 and not re.search(r"\d{3,}|\$\s*\d", name)
        and BW_HEADER_WORDS.search(name) is not None
    )


def _bw_headers(cells, next_cells=None):
    """[(column index, header name)] if the row is a header row, else None.

    The headings can be in any order and start in any column. A row counts if
      - two or more of its cells read like headings (PLU, SUPPLIER, MIN QTY ...), or
      - it has a PLU or MYCHEM ID heading and another cell, or
      - it has three or more text cells and the next row (next_cells) has a
        number in one of those columns, so it reads like headings over items.
    """
    headers = [(i, _norm_header(c)) for i, c in enumerate(cells) if c]
    if len(headers) < 2 or any(_is_number(c) for _, c in headers):
        return None
    names = {name for _, name in headers}
    if names & {TARGET_HEADER, BW_ID_HEADER}:
        return headers
    if sum(_looks_like_header(name) for _, name in headers) >= 2:
        return headers
    if len(headers) >= 3 and next_cells and any(
        _is_number(_cell(next_cells, i)) for i, _ in headers
    ):
        return headers
    return None


def _bw_key_columns(headers):
    """Columns that hold a number on every item row: PLU and MYCHEM ID."""
    return [i for i, name in headers or () if name in (TARGET_HEADER, BW_ID_HEADER)]


def _is_number(text):
    return re.fullmatch(r"\d+(\.0+)?", text) is not None


# Text rows that are notes to the reader, not section titles
BW_NOTE_START = re.compile(r"^\s*(\*|N\.?B\b|NOTE\b|PLEASE\b|IMPORTANT\b)", re.IGNORECASE)
BW_NOTE_WORDS = re.compile(r"\b(NEEDS? TO|MUST|SHOULD|PLEASE)\b", re.IGNORECASE)


def _is_bw_note(text):
    """Whether a text row reads like a note ('*WA and SA stores need to order ...')
    rather than a section title ('#1 P&G: ORAL B & VOOST END')."""
    text = text.strip()
    return bool(
        BW_NOTE_START.search(text) or BW_NOTE_WORDS.search(text) or text.endswith((".", "!"))
    )


# Category headings in B&W flight files, in the order they are listed. Other headings
# that read like categories are grouped under BW_OTHER_CATEGORIES.
BW_CATEGORIES = [
    "GONDOLA ENDS", "MESH ENDS", "TILL MESH", "SCRIPTS MESH", "FLOOR STANDS", "MEZZANINE",
    "FRAGRANCE PODIUM", "WOITW", "P&G SCRIPTS IN PODS 1-12", "OPELLA SCRIPTS OUT PODS 1-12",
]
BW_OTHER_CATEGORIES = "Other categories"
# Which sections are ticked when a file is opened, by category:
BW_SELECT_ALL = {"GONDOLA ENDS", "MESH ENDS", "TILL MESH", "SCRIPTS MESH", "FRAGRANCE PODIUM"}
BW_SKIP_NUMBERS = {"TILL MESH": {5}}  # ...except these section numbers ('#5 ...')
BW_SELECT_GRADES = {"WOITW": {"D"}}  # ...and these grades ('#48D ...' is grade D)
BW_SKIP_TITLES = ("SELECTED STORES",)  # sections whose title starts so are never ticked


def _category_key(text):
    """'Gondola  Ends:' -> 'GONDOLA ENDS', for comparing headings with BW_CATEGORIES."""
    return re.sub(r"\s+", " ", re.sub(r"[:.\s]+$", "", str(text))).strip().upper()


_BW_CATEGORY_KEYS = {_category_key(c): c for c in BW_CATEGORIES}


def _known_category(text):
    """The BW_CATEGORIES name a heading is, or None."""
    return _BW_CATEGORY_KEYS.get(_category_key(text))


def bw_group(section):
    """The category a section is listed under: one of BW_CATEGORIES or BW_OTHER_CATEGORIES."""
    return _known_category(section.category) or BW_OTHER_CATEGORIES


def _section_number_grade(title):
    """('48', 'D') for '#48D CWH: WAGNER ...', (None, None) if it has no '#' number."""
    found = re.match(r"\s*#\s*(\d+)([A-Z])?(?![A-Z])", title, re.IGNORECASE)
    if not found:
        return None, None
    return found.group(1), (found.group(2) or "").upper() or None


def bw_default_selected(section):
    """Whether a section is ticked when the file is opened (see BW_SELECT_ALL)."""
    if section.title.strip().upper().startswith(BW_SKIP_TITLES):
        return False
    group = bw_group(section)
    number, grade = _section_number_grade(section.title)
    if group in BW_SELECT_ALL:
        return not (number and int(number) in BW_SKIP_NUMBERS.get(group, ()))
    if group in BW_SELECT_GRADES:
        return grade in BW_SELECT_GRADES[group]
    return False


def _is_bw_category(text):
    """Whether a text row reads like a category heading ('GONDOLA ENDS', 'MESH ENDS'):
    a few words with no figures, unlike section titles ('#1 P&G: ORAL B END ...')."""
    return (
        not _is_bw_note(text) and not re.search(r"[\d#]", text)
        and len(text.split()) <= 5 and len(text) <= 40
    )


def _pick_bw_title(texts):
    """(title, notes, category) from the text rows above a section's items, given as
    (text, above a header row).

    The title is the last row that isn't a note; if every row reads like a note, the
    last row is used. The category is the last category-like row before the title
    ('' if none). A heading from BW_CATEGORIES wins over other category-like rows and
    counts even above a header row ('WOITW'); other rows above a header row are
    banners ('COMPLIANCE ENDS DECEMBER CATALOGUE'), not categories.
    """
    plain = [t for t, _ in texts if not _is_bw_note(t)]
    title = (plain or [t for t, _ in texts])[-1]
    at = next(i for i in reversed(range(len(texts))) if texts[i][0] is title)
    notes = [t for t, _ in texts if t is not title and _is_bw_note(t)]
    if _known_category(title):
        return title, notes, _known_category(title)  # a category with no section title
    # Headings from BW_CATEGORIES count wherever they are; other category-like rows
    # only if they are not banners
    known = [_known_category(t) for t, _ in texts[:at] if _known_category(t)]
    if known:
        return title, notes, known[-1]
    others = [t for t, banner in texts[:at] if not banner and _is_bw_category(t)]
    return title, notes, others[-1] if others else ""


def parse_bw_sections(rows):
    """Split B&W flight rows into sections. Sections with no items are left out.

    Item rows are told apart from text rows by a number in the PLU or MYCHEM ID
    column, or, before any header row is known (or if it has neither), by a number in
    the first filled cell. The text rows just above a section's items give its title,
    notes and category (see _pick_bw_title); a category carries on to the sections
    after it until the next one.
    """
    sections = []
    current = None
    category = ""
    headers = None  # [(column index, header name)] from the latest header row
    texts = []  # (text, above a header row) for the text rows since the last item
    rows = [[str(c).strip() for c in row] for row in rows]
    rows = [r for r in rows if any(r)]  # leave out blank (or colour-only) rows

    for n, cells in enumerate(rows):
        first = next(i for i, c in enumerate(cells) if c)
        lead = cells[first]
        keys = _bw_key_columns(headers)
        if keys:
            is_item = any(_is_number(_cell(cells, i)) for i in keys)
        else:
            is_item = _is_number(lead)

        if not is_item:
            row_headers = _bw_headers(cells, rows[n + 1] if n + 1 < len(rows) else None)
            if row_headers:
                headers = row_headers
                texts = [(t, True) for t, _ in texts]
            else:
                texts.append((re.sub(r"\s+", " ", lead).strip(), False))
            continue

        if texts:
            title, notes, new_category = _pick_bw_title(texts)
            category = new_category or category
            current = BWSection(title, notes=notes, category=category)
            sections.append(current)
            texts = []
        if current is None:
            continue  # items before any section title: nothing to file them under
        cols = headers or list(enumerate(BW_DEFAULT_HEADERS, start=first))
        current.items.append({name: _cell(cells, i) for i, name in cols})

    return sections


def load_bw_sections(path):
    sections = parse_bw_sections(read_rows(path))
    if not sections:
        raise ExtractError(
            "No sections found in this file.\n\nExpected section title rows (e.g. "
            "'HEALTHY BREAK 24A BANANA BOAT') followed by rows starting with a MYCHEM ID "
            "or PLU."
        )
    return sections


def section_suppliers(section):
    """The different suppliers of the section's items, in the order they first appear."""
    names = {}
    for item in section.items:
        name = re.sub(r"\s+", " ", item.get("SUPPLIER", "")).strip()
        if name:
            names.setdefault(name.upper(), name)
    return list(names.values())


def section_supplier(section):
    """The section's suppliers as one line, e.g. 'DIRECT, SIGMA'."""
    return ", ".join(section_suppliers(section))


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
    body: list = field(default_factory=list)  # [(index into the rows parsed, row)], every item row


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
    for i, row in enumerate(rows[ri + 1:], start=ri + 1):
        key = _plu_key(_cell(row, plu_col))
        if not key or key == TARGET_HEADER:
            continue  # footer, or the header row repeated
        table.body.append((i, row))
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
        # System aware (Win 8.1+). Tk can't rescale a window moved to a monitor with another
        # scaling, so leave that to Windows rather than claim to be per-monitor aware.
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
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
# States of the tick boxes in the lists. BLOCKED: there is nothing to tick.
UNCHECKED, CHECKED, PARTLY, BLOCKED = "☐", "☑", "▣", ""
CHECK_WIDTH = 44  # of the column they are in

# Display scaling: 1.0 at 100%, 1.5 at 150%. Set when the window is made.
_scale = 1.0

LIGHT = {
    "field": "#ffffff", "text": "#1c1c1c", "hint": "#5f6368", "ok": "#1a7f37", "warn": "#b3261e",
    "step": "#005fb8", "on": "#e6f0fb", "off": "#80868b", "section": "#e9ebee", "border": "#d9d9d9", "box": "#868686", "mark": "#ffffff",
    "selected": "#b9d7f5",
}
DARK = {
    "field": "#202020", "text": "#fafafa", "hint": "#a8adb3", "ok": "#6ccb5f", "warn": "#ff99a4",
    "step": "#57c8ff", "on": "#1d3146", "off": "#8b9096", "section": "#343434", "border": "#3a3a3a", "box": "#9a9a9a", "mark": "#000000",
    "selected": "#2b5a88",
}


def _px(n):
    """A size in pixels at 100% display scaling, in pixels at the current scaling."""
    return round(n * _scale)


def _check_png(state, size, colors, pad=0):
    """PNG of a tick box, drawn here so that it is sharp at any size and fits the colours.

    pad is the width of the empty space on its left and right.
    """
    import struct
    import zlib

    def rgb(name):
        return tuple(int(colors[name][i:i + 2], 16) for i in (1, 3, 5))

    filled = state in (CHECKED, PARTLY)
    edge = rgb("step") if filled else rgb("border" if state == BLOCKED else "box")
    inside, mark = rgb("step") if filled else rgb("field"), rgb("mark")
    line = 1.25 * _scale / size  # width of the border

    def in_box(x, y, inset):
        """Whether a point is in the rounded square, shrunk by inset."""
        half, radius = 0.5 - inset, 0.2 - inset
        dx, dy = abs(x - 0.5) - (half - radius), abs(y - 0.5) - (half - radius)
        if dx > 0 and dy > 0:
            return math.hypot(dx, dy) <= radius
        return max(dx, dy) <= radius

    def near(x, y, x1, y1, x2, y2):
        """Whether a point is on a line drawn with a round pen."""
        t = ((x - x1) * (x2 - x1) + (y - y1) * (y2 - y1)) / ((x2 - x1) ** 2 + (y2 - y1) ** 2)
        t = min(max(t, 0), 1)
        return math.hypot(x - x1 - t * (x2 - x1), y - y1 - t * (y2 - y1)) <= 0.06

    def color(x, y):
        if not in_box(x, y, 0.02):
            return None
        if state == CHECKED and (
            near(x, y, 0.27, 0.52, 0.43, 0.68) or near(x, y, 0.43, 0.68, 0.74, 0.35)
        ):
            return mark
        if state == PARTLY and near(x, y, 0.28, 0.5, 0.72, 0.5):
            return mark
        return inside if in_box(x, y, 0.02 + line) else edge

    steps = [(i + 0.5) / 4 for i in range(4)]  # 4 x 4 samples in each pixel, for smooth edges
    data = bytearray()
    for py in range(size):
        data.append(0)  # PNG filter type: none
        data += bytes(4 * pad)
        for px in range(size):
            hits = [color((px + a) / size, (py + b) / size) for a in steps for b in steps]
            hits = [c for c in hits if c]
            n = len(hits) or 1
            data += bytes(sum(c[i] for c in hits) // n for i in range(3))
            data.append(255 * len(hits) // 16)
        data += bytes(4 * pad)

    def chunk(kind, body):
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))

    header = struct.pack(">IIBBBBB", size + 2 * pad, size, 8, 6, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(
        b"IDAT", zlib.compress(bytes(data))
    ) + chunk(b"IEND", b"")


def _check_column(tree, root, on_heading):
    """Set up a list's first column (#0) for tick boxes; clicking its heading calls on_heading."""
    tree.configure(show="tree headings", style="Check.Treeview")
    tree.column("#0", width=_px(CHECK_WIDTH), minwidth=_px(CHECK_WIDTH), stretch=False)
    tree.heading("#0", image=root.heading_checks[BLOCKED], anchor="center", command=on_heading)

    def motion(event):
        over = tree.identify_column(event.x) == "#0" and (
            tree.identify_region(event.x, event.y) == "heading" or tree.identify_row(event.y)
        )
        tree.configure(cursor="hand2" if over else "")

    tree.bind("<Motion>", motion, add="+")


def apply_theme(root, dark=False):
    """Colours and fonts. root.themed says whether the modern theme is available.

    Functions in root.on_theme are called with the colours, for what styles don't reach.
    """
    from tkinter import font, ttk

    style = ttk.Style(root)
    try:
        import sv_ttk

        sv_ttk.set_theme("dark" if dark else "light", root)
        root.themed = True
    except Exception:  # noqa: BLE001 - the look is a nicety; fall back to the Windows theme
        root.themed = dark = False
        if "vista" in style.theme_names():
            style.theme_use("vista")
    colors = DARK if dark else LIGHT

    # sv_ttk's fonts are sized in pixels, so they ignore display scaling, and come with
    # Windows 11 only. Sizes in points follow the scaling.
    sizes = {
        "TkDefaultFont": 10, "TkTextFont": 10, "TkHeadingFont": 9,
        "SunValleyCaptionFont": 9, "SunValleyBodyFont": 10, "SunValleyBodyStrongFont": 10,
    }
    existing = set(font.names(root))
    for name, size in sizes.items():
        if name in existing:
            font.nametofont(name, root).configure(family="Segoe UI", size=size)
    root.option_add("*TCombobox*Listbox.font", "TkTextFont")

    for name in ("Ok", "Warn", "Hint"):
        style.configure(f"{name}.TLabel", foreground=colors[name.lower()])
    style.configure("Step.TLabel", font=("Segoe UI", 10, "bold"), foreground=colors["step"])
    style.configure("AppTitle.TLabel", font=("Segoe UI", 18, "bold"))
    style.configure("PageTitle.TLabel", font=("Segoe UI", 13, "bold"))
    style.configure("DialogOk.TLabel", font=("Segoe UI", 13, "bold"), foreground=colors["ok"])
    style.configure("DialogWarn.TLabel", font=("Segoe UI", 13, "bold"), foreground=colors["warn"])
    style.configure("Big.TButton", font=("Segoe UI", 12), padding=(_px(16), _px(10)))
    row_height = font.nametofont("TkDefaultFont", root).metrics("linespace") + _px(8)
    style.configure("Treeview", rowheight=row_height)
    if root.themed:
        if "White.field" not in style.element_names():
            # Plain box: the theme's own is drawn in the window's colour
            style.element_create("White.field", "from", "clam", "Treeview.field")
        style.layout("Treeview", [("White.field", {"sticky": "nswe", "border": 1, "children": [
            ("Treeview.padding", {"sticky": "nswe", "children": [
                ("Treeview.treearea", {"sticky": "nswe"})]})]})])
        style.configure(
            "Treeview", background=colors["field"], fieldbackground=colors["field"],
            bordercolor=colors["border"], lightcolor=colors["field"],
        )
        style.configure("Heading", padding=(_px(4), _px(3)))  # in line with the cells' text
        style.map(
            "Treeview", background=[("selected", colors["selected"])],
            foreground=[("selected", colors["text"])],
        )
    # Tick boxes: the same images are redrawn, so the lists using them follow
    if not hasattr(root, "checks"):
        import tkinter

        states = (UNCHECKED, CHECKED, PARTLY, BLOCKED)
        root.checks = {state: tkinter.PhotoImage(master=root) for state in states}
        root.heading_checks = {state: tkinter.PhotoImage(master=root) for state in states}
    size = _px(17)
    # Headings put an image on the right, so fill the heading to line up with the rows
    pad = (_px(CHECK_WIDTH) - 2 * _px(4) - size) // 2
    for state in root.checks:
        root.checks[state].configure(data=_check_png(state, size, colors))
        root.heading_checks[state].configure(data=_check_png(state, size, colors, pad))
    style.layout("Check.Treeview.Item", [("Treeitem.padding", {"sticky": "nswe", "children": [
        ("Treeitem.image", {"sticky": ""})]})])

    # The tree's own cells are too low for the theme's entry box
    style.layout("Cell.TEntry", [("Entry.padding", {"sticky": "nswe", "children": [
        ("Entry.textarea", {"sticky": "nswe"})]})])
    style.configure("Cell.TEntry", padding=(_px(3), 0))

    for update in root.on_theme:
        update(colors)
    _dark_title_bar(root, dark)
    return colors


def _dark_title_bar(window, dark):
    if sys.platform != "win32":
        return
    import ctypes
    from ctypes import wintypes

    try:
        window.update_idletasks()
        user32 = ctypes.windll.user32
        hwnd = wintypes.HWND(user32.GetParent(wintypes.HWND(window.winfo_id())))
        value = ctypes.c_int(int(dark))
        # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), 4)
        # Windows 10 only repaints the title bar when it changes between active and
        # inactive, so draw it as the other one and back (0x86 = WM_NCACTIVATE)
        active = int(user32.GetForegroundWindow() == hwnd.value)
        for state in (1 - active, active):
            user32.SendMessageW(hwnd, 0x86, wintypes.WPARAM(state), wintypes.LPARAM(0))
    except Exception:  # noqa: BLE001 - cosmetic only
        pass


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


def _natural_key(text):
    """Sort key where numbers sort by value, so 'BREAK 9' comes before 'BREAK 24A'."""
    return [
        (0, int(part), "") if part.isdecimal() else (1, 0, part)
        for part in re.split(r"(\d+)", str(text).upper()) if part
    ]


def _page(parent, title, go_home, step=None):
    """A task page: header, body that takes the spare space. Returns (frame, body)."""
    from tkinter import ttk

    frm = ttk.Frame(parent, padding=_px(16))
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
    button = ttk.Button(
        bar, text=text, command=command, default="active", width=14, style="Accent.TButton"
    )
    button.grid(row=0, column=2)

    def set_reason(reason):
        reason_lbl.configure(text=reason)
        button.state(["disabled" if reason else "!disabled"])

    def run():
        if not button.instate(["disabled"]):
            command()

    return run, set_reason


def _scrollbar(widget, orient, row, column, on_scroll=None):
    """Scrollbar beside a widget in its parent's grid, hidden while there is nothing to scroll."""
    from tkinter import ttk

    vertical = orient == "vertical"
    bar = ttk.Scrollbar(
        widget.master, orient=orient, command=widget.yview if vertical else widget.xview
    )
    bar.grid(row=row, column=column, sticky="ns" if vertical else "ew")

    def scrolled(first, last):
        (bar.grid_remove if float(first) <= 0 and float(last) >= 1 else bar.grid)()
        bar.set(first, last)
        if on_scroll:
            on_scroll()

    widget.configure(**{"yscrollcommand" if vertical else "xscrollcommand": scrolled})


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
    frm = ttk.Frame(win, padding=_px(20))
    frm.pack(fill="both", expand=True)
    ttk.Label(frm, text=title, style="DialogWarn.TLabel" if warn else "DialogOk.TLabel").pack(anchor="w")
    ttk.Label(frm, text=message, wraplength=_px(520), justify="left").pack(anchor="w", pady=(10, 16))
    bar = ttk.Frame(frm)
    bar.pack(anchor="e")

    def press(command=None):
        win.grab_release()
        win.destroy()
        if command:
            command()

    for label, command in buttons:
        ttk.Button(bar, text=label, command=lambda c=command: press(c)).pack(side="left", padx=(0, 8))
    close = ttk.Button(bar, text="Close", command=press, default="active", style="Accent.TButton")
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
    _dark_title_bar(win, root.dark)
    win.grab_set()
    close.focus_set()
    return win


def _section_preview(root, section):
    """Window listing a B&W section's items, read only. Several can be open at once."""
    import tkinter as tk
    from tkinter import font, ttk

    # Columns in the order they first appear, leaving out ones with nothing in them
    names = []
    for item in section.items:
        names += [n for n in item if n not in names]
    names = [n for n in names if any(str(i.get(n, "")).strip() for i in section.items)]

    win = tk.Toplevel(root)
    win.withdraw()
    win.title(f"{section.title} - {APP_NAME}")
    win.transient(root)
    frm = ttk.Frame(win, padding=_px(16))
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(3, weight=1)
    ttk.Label(frm, text=section.title, style="PageTitle.TLabel").grid(row=0, column=0, sticky="w")
    supplier = section_supplier(section)
    ttk.Label(
        frm, style="Hint.TLabel",
        text=(f"{section.category}  ·  " if section.category else "")
        + _plural(len(section.items), "item") + (f"  ·  Supplier: {supplier}" if supplier else ""),
    ).grid(row=1, column=0, sticky="w", pady=(2, 10))
    if section.notes:
        ttk.Label(
            frm, text="\n".join(section.notes), style="Hint.TLabel", wraplength=_px(700),
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(0, 10))

    box = ttk.Frame(frm)
    box.grid(row=3, column=0, sticky="nsew")
    box.columnconfigure(0, weight=1)
    box.rowconfigure(0, weight=1)
    cols = [f"c{i}" for i in range(len(names))]
    tree = ttk.Treeview(box, columns=cols, show="headings", height=14, selectmode="extended")
    tree.grid(row=0, column=0, sticky="nsew")
    _scrollbar(tree, "vertical", 0, 1)
    _scrollbar(tree, "horizontal", 1, 0)
    rows = [[_one_line(item.get(n, "")) for n in names] for item in section.items]
    measure = font.nametofont("TkDefaultFont").measure
    for i, name in enumerate(names):
        widest = max([r[i] for r in rows], key=len, default="")
        width = max(measure(name), measure(widest)) + _px(24)
        numeric = all(_is_number(r[i]) for r in rows if r[i])
        tree.column(
            cols[i], width=min(max(width, _px(50)), _px(360)), minwidth=_px(30),
            stretch=False, anchor="e" if numeric else "w",
        )
        tree.heading(cols[i], text=name, anchor="w")
    for r in rows:
        tree.insert("", "end", values=r)

    close = ttk.Button(frm, text="Close", command=win.destroy, style="Accent.TButton")
    close.grid(row=4, column=0, sticky="e", pady=(12, 0))
    win.bind("<Escape>", lambda e: win.destroy())
    win.bind("<Return>", lambda e: win.destroy())

    win.update_idletasks()
    w = min(max(win.winfo_reqwidth(), _px(480)), round(root.winfo_screenwidth() * 0.9))
    h = min(win.winfo_reqheight(), round(root.winfo_screenheight() * 0.85))
    x = root.winfo_rootx() + (root.winfo_width() - w) // 2
    y = root.winfo_rooty() + (root.winfo_height() - h) // 3
    win.geometry(f"{w}x{h}+{max(x, 0)}+{max(y, 0)}")
    win.minsize(_px(360), _px(240))
    win.deiconify()
    _dark_title_bar(win, root.dark)
    tree.focus_set()
    return win


class CellEditor:
    """Entry placed over a Treeview cell to change its value.

    on_done(iid, column, text) is called when it closes, with None as the text if the
    user cancelled.
    """

    def __init__(self, tree, on_done, justify="left"):
        self.tree, self.on_done, self.justify = tree, on_done, justify
        self.entry = self.cell = None
        for event in ("<Configure>", "<Unmap>"):
            tree.bind(event, lambda e: self.close(), add="+")

    def open(self, iid, column, text):
        from tkinter import ttk

        tree = self.tree
        self.close()
        if not tree.exists(iid):
            return
        tree.see(iid)
        tree.update_idletasks()
        bbox = tree.bbox(iid, column)
        if not bbox:
            return  # scrolled out of sight sideways
        x, y, w, h = bbox
        left = max(x, 0)
        entry = ttk.Entry(tree, justify=self.justify, style="Cell.TEntry")
        entry.insert(0, text)
        entry.select_range(0, "end")
        entry.place(x=left, y=y, width=min(x + w, tree.winfo_width()) - left, height=h)
        self.entry, self.cell = entry, (iid, column)

        def done(apply):
            self.close(apply)
            tree.focus_set()
            return "break"

        def left_entry():
            # Only if the focus really moved to something else in this window
            if self.entry is entry and entry.focus_get() not in (None, entry):
                self.close()

        entry.bind("<Return>", lambda e: done(True))
        entry.bind("<Tab>", lambda e: done(True))
        entry.bind("<Escape>", lambda e: done(False))
        entry.bind("<FocusOut>", lambda e: tree.after(50, left_entry))
        entry.focus_set()

    def close(self, apply=True):
        """Close the entry if it is open. Returns what on_done returned, else True."""
        if self.entry is None:
            return True
        entry, cell = self.entry, self.cell
        self.entry = None  # first: destroying the entry fires FocusOut
        text = entry.get().strip()
        entry.destroy()
        return self.on_done(*cell, text if apply else None)


class RowsTable:
    """Table of rows of text, whose values can be changed and rows deleted.

    .rows holds all the rows read; show() displays some of their rows and columns.
    on_change() is called after the user changes .rows, and should call show().
    on_pick(column number) is called when a column heading is clicked.
    Place .frame, and the frame made by buttons().
    """

    MAX_ROWS = 5000

    def __init__(self, parent, height, on_change, on_pick=None):
        from tkinter import font, ttk

        self.frame = ttk.Frame(parent)
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(self.frame, show="headings", height=height, selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")
        # An open entry doesn't scroll with the table, so close it
        _scrollbar(self.tree, "vertical", 0, 1, lambda: self.finish())
        _scrollbar(self.tree, "horizontal", 1, 0, lambda: self.finish())
        self.font = font.nametofont("TkDefaultFont")
        self.rows, self.headings, self.col_ids, self.all_ids = [], [], [], []
        self.shown = {}  # {row index: values displayed}
        self.undo_steps = []  # ("edit", row, column, old text) or ("delete", [(row, cells)])
        self.on_change, self.on_pick = on_change, on_pick
        self.delete_btn = self.undo_btn = None
        self.editor = CellEditor(self.tree, self._edited)
        self.tree.bind("<Button-1>", self._clicked, add="+")
        self.tree.bind("<Double-1>", self._double_clicked)
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Control-z>", lambda e: self.undo())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_buttons())

    def buttons(self, parent):
        from tkinter import ttk

        frm = ttk.Frame(parent)
        self.delete_btn = ttk.Button(frm, text="Delete rows", command=self.delete_selected)
        self.delete_btn.pack(side="left")
        self.undo_btn = ttk.Button(frm, text="Undo", command=self.undo)
        self.undo_btn.pack(side="left", padx=(6, 0))
        self._update_buttons()
        return frm

    def _update_buttons(self):
        if self.delete_btn:
            self.delete_btn.state(["!disabled" if self._selected() else "disabled"])
            self.undo_btn.state(["!disabled" if self.undo_steps else "disabled"])

    def _selected(self):
        return [int(i) for i in self.tree.selection() if i.isdigit()]

    # --- display
    def load(self, rows):
        """Take new rows, forgetting earlier changes. Call show() afterwards."""
        self.editor.close(False)
        self.rows, self.undo_steps = [list(r) for r in rows], []

    def show(self, headings, row_ids, col_ids, marked=None, keep_place=False):
        """Show the given rows and columns (indexes into .rows) under the headings."""
        self.editor.close(False)
        tree = self.tree
        place = tree.yview()[0], tree.xview()[0], tree.selection()
        before, self.headings = (self.headings, self.col_ids, len(self.all_ids)), [
            _one_line(h) for h in headings
        ]
        headings, self.col_ids, self.all_ids = self.headings, list(col_ids), list(row_ids)
        old, shown = self.shown, {
            r: [_one_line(_cell(self.rows[r], c)) for c in self.col_ids]
            for r in row_ids[: self.MAX_ROWS]
        }
        self.shown = shown
        if keep_place and before == (headings, self.col_ids, len(row_ids)) and list(old) == list(shown):
            for r, values in shown.items():
                if values != old[r]:
                    tree.item(str(r), values=values)
            self.mark(marked)
            self._update_buttons()
            return
        tree.delete(*tree.get_children())
        tree.configure(columns=[f"c{i}" for i in range(len(headings))])
        first = list(shown.values())[:100]
        for i, text in enumerate(headings):
            widest = max([r[i] for r in first], key=len, default="")
            heading = f"✔ PLU: {text}" if self.on_pick else text
            width = max(self.font.measure(heading), self.font.measure(widest)) + _px(24)
            tree.column(
                f"c{i}", width=min(max(width, _px(50)), _px(300)), minwidth=_px(30),
                stretch=False, anchor="w",
            )
        self.mark(marked)
        for r, values in shown.items():
            tree.insert("", "end", iid=str(r), values=values)
        if headings and len(row_ids) > len(shown):
            tree.insert("", "end", values=[f"… and {len(row_ids) - len(shown):,} more rows"])
        if keep_place:
            tree.update_idletasks()
            tree.yview_moveto(place[0])
            tree.xview_moveto(place[1])
            tree.selection_set([i for i in place[2] if tree.exists(i)])
        self._update_buttons()

    def mark(self, column):
        """Show which column is the PLU column."""
        for i, text in enumerate(self.headings):
            if i == column:
                text = "✔ PLU" if text.upper() == TARGET_HEADER else f"✔ PLU: {text}"
            self.tree.heading(f"c{i}", text=text, anchor="w")

    # --- clicks
    def _clicked(self, event):
        self.finish()
        if self.on_pick and self.tree.identify_region(event.x, event.y) == "heading":
            column = int(self.tree.identify_column(event.x)[1:]) - 1
            if 0 <= column < len(self.headings):
                self.on_pick(column)

    def _double_clicked(self, event):
        tree = self.tree
        row = tree.identify_row(event.y)
        if tree.identify_region(event.x, event.y) != "cell" or not row.isdigit():
            return None
        column = int(tree.identify_column(event.x)[1:]) - 1
        text = _one_line(_cell(self.rows[int(row)], self.col_ids[column]))
        self.editor.open(row, f"c{column}", text)
        return "break"

    # --- changes
    def finish(self):
        """Apply the value being changed, if any."""
        self.editor.close()

    def _edited(self, iid, column, text):
        row, col = int(iid), self.col_ids[int(column[1:])]
        old = str(_cell(self.rows[row], col))
        if text is None or text == _one_line(old):
            return True
        self._put(row, col, text)
        self._changed(("edit", row, col, old))
        return True

    def _put(self, row, col, text):
        cells = self.rows[row]
        cells.extend([""] * (col + 1 - len(cells)))
        cells[col] = text

    def delete_selected(self):
        self.finish()
        ids = sorted(self._selected())
        if not ids:
            return
        removed = [(i, self.rows[i]) for i in ids]
        for i in reversed(ids):
            del self.rows[i]
        self.tree.selection_set(())
        self._changed(("delete", removed))

    def undo(self):
        self.editor.close(False)
        if not self.undo_steps:
            return
        kind, *step = self.undo_steps.pop()
        if kind == "edit":
            self._put(*step)
        else:
            for i, cells in step[0]:
                self.rows.insert(i, cells)
        self.tree.selection_set(())
        self._changed()

    def _changed(self, step=None):
        if step:
            self.undo_steps.append(step)
        self.on_change()
        self._update_buttons()


def _raw_table(rows):
    """What to show for rows without a recognised header row, as RowsTable.show arguments."""
    ids = [i for i, r in enumerate(rows) if any(str(c).strip() for c in r)]
    cols = range(max((len(rows[i]) for i in ids), default=0))
    return [f"Column {c + 1}" for c in cols], ids, cols


def _paste_box(parent, root, on_change, view, height=6, on_pick=None):
    """Box for cells pasted from Excel: a text box that turns into a table after pasting.

    on_change(edited) is called when the rows (table.rows) change; edited is True if a
    value was changed or rows deleted in the table. view() gives the RowsTable.show
    arguments. Returns (frame, table, focus).
    """
    import io
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
    _scrollbar(txt, "vertical", 0, 1)
    root.on_theme.append(lambda colors: txt.configure(
        background=colors["field"], foreground=colors["text"], insertbackground=colors["text"]
    ))
    table = RowsTable(box, height, lambda: table_changed(), on_pick)
    for w in (text_frame, table.frame):
        w.grid(row=1, column=0, sticky="nsew")
    table.frame.grid_remove()
    # table: whether the table is showing; text: the text that table.rows matches
    state = {"table": False, "text": ""}

    foot = ttk.Frame(box)
    foot.grid(row=2, column=0, sticky="ew", pady=(4, 0))
    foot.columnconfigure(0, weight=1)
    ttk.Label(foot, textvariable=info, style="Hint.TLabel").grid(row=0, column=0, sticky="w")

    def get_text():
        return txt.get("1.0", "end-1c")

    def update_info():
        n = sum(1 for r in table.rows if any(c.strip() for c in r))
        text = f"{_plural(n, 'row')} pasted" if n else ""
        if n and state["table"]:
            text += ". Double-click a value to change it."
        info.set(text)
        toggle_btn.state(["!disabled" if get_text().strip() else "disabled"])

    def focus():
        (table.tree if state["table"] else txt).focus_set()

    def read_text():
        """Take the rows from the text, if it is different from the rows held."""
        text = get_text()
        if text == state["text"]:
            return
        state["text"] = text
        try:
            table.load(parse_pasted(text))
        except csv.Error:
            table.load([])
        on_change(False)
        if state["table"]:
            table.show(*view())

    def table_changed():
        """A value was changed or rows were deleted: rewrite the text to match."""
        out = io.StringIO()
        csv.writer(out, delimiter="\t", lineterminator="\n").writerows(table.rows)
        state["text"] = out.getvalue()
        txt.delete("1.0", "end")
        txt.insert("1.0", state["text"])
        on_change(True)
        table.show(*view(), keep_place=True)
        update_info()

    def show_table(on):
        table.finish()
        read_text()
        on = on and bool(get_text().strip())
        state["table"] = on
        (text_frame if on else table.frame).grid_remove()
        (table.frame if on else text_frame).grid()
        (table_btns.grid if on else table_btns.grid_remove)()
        toggle_btn.configure(text="Edit as text" if on else "Show as table")
        if on:
            table.show(*view())
        update_info()
        focus()

    def paste_clipboard(_=None):
        try:
            data = root.clipboard_get()
        except tk.TclError:
            messagebox.showerror(APP_NAME, "The clipboard is empty.\nCopy the cells in Excel first.")
            return
        table.editor.close(False)
        txt.delete("1.0", "end")
        txt.insert("1.0", data)
        show_table(True)

    def clear():
        table.editor.close(False)
        txt.delete("1.0", "end")
        show_table(False)

    table_btns = table.buttons(foot)
    table_btns.grid(row=0, column=1, padx=(0, 12))
    table_btns.grid_remove()
    toggle_btn = ttk.Button(
        foot, text="Show as table", command=lambda: show_table(not state["table"])
    )
    toggle_btn.grid(row=0, column=2)
    toggle_btn.state(["disabled"])
    ttk.Button(foot, text="Paste", command=paste_clipboard).grid(row=0, column=3, padx=(6, 0))
    ttk.Button(foot, text="Clear", command=clear).grid(row=0, column=4, padx=(6, 0))

    def modified(_=None):
        if not txt.edit_modified():
            return  # fired by the reset below
        txt.edit_modified(False)
        read_text()
        update_info()

    txt.bind("<<Modified>>", modified)
    txt.bind("<<Paste>>", lambda e: txt.after_idle(show_table, True))
    table.tree.bind("<<Paste>>", paste_clipboard)
    return box, table, focus


def build_home_page(parent, root, open_page):
    import tkinter as tk
    from tkinter import ttk

    page_frm = ttk.Frame(parent, padding=_px(24))
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
    if root.themed:
        dark = tk.BooleanVar(value=root.dark)

        def switch():
            root.dark = dark.get()
            apply_theme(root, root.dark)

        ttk.Checkbutton(
            page_frm, text="Dark mode", variable=dark, command=switch, style="Switch.TCheckbutton"
        ).place(relx=1, rely=1, anchor="se")
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
    # For each input: the columns [(index, label)] and rows below the header [(row index,
    # row)] found, and the index of the PLU column (None: not chosen)
    # header: whether it has a header row; detected: whether a PLU heading was found
    sources = {
        m: {"columns": [], "body": [], "plu": None, "header": False, "detected": False}
        for m in ("file", "paste")
    }
    # path: the file the preview shows; stamp: its version; error: why it couldn't be read
    preview = {"path": "", "stamp": None, "error": "", "timer": None}

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
    def table_of(which):
        return file_rows if which == "file" else paste_rows

    def find_columns(which, keep_plu=False):
        """Work out an input's columns, going back to the assumed PLU column."""
        src, chosen = sources[which], sources[which]["plu"]
        src["columns"], src["body"], src["plu"] = plu_table(table_of(which).rows, src["header"])
        if keep_plu and any(i == chosen for i, _ in src["columns"]):
            src["plu"] = chosen

    def view(which):
        """The preview of an input, as RowsTable.show arguments."""
        src = sources[which]
        return (
            [label for _, label in src["columns"]], [i for i, _ in src["body"]],
            [i for i, _ in src["columns"]], plu_position(which),
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
        """Update the options, the mark in the preview and the Convert button."""
        which = mode.get()
        has_header.set(sources[which]["header"])
        header_info.configure(text="PLU heading found" if sources[which]["detected"] else "")
        labels = [label for _, label in sources[which]["columns"]]
        position = plu_position(which)
        plu_box.configure(values=labels, state="readonly" if labels else "disabled")
        plu_var.set("" if position is None else labels[position])
        table_of(which).mark(position)
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
    file_rows = RowsTable(file_box, 6, lambda: file_edited(), pick)
    file_rows.frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
    file_foot = ttk.Frame(file_box)
    file_foot.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))
    file_foot.columnconfigure(0, weight=1)
    file_info = ttk.Label(file_foot)
    file_info.grid(row=0, column=0, sticky="w")
    file_rows.buttons(file_foot).grid(row=0, column=1)

    def show_file_info():
        text, style = preview["error"], "Warn.TLabel"
        if not text:
            style = "Hint.TLabel"
            if os.path.isfile(preview["path"]):
                text = (
                    f"{_plural(len(file_rows.rows), 'row')} in the preview. "
                    "Double-click a value to change it (the file itself isn't changed)."
                )
            elif root.can_drop:
                text = "You can also drop a file onto this window."
        file_info.configure(text=text, style=style)

    def file_stamp(path):
        try:
            info = os.stat(path)
        except OSError:
            return None
        return info.st_mtime_ns, info.st_size

    def file_edited():
        find_columns("file", keep_plu=True)
        file_rows.show(*view("file"), keep_place=True)
        show_file_info()
        show_plu()

    def load_preview():
        """Read the chosen file and show its rows."""
        if preview["timer"]:
            root.after_cancel(preview["timer"])
        path = in_var.get().strip()
        rows, error, stamp = [], "", file_stamp(path)
        if os.path.isfile(path):
            try:
                rows = [r for r in read_rows(path) if any(str(c).strip() for c in r)]
            except PermissionError:
                error = "Could not open the file. Is it open in Excel?"
            except Exception as e:  # noqa: BLE001
                error = f"Could not read the file: {_one_line(e)}"
        preview.update(path=path, stamp=stamp, error=error, timer=None)
        file_rows.load(rows)
        sources["file"]["header"] = sources["file"]["detected"] = has_plu_header(rows)
        find_columns("file")
        file_rows.show(*view("file"))
        show_file_info()
        show_plu()

    def input_changed(*_):
        # Wait for a pause in typing before reading the file
        if preview["timer"]:
            root.after_cancel(preview["timer"])
        preview["timer"] = root.after(250, load_preview)
        update_ready()

    # --- paste input
    def on_paste_change(edited):
        src, found = sources["paste"], has_plu_header(paste_rows.rows)
        if not edited and found != src["detected"]:  # otherwise keep the user's choice
            src["header"] = src["detected"] = found
        find_columns("paste", keep_plu=edited)
        show_plu()

    paste_box, paste_rows, focus_paste = _paste_box(
        body, root, on_paste_change, lambda: view("paste"), on_pick=pick
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
        which = mode.get()
        table_of(which).finish()
        sources[which]["header"] = has_header.get()
        find_columns(which)
        table_of(which).show(*view(which))
        show_plu()

    header_row = ttk.Frame(body)
    header_row.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    ttk.Checkbutton(
        header_row, text="Data has a header row", variable=has_header, command=header_changed
    ).pack(side="left")
    header_info = ttk.Label(header_row, style="Hint.TLabel")
    header_info.pack(side="left", padx=8)
    ttk.Label(body, text="PLU column:").grid(row=4, column=0, sticky="w", pady=(8, 0))
    plu_row = ttk.Frame(body)
    plu_row.grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=(8, 0))
    plu_box = ttk.Combobox(plu_row, textvariable=plu_var, state="disabled", width=30, height=15)
    plu_box.pack(side="left")
    plu_box.bind("<<ComboboxSelected>>", lambda e: pick(plu_box.current()))
    ttk.Label(plu_row, text="or click a column heading in the preview", style="Hint.TLabel").pack(
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
        table = table_of(mode.get())
        table.finish()
        if update_ready():
            return  # the value just changed left nothing to convert
        qty = qty_var.get().strip()
        pasting = mode.get() == "paste"
        src = in_var.get().strip()
        if not pasting and file_stamp(src) != preview["stamp"]:
            if messagebox.askyesno(
                APP_NAME,
                "The file has changed since it was opened here.\nLoad it again"
                + (" (the changes you made in the preview will be lost)" if table.undo_steps else "")
                + "?\n\nChoose No to convert what the preview shows.",
            ):
                load_preview()
                return
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
            n, skipped = extract_rows(
                table.rows, out, sources[mode.get()]["header"], qty, plu_col,
                source="pasted data" if pasting else "file",
            )
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
            if not sources["paste"]["columns"]:
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
            reason = "There are no rows"
        if not reason and sources[mode.get()]["plu"] is None:
            reason = "Choose the PLU column"
        if not reason:
            try:
                if not math.isfinite(float(qty_var.get().strip())):
                    raise ValueError
            except ValueError:
                reason = "QTY must be a number"
        set_reason(reason)
        return reason

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
    group_var = tk.BooleanVar(value=True)
    # selected holds indexes into sections; sort is (column or None, descending);
    # shown lists the indexes that pass the filter, in display order; categories lists
    # the category names in file order and collapsed the indexes of the folded ones
    state = {
        "sections": [], "selected": set(), "sort": (None, False), "shown": [],
        "categories": [], "collapsed": set(),
    }

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
    group_check = ttk.Checkbutton(
        filter_row, text="Group by category", variable=group_var, command=lambda: refresh_view()
    )
    group_check.grid(row=0, column=4, padx=(12, 0))

    list_box = ttk.Frame(body)
    list_box.grid(row=3, column=0, sticky="nsew")
    list_box.columnconfigure(0, weight=1)
    list_box.rowconfigure(0, weight=1)
    tree = ttk.Treeview(
        list_box, columns=("check", "section", "supplier", "items", "preview"), height=8,
        selectmode="none", displaycolumns=("section", "supplier", "items", "preview"),
    )
    _check_column(tree, root, lambda: set_all(not state["shown"] or not all_shown_selected()))
    headings = {"section": ("Section", "w"), "supplier": ("Supplier", "w"), "items": ("Items", "center")}
    tree.column("section", width=_px(420), minwidth=_px(120), anchor="w")
    tree.column("supplier", width=_px(170), minwidth=_px(60), stretch=False, anchor="w")
    tree.column("items", width=_px(80), minwidth=_px(60), stretch=False, anchor="center")
    tree.column("preview", width=_px(90), minwidth=_px(70), stretch=False, anchor="center")
    tree.heading("preview", text="Preview", anchor="center")
    tree.grid(row=0, column=0, sticky="nsew")
    _scrollbar(tree, "vertical", 0, 1)
    def recolor(colors):
        tree.tag_configure("on", background=colors["on"])
        tree.tag_configure("group", background=colors["section"], font=("Segoe UI", 10, "bold"))

    root.on_theme.append(recolor)
    def category_of(idx):
        """Index into state["categories"] of the group a section is listed under."""
        return state["categories"].index(bw_group(state["sections"][idx]))

    def grouped():
        return group_var.get() and any(s.category for s in state["sections"])

    def group_ids(cat):
        """A category's shown sections."""
        return [i for i in state["shown"] if category_of(i) == cat]

    def box_state(ids):
        on = sum(i in state["selected"] for i in ids)
        if not ids:
            return BLOCKED
        return CHECKED if on == len(ids) else PARTLY if on else UNCHECKED

    def refresh_row(idx):
        if not tree.exists(str(idx)):
            return
        on = idx in state["selected"]
        box = CHECKED if on else UNCHECKED
        tree.set(str(idx), "check", box)
        tree.item(str(idx), image=root.checks[box], tags=("on",) if on else ())

    def refresh_group(cat):
        iid = f"c{cat}"
        if not tree.exists(iid):
            return
        ids = group_ids(cat)
        box = box_state(ids)
        arrow = "▸" if cat in state["collapsed"] else "▾"
        picked = sum(i in state["selected"] for i in ids)
        tree.item(iid, image=root.checks[box], tags=("group",), values=(
            box, f"{arrow}  {state['categories'][cat]}", f"{picked} of {len(ids)} selected",
            sum(len(state["sections"][i].items) for i in ids), "",
        ))

    def refresh_groups(ids=None):
        cats = range(len(state["categories"])) if ids is None else {category_of(i) for i in ids}
        for cat in cats:
            refresh_group(cat)

    def all_shown_selected():
        return state["selected"].issuperset(state["shown"])

    def update_sel_info():
        secs, sel = state["sections"], state["selected"]
        if not secs:
            sel_info.set("")
            set_reason("Choose a flight file")
            return
        items = sum(len(secs[i].items) for i in sel)
        shown = set(state["shown"])
        hidden = len(sel - shown)
        if not shown:
            box = BLOCKED
        elif all_shown_selected():
            box = CHECKED
        else:
            box = PARTLY if sel & shown else UNCHECKED
        tree.heading("#0", image=root.heading_checks[box])
        sel_info.set(
            f"{len(sel)} of {_plural(len(secs), 'section')} selected"
            + (f" ({_plural(items, 'item')})" if sel else "")
            + (f", including {hidden} hidden by the filter" if hidden else "")
        )
        set_reason("" if sel else "Select at least one section")

    def toggle_ids(ids):
        """Tick the sections, or clear them if they are all ticked already."""
        on = not state["selected"].issuperset(ids)
        (state["selected"].update if on else state["selected"].difference_update)(ids)
        for i in ids:
            refresh_row(i)
        refresh_groups(ids)
        update_sel_info()

    def toggle(idx):
        toggle_ids([idx])

    def toggle_group(cat):
        toggle_ids(group_ids(cat))

    def fold(cat, open_it=None):
        """Fold a category's sections away or show them again (open_it: force either)."""
        folded = cat in state["collapsed"]
        if open_it is not None and open_it != folded:
            return  # already open, or already folded
        state["collapsed"].symmetric_difference_update({cat})
        refresh_view()
        tree.focus(f"c{cat}")
        tree.see(f"c{cat}")

    def sort_key(col, idx):
        section = state["sections"][idx]
        if col == "items":
            return len(section.items)
        text = section.title if col == "section" else section_supplier(section)
        return _natural_key(text)

    def passes_filter(section):
        supplier = supplier_var.get()
        # A section with several suppliers shows under each of them
        if supplier != ALL_SUPPLIERS and supplier.upper() not in {
            s.upper() for s in section_suppliers(section) or [NO_SUPPLIER]
        }:
            return False
        text = f"{section.title} {section.category}".upper()
        return all(word in text for word in search_var.get().upper().split())

    def refresh_view(*_):
        """List the sections that pass the filter, in the chosen order (default: as in the
        file), under their categories if grouped."""
        col, descending = state["sort"]
        order = [i for i, s in enumerate(state["sections"]) if passes_filter(s)]
        if col:
            order.sort(key=lambda i: sort_key(col, i), reverse=descending)
        state["shown"] = order
        focus = tree.focus()
        tree.delete(*tree.get_children())

        def add(idx):
            s = state["sections"][idx]
            values = ("", s.title, section_supplier(s), len(s.items), "View…")
            tree.insert("", "end", iid=str(idx), values=values)
            refresh_row(idx)

        if grouped():
            members = {}
            for idx in order:
                members.setdefault(category_of(idx), []).append(idx)
            for cat in sorted(members):
                tree.insert("", "end", iid=f"c{cat}")
                refresh_group(cat)
                if cat not in state["collapsed"]:
                    for idx in members[cat]:
                        add(idx)
        else:
            for idx in order:
                add(idx)
        if focus and tree.exists(focus):
            tree.focus(focus)
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

    def column_at(x):
        column = tree.identify_column(x)
        return tree.column(column, "id") if column not in ("", "#0") else column

    def preview(idx):
        _section_preview(root, state["sections"][idx])

    def on_click(event):
        if tree.identify_region(event.x, event.y) == "heading":
            return None  # let the heading handle it (sorts the list)
        row = tree.identify_row(event.y)
        if row and tree.identify_region(event.x, event.y) in ("cell", "tree"):
            tree.focus(row)
            tree.focus_set()
            column = column_at(event.x)
            if row.startswith("c"):
                # Category row: its tick box ticks all its sections, elsewhere folds it
                (toggle_group if column == "#0" else fold)(int(row[1:]))
            else:
                (preview if column == "preview" else toggle)(int(row))
        return "break"

    def on_motion(event):
        # After _check_column's handler, which clears the cursor elsewhere
        row = tree.identify_row(event.y)
        if row and not row.startswith("c") and column_at(event.x) == "preview":
            tree.configure(cursor="hand2")

    def on_key(section_action=None, group_action=None):
        """Key handler that runs the action for the focused section or category row.
        Keys with no action for that kind of row keep their usual meaning."""
        def handler(_):
            row = tree.focus()
            if row.startswith("c"):
                if not group_action:
                    return None
                group_action(int(row[1:]))
            elif row and int(row) in state["shown"]:
                if not section_action:
                    return None
                section_action(int(row))
            return "break"
        return handler

    tree.bind("<Button-1>", on_click)
    tree.bind("<space>", on_key(toggle, toggle_group))
    tree.bind("<Motion>", on_motion, add="+")
    for key in ("p", "P"):
        tree.bind(f"<KeyPress-{key}>", on_key(preview, lambda c: None))
    tree.bind("<Left>", on_key(group_action=lambda c: fold(c, open_it=False)))
    tree.bind("<Right>", on_key(group_action=lambda c: fold(c, open_it=True)))

    def set_all(on):
        """Select or clear the sections that are showing."""
        (state["selected"].update if on else state["selected"].difference_update)(state["shown"])
        for i in state["shown"]:
            refresh_row(i)
        refresh_groups()
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
        state["sections"], state["sort"] = sections, (None, False)
        # Known categories in file order, then the rest
        groups = list(dict.fromkeys(bw_group(s) for s in sections))
        state["categories"] = [g for g in groups if g != BW_OTHER_CATEGORIES] + [
            g for g in groups if g == BW_OTHER_CATEGORIES
        ]
        if any(_known_category(s.category) for s in sections):
            state["selected"] = {i for i, s in enumerate(sections) if bw_default_selected(s)}
        else:
            state["selected"] = set(range(len(sections)))  # no known categories: tick all
        state["collapsed"] = set()
        (group_check.grid if any(s.category for s in sections) else group_check.grid_remove)()
        suppliers = {}
        for s in sections:
            for supplier in section_suppliers(s) or [NO_SUPPLIER]:
                suppliers.setdefault(supplier.upper(), supplier)
        supplier_box.configure(values=(ALL_SUPPLIERS, *(suppliers[k] for k in sorted(suppliers))))
        search_var.set("")
        supplier_var.set(ALL_SUPPLIERS)
        cats = len([g for g in state["categories"] if g != BW_OTHER_CATEGORIES])
        found_var.set(
            f"{_plural(len(sections), 'section')} found"
            + (f" in {cats} {'category' if cats == 1 else 'categories'}" if cats else "")
            + ". Click a section to include or leave it out, View… (or P) to see its items, "
            "or a heading to sort:"
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
    group_check.grid_remove()  # until a file with categories is loaded
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
        "file": None,  # StockTable from the chosen file
        "paste": None,  # StockTable from the pasted cells
        "file_error": "",
        "paste_error": "",
    }

    ttk.Label(body, textvariable=chosen_info).grid(row=1, column=0, columnspan=3, sticky="w")

    # --- stock report source
    ttk.Label(body, text="Stock report:").grid(row=2, column=0, sticky="w", pady=(10, 0))
    src_row = ttk.Frame(body)
    src_row.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def table_of(which):
        return file_rows if which == "file" else paste_rows

    def parse(which):
        """Read the items from an input's rows."""
        state[which], state[which + "_error"] = None, ""
        rows = table_of(which).rows
        if any(any(str(c).strip() for c in r) for r in rows):
            try:
                state[which] = parse_stock_rows(rows, "preview" if which == "file" else "pasted data")
            except ExtractError as e:
                state[which + "_error"] = str(e)
        elif which == "file" and in_var.get():
            state["file_error"] = "All the rows have been deleted."

    def view(which):
        """The labelled columns of an input's items, in report order, as RowsTable.show arguments."""
        table = state[which]
        if table is None:
            return _raw_table(table_of(which).rows)
        cols = {table.plu_col: TARGET_HEADER, table.soh_col: SOH_HEADER}
        cols.update((ci, label) for label, ci in table.columns.items())
        order = sorted(cols)
        return [cols[c] for c in order], [i for i, _ in table.body], order

    file_box = ttk.Frame(body)
    file_box.columnconfigure(1, weight=1)
    file_box.rowconfigure(1, weight=1)
    ttk.Label(file_box, text="Stock file:").grid(row=0, column=0, sticky="w")
    in_entry = _path_entry(file_box, in_var)
    in_entry.grid(row=0, column=1, sticky="ew", padx=6)
    file_rows = RowsTable(file_box, 5, lambda: file_edited())
    file_rows.frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
    file_foot = ttk.Frame(file_box)
    file_foot.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))
    file_foot.columnconfigure(0, weight=1)
    file_hint = ttk.Label(file_foot, style="Hint.TLabel")
    file_hint.grid(row=0, column=0, sticky="w")
    file_rows.buttons(file_foot).grid(row=0, column=1)

    def file_edited():
        parse("file")
        file_rows.show(*view("file"), keep_place=True)
        refresh_table()

    def browse_input():
        path = filedialog.askopenfilename(title="Select the stock report", filetypes=FILE_TYPES)
        if path:
            load(path)

    def load(path):
        try:
            rows = read_rows(path)
            parse_stock_rows(rows)
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
        file_rows.load(rows)
        parse("file")
        file_rows.show(*view("file"))
        file_hint.configure(
            text="Double-click a value to change it (the file itself isn't changed)."
        )
        mode.set("file")
        switch_mode()

    ttk.Button(file_box, text="Browse…", command=browse_input).grid(row=0, column=2)

    def on_paste_change(_):
        parse("paste")
        refresh_table()

    paste_box, paste_rows, focus_paste = _paste_box(
        body, root, on_paste_change, lambda: view("paste"), height=5
    )

    for box in (file_box, paste_box):
        box.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
    paste_box.grid_remove()

    def current_table():
        return state[mode.get()]

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
            text = state[mode.get() + "_error"]
            if not text and mode.get() == "file" and root.can_drop:
                text = "You can also drop the file onto this window."
            stock_info.set(text)
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
            if state[mode.get() + "_error"]:
                reason = "The stock report has no items to use"
        elif planogram_var.get() not in table.columns:
            reason = "Select the planogram column"
        threshold = threshold_var.get().strip()
        if not reason and (not threshold or _to_number(threshold) is None):
            reason = "Threshold must be a number"
        set_reason(reason)
        return reason

    def next_page():
        table_of(mode.get()).finish()
        if update_ready():
            return  # the value just changed left nothing to use
        go_next(
            current_table(), planogram_var.get(), _to_number(threshold_var.get()),
            in_var.get() if mode.get() == "file" else "",
        )

    run, set_reason = _footer(frm, "Next →", next_page, go_back)
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
    body.rowconfigure(2, weight=1)

    ALL_SECTIONS = "All sections"
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
    search_var = tk.StringVar()
    section_var = tk.StringVar(value=ALL_SECTIONS)
    group_var = tk.BooleanVar(value=True)
    HINT = "Click a box to include or leave out an item. Click an order quantity to change it."
    state = {
        "flight": "", "sections": [], "inputs": None, "items": [],
        "sources": [],  # input files, which must not be overwritten
        "shown": [],  # indexes of the items that pass the filters
        "collapsed": set(),  # indexes of the sections whose items are folded away
        "sort": (None, False),  # column or None (as in the flight), descending
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

    find_row = ttk.Frame(body)
    find_row.grid(row=1, column=0, sticky="ew", pady=(0, 6))
    find_row.columnconfigure(1, weight=1)
    ttk.Label(find_row, text="Search:").grid(row=0, column=0)
    search_entry = ttk.Entry(find_row, textvariable=search_var)
    search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 12))
    search_entry.bind("<Return>", lambda e: "break")  # filter only, don't generate
    ttk.Label(find_row, text="Section:").grid(row=0, column=2)
    section_box = ttk.Combobox(
        find_row, textvariable=section_var, state="readonly", width=32, values=(ALL_SECTIONS,)
    )
    section_box.grid(row=0, column=3, padx=(6, 12))
    ttk.Checkbutton(
        find_row, text="Group by section", variable=group_var, command=lambda: refresh()
    ).grid(row=0, column=4)

    # --- items list
    list_box = ttk.Frame(body)
    list_box.grid(row=2, column=0, sticky="nsew")
    list_box.columnconfigure(0, weight=1)
    list_box.rowconfigure(0, weight=1)
    columns = (
        ("check", "", 0, "center"),  # not displayed: the tick box is in column #0
        ("desc", "Item", 270, "w"), ("section", "Section", 160, "w"),
        ("plu", "PLU", 75, "w"), ("min", "MIN QTY", 80, "center"),
        ("plano", "Planogram", 90, "center"), ("soh", "SOH", 60, "center"),
        ("qty", "Order", 70, "center"), ("note", "Notes", 240, "w"),
    )
    tree = ttk.Treeview(
        list_box, columns=[c[0] for c in columns], height=8, selectmode="browse"
    )
    _check_column(tree, root, lambda: toggle_items(state["shown"]))
    for name, label, width, anchor in columns:
        tree.heading(name, text=label, anchor=anchor, command=lambda c=name: sort_by(c))
        tree.column(
            name, width=_px(width), minwidth=_px(40), stretch=name in ("desc", "note"), anchor=anchor
        )
    tree.grid(row=0, column=0, sticky="nsew")
    # An open entry doesn't scroll with the list, so close it
    _scrollbar(tree, "vertical", 0, 1, lambda: finish_edit())
    _scrollbar(tree, "horizontal", 1, 0, lambda: finish_edit())

    def recolor(colors):
        tree.tag_configure("on", background=colors["on"])
        tree.tag_configure("off", foreground=colors["off"])
        tree.tag_configure("section", background=colors["section"], font=("Segoe UI", 10, "bold"))

    root.on_theme.append(recolor)

    info_row = ttk.Frame(body)
    info_row.grid(row=3, column=0, sticky="ew", pady=(6, 0))
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
        indent = "      " if group_var.get() else ""
        return (
            CHECKED if item.include else UNCHECKED, indent + _one_line(item.description),
            state["sections"][item.section].title, item.plu, item.min_qty, item.planogram, item.soh,
            "" if item.qty is None else _format_qty(item.qty), "; ".join(notes),
        )

    def draw_row(idx):
        item = state["items"][idx]
        if tree.exists(str(idx)):
            box = CHECKED if item.include else BLOCKED if item.qty is None else UNCHECKED
            tree.item(
                str(idx), values=row_values(item), image=root.checks[box],
                tags=("on" if item.include else "off",),
            )

    def refresh_row(idx):
        draw_row(idx)
        refresh_section(state["items"][idx].section)

    def box_items(ids):
        """Of the items with these indexes: (those a box for them all ticks, those ticked)."""
        items = [state["items"][i] for i in ids]
        # The box adds what needs ordering or was given a quantity, not what isn't needed
        wanted = [
            i for i in items if i.qty is not None and (i.status == ORDER or i.qty != i.calculated)
        ]
        return wanted, [i for i in items if i.include]

    def box_state(ids):
        wanted, ticked = box_items(ids)
        if not ticked:
            return UNCHECKED if wanted else BLOCKED
        return CHECKED if all(i.include for i in wanted) else PARTLY

    def section_ids(section):
        """A section's shown items."""
        return [i for i in state["shown"] if state["items"][i].section == section]

    def refresh_section(section):
        """Update a section's heading row."""
        iid = f"s{section}"
        if not tree.exists(iid):
            return
        box = box_state(section_ids(section))
        items = [i for i in state["items"] if i.section == section]
        arrow = "▸" if section in state["collapsed"] else "▾"
        values = dict.fromkeys((c[0] for c in columns), "")
        values.update(
            check=box, desc=f"{arrow}  {state['sections'][section].title}",
            note=f"{sum(i.include for i in items)} of {len(items)} to order",
        )
        tree.item(
            iid, values=list(values.values()), image=root.checks[box], tags=("section",)
        )

    def update_totals():
        items = state["items"]
        for name, (label, test) in FILTERS.items():
            filter_btns[name].configure(text=f"{label} ({sum(1 for i in items if test(i))})")
        tree.heading("#0", image=root.heading_checks[box_state(state["shown"])])
        chosen = [i for i in items if i.include]
        files = len({i.section for i in chosen})
        hidden = sum(1 for i in chosen if not found(i))
        total_var.set(
            f"{_plural(len(chosen), 'item')} to order in {_plural(files, 'file')}"
            + (f", including {hidden} hidden by the search" if hidden else "")
        )
        if not chosen:
            set_reason("Tick at least one item")
        elif not out_var.get().strip():
            set_reason("Choose an output folder")
        else:
            set_reason("")

    def found(item):
        """Whether an item passes the search and the section filter."""
        if section_var.get() != ALL_SECTIONS and section_box.current() - 1 != item.section:
            return False
        text = f"{item.plu} {item.description}".upper()
        return all(word in text for word in search_var.get().upper().split())

    def sort_items(ids):
        col, descending = state["sort"]
        if col is None:
            return ids
        items, sections = state["items"], state["sections"]

        def key(idx):
            """None for blanks and non-numbers in number columns, which go last either way."""
            item = items[idx]
            if col == "qty":
                return item.qty
            if col in ("min", "plano", "soh"):
                text = {"min": item.min_qty, "plano": item.planogram, "soh": item.soh}[col]
                return _to_number(text) if text.strip() else None
            return _natural_key({
                "desc": item.description, "section": sections[item.section].title,
                "plu": item.plu, "note": row_values(item)[-1],
            }[col]) or None

        keys = {i: key(i) for i in ids}
        blank = [i for i in ids if keys[i] is None]
        rest = sorted((i for i in ids if keys[i] is not None), key=keys.get, reverse=descending)
        return rest + blank

    def sort_by(col):
        prev, descending = state["sort"]
        state["sort"] = (col, col == prev and not descending)
        refresh()

    def refresh(*_):
        """List the items that pass the filters, under their sections if grouped."""
        finish_edit()
        tree.delete(*tree.get_children())
        test, grouped = FILTERS[filter_var.get()][1], group_var.get()
        shown = [i for i, item in enumerate(state["items"]) if test(item) and found(item)]
        state["shown"] = shown = sort_items(shown)
        tree.configure(displaycolumns=[
            c[0] for c in columns if c[0] != "check" and not (grouped and c[0] == "section")
        ])
        for name, label, _, anchor in columns:
            arrow = (" ▼" if state["sort"][1] else " ▲") if name == state["sort"][0] else ""
            tree.heading(name, text=label + arrow, anchor=anchor)

        def add(ids):
            for idx in ids:
                tree.insert("", "end", iid=str(idx))
                draw_row(idx)

        if grouped:
            members = {}
            for idx in shown:
                members.setdefault(state["items"][idx].section, []).append(idx)
            for section in sorted(members):
                tree.insert("", "end", iid=f"s{section}")
                refresh_section(section)
                if section not in state["collapsed"]:
                    add(members[section])
        else:
            add(shown)
        tree.yview_moveto(0)
        update_totals()

    def fold(section):
        """Fold a section's items away, or show them again."""
        place = tree.yview()[0]
        state["collapsed"].symmetric_difference_update({section})
        refresh()
        tree.update_idletasks()
        tree.yview_moveto(place)
        tree.selection_set(f"s{section}")
        tree.focus(f"s{section}")

    def toggle_items(ids):
        """Tick or untick several items, as the box for them all would."""
        finish_edit()
        wanted, ticked = box_items(ids)
        tick = any(not i.include for i in wanted)
        for item in wanted if tick else ticked:
            item.include = tick
        hint()
        for idx in ids:
            draw_row(idx)
        for section in {state["items"][idx].section for idx in ids}:
            refresh_section(section)
        update_totals()

    def toggle_section(section):
        toggle_items(section_ids(section))

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
        editor.open(str(idx), "qty", "" if item.qty is None else _format_qty(item.qty))
        hint("Type the quantity and press Enter. Clear it to leave the item out.")

    def edited(iid, _, text):
        """Take the quantity typed. False if it isn't usable."""
        hint()
        if text is None:
            return True
        item = state["items"][int(iid)]
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
        refresh_row(int(iid))
        update_totals()
        return True

    editor = CellEditor(tree, edited, "center")

    def finish_edit():
        """Apply the quantity being changed, if any. False if it isn't usable."""
        return editor.close()

    def on_click(event):
        region = tree.identify_region(event.x, event.y)
        row = tree.identify_row(event.y)
        if region not in ("cell", "tree") or not row:
            return None
        finish_edit()
        number = int(tree.identify_column(event.x)[1:])
        col = tree["displaycolumns"][number - 1] if number else "check"
        tree.selection_set(row)
        tree.focus(row)
        tree.focus_set()
        if row.startswith("s"):
            (toggle_section if col == "check" else fold)(int(row[1:]))
        elif col == "check":
            toggle(int(row))
        elif col == "qty":
            start_edit(int(row))
        return "break"

    def on_key(item_action, section_action):
        def handler(_):
            row = tree.focus()
            if row and tree.exists(row):
                if row.startswith("s"):
                    section_action(int(row[1:]))
                else:
                    item_action(int(row))
            return "break"
        return handler

    tree.bind("<Button-1>", on_click)
    tree.bind("<space>", on_key(toggle, toggle_section))
    tree.bind("<Return>", on_key(start_edit, fold))
    tree.bind("<F2>", on_key(start_edit, fold))

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
    out_row.grid(row=4, column=0, sticky="ew", pady=(12, 0))
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
        search_var.set("")
        section_var.set(ALL_SECTIONS)
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
    for var in (search_var, section_var):
        var.trace_add("write", refresh)
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
        state["collapsed"], state["sort"] = set(), (None, False)
        titles = []
        for s in sections:
            same = sum(1 for other in sections[: len(titles)] if other.title == s.title)
            titles.append(f"{s.title} ({same + 1})" if same else s.title)
        section_box.configure(values=(ALL_SECTIONS, *titles))
        section_var.set(ALL_SECTIONS)
        search_var.set("")
        hint()
        refresh()
        return True

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
    from tkinter import messagebox, ttk

    root = tk.Tk()
    root.withdraw()  # until it is laid out
    root.title(APP_NAME)
    set_window_icon(root)

    global _scale
    _scale = root.winfo_fpixels("1i") / 96
    root.on_theme, root.dark = [], False
    apply_theme(root)

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
    pages["home"] = build_home_page(container, root, show)
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
    apply_theme(root)  # again, for the widgets made since

    # One size for all pages, so the buttons stay put when moving between them
    root.update_idletasks()
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    min_w = min(max(_px(900), pages["plu"].frame.winfo_reqwidth()), screen_w)
    min_h = min(max(p.frame.winfo_reqheight() for p in pages.values()), screen_h - 80)
    w, h = min(max(_px(1060), min_w), screen_w), min(max(_px(700), min_h), screen_h - 80)
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

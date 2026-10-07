"""
Bookings_merge_ui.py
Desktop app: add one or more files (.tsv, .csv, .xlsx), pick the columns to keep,
remove duplicate rows, and merge everything into ONE Excel sheet.

NO INSTALLS: Python standard library only. The .xlsx is written directly
(no PowerShell or Excel automation), then opened in Excel.
"""

import csv
import os
import re
import threading
import tkinter as tk
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]
EXCEL_EPOCH = datetime(1899, 12, 30)
SOURCE_COL = "Source file"


def downloads_folder():
    """Ask Windows for the real Downloads folder (handles redirected folders)."""
    try:
        import ctypes, uuid
        guid = (ctypes.c_char * 16).from_buffer_copy(
            uuid.UUID("374DE290-123F-4565-9164-39C4925E467B").bytes_le)
        ptr = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(ptr)) == 0:
            path = ptr.value
            ctypes.windll.ole32.CoTaskMemFree(ptr)
            return Path(path)
    except Exception:
        pass
    return Path.home() / "Downloads"


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def dedupe_names(names):
    seen, out = set(), []
    for h in names:
        name = (h or "Column").replace("\n", " ").strip() or "Column"
        base, n = name, 2
        while name.lower() in seen:
            name = "{} ({})".format(base, n)
            n += 1
        seen.add(name.lower())
        out.append(name)
    return out


# ---------------- Readers ----------------

def read_text_table(path):
    """.tsv / .csv / .txt -> (headers, rows as dicts)"""
    ext = Path(path).suffix.lower()
    raw = None
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            with open(path, newline="", encoding=enc) as f:
                raw = f.read()
            if enc == "utf-16" and "\t" not in raw and "," not in raw:
                continue
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    if raw is None:
        raise ValueError("Couldn't read the text encoding of this file.")

    if ext == ".tsv":
        delim = "\t"
    elif ext == ".csv":
        delim = ","
    else:
        first = raw.splitlines()[0] if raw else ""
        delim = "\t" if first.count("\t") >= first.count(",") else ","

    reader = csv.reader(raw.splitlines(), delimiter=delim)
    table = [r for r in reader if any(c.strip() for c in r)]
    return table_to_dicts(table)


def table_to_dicts(table):
    if not table:
        raise ValueError("No data found in this file.")
    headers = dedupe_names(table[0])
    rows = []
    for r in table[1:]:
        r = list(r) + [""] * (len(headers) - len(r))
        row = {h: r[i] for i, h in enumerate(headers)}
        if any(str(v).strip() for v in row.values()):
            rows.append(row)
    if not rows:
        raise ValueError("The file has headers but no data rows.")
    return headers, rows


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def _is_date_format(code):
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "").lower()
    if not code or code == "general":
        return False
    return any(ch in code for ch in "dmyhs")


def read_xlsx(path):
    """.xlsx (first sheet) -> (headers, rows as dicts). Standard library only."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        # Workbook -> first sheet file
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet_el = next(el for el in wb.iter() if _local(el.tag) == "sheet")
        rid = next(v for k, v in sheet_el.attrib.items() if _local(k) == "id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(el.get("Target") for el in rels if el.get("Id") == rid)
        sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        sheet_path = os.path.normpath(sheet_path).replace("\\", "/")

        # Shared strings
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
                shared.append("".join(t.text or "" for t in si.iter() if _local(t.tag) == "t"))

        # Which cell styles are dates
        date_styles = set()
        if "xl/styles.xml" in names:
            st = ET.fromstring(z.read("xl/styles.xml"))
            custom = {}
            for el in st.iter():
                if _local(el.tag) == "numFmt":
                    custom[int(el.get("numFmtId"))] = el.get("formatCode")
            builtin_dates = set(range(14, 23)) | {45, 46, 47}
            cell_xfs = next((el for el in st if _local(el.tag) == "cellXfs"), None)
            if cell_xfs is not None:
                for i, xf in enumerate(x for x in cell_xfs if _local(x.tag) == "xf"):
                    fid = int(xf.get("numFmtId", "0"))
                    if fid in builtin_dates or (fid in custom and _is_date_format(custom[fid])):
                        date_styles.add(i)

        # Cells
        table = []
        sheet = ET.fromstring(z.read(sheet_path))
        for row_el in sheet.iter():
            if _local(row_el.tag) != "row":
                continue
            row = {}
            next_col = 0
            for c in row_el:
                if _local(c.tag) != "c":
                    continue
                ref = c.get("r")
                col = _col_index(ref) if ref else next_col
                next_col = col + 1
                t = c.get("t", "n")
                s = int(c.get("s", "0"))
                v_el = next((x for x in c if _local(x.tag) == "v"), None)
                v = v_el.text if v_el is not None else None

                if t == "s" and v is not None:
                    val = shared[int(v)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter() if _local(x.tag) == "t")
                elif t in ("str", "e"):
                    val = v or ""
                elif t == "b":
                    val = "TRUE" if v == "1" else "FALSE"
                elif v is None:
                    val = ""
                else:
                    num = float(v)
                    if s in date_styles:
                        val = EXCEL_EPOCH + timedelta(days=num)
                        val = val.replace(microsecond=0)
                    elif num.is_integer():
                        val = str(int(num))
                    else:
                        val = repr(num)
                row[col] = val
            if row:
                width = max(row) + 1
                table.append([row.get(i, "") for i in range(width)])

    table = [r for r in table if any(str(c).strip() for c in r)]
    if table:
        width = max(len(r) for r in table)
        table = [r + [""] * (width - len(r)) for r in table]
        table[0] = [str(h) for h in table[0]]
    return table_to_dicts(table)


def read_any(path):
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext == ".xls":
        raise ValueError("Old .xls files aren't supported. Open it in Excel and Save As .xlsx.")
    return read_text_table(path)


# ---------------- Merge ----------------

def norm(v):
    """Normalize a value for duplicate checking."""
    if isinstance(v, datetime):
        return v.isoformat()
    s = str(v).strip()
    d = parse_date(s)
    if d:
        return d.isoformat()
    return re.sub(r"\s+", " ", s).lower()


def merge(files, columns, dedupe, match_col, add_source):
    """files: list of dicts {name, map(lower->header), rows}. columns: display names."""
    out_headers = list(columns) + ([SOURCE_COL] if add_source else [])
    data, seen = [], set()
    total = dupes = 0
    for f in files:
        for row in f["rows"]:
            vals = []
            for c in columns:
                h = f["map"].get(c.strip().lower())
                v = row.get(h, "") if h else ""
                vals.append(v if isinstance(v, datetime) else str(v).strip())
            if not any(v != "" for v in vals):
                continue
            total += 1
            if dedupe:
                if match_col:
                    i = columns.index(match_col)
                    key = norm(vals[i]) if vals[i] != "" else None
                else:
                    key = tuple(norm(v) for v in vals)
                if key is not None:
                    if key in seen:
                        dupes += 1
                        continue
                    seen.add(key)
            if add_source:
                vals.append(f["name"])
            data.append(vals)

    # Date columns: every non-blank value is a date
    date_cols = set()
    for i in range(len(out_headers)):
        values = [r[i] for r in data if r[i] != ""]
        if values and all(isinstance(v, datetime) or parse_date(v) for v in values):
            date_cols.add(i)
    for r in data:
        for i, v in enumerate(r):
            if v == "":
                continue
            if i in date_cols:
                d = v if isinstance(v, datetime) else parse_date(v)
                r[i] = (d - EXCEL_EPOCH).total_seconds() / 86400
            elif isinstance(v, datetime):
                r[i] = v.strftime("%m/%d/%Y %I:%M %p")
    return out_headers, data, date_cols, total, dupes


# ---------------- Excel file writer (standard library only) ----------------

_BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _x(text):
    return escape(_BAD_XML.sub("", str(text)), {'"': "&quot;"})


def _col(n):
    """1 -> A, 27 -> AA"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _unique_headers(headers):
    seen, out = {}, []
    for h in headers:
        name = (h or "Column").replace("\n", " ").strip() or "Column"
        base, n = name, 2
        while name.lower() in seen:
            name = "{} ({})".format(base, n)
            n += 1
        seen[name.lower()] = True
        out.append(name)
    return out


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Appointments" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

SHEET_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"""

# Style 0 = normal, 1 = date/time, 2 = text
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="mm/dd/yyyy h:mm AM/PM"/></numFmts>
<fonts count="1"><font><sz val="10"/><name val="Arial"/><family val="2"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def write_xlsx(headers, data, date_cols, out_path):
    headers = _unique_headers(headers)
    n_cols, n_rows = len(headers), len(data) + 1
    ref = "A1:{}{}".format(_col(n_cols), n_rows)

    # Column widths
    widths = [len(h) for h in headers]
    for row in data:
        for i, v in enumerate(row):
            w = 20 if i in date_cols else len(str(v))
            widths[i] = max(widths[i], w)
    cols_xml = "".join('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(i + 1, min(w + 3, 60))
                       for i, w in enumerate(widths))

    def text_cell(ref_, value, style):
        return '<c r="{}" t="inlineStr" s="{}"><is><t xml:space="preserve">{}</t></is></c>'.format(
            ref_, style, _x(value))

    rows_xml = ['<row r="1">' + "".join(text_cell("{}1".format(_col(i + 1)), h, 0)
                                        for i, h in enumerate(headers)) + "</row>"]
    for r, row in enumerate(data, start=2):
        cells = []
        for i, v in enumerate(row):
            ref_ = "{}{}".format(_col(i + 1), r)
            if i in date_cols and v != "":
                cells.append('<c r="{}" s="1"><v>{}</v></c>'.format(ref_, v))
            elif v != "":
                cells.append(text_cell(ref_, v, 2))
        rows_xml.append('<row r="{}">{}</row>'.format(r, "".join(cells)))

    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             '<sheetViews><sheetView workbookViewId="0">'
             '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
             '</sheetView></sheetViews>'
             '<cols>' + cols_xml + '</cols>'
             '<sheetData>' + "".join(rows_xml) + '</sheetData>'
             '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
             '</worksheet>')

    table = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'id="1" name="Bookings" displayName="Bookings" ref="{0}" totalsRowShown="0">'
             '<autoFilter ref="{0}"/><tableColumns count="{1}">{2}</tableColumns>'
             '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
             'showRowStripes="1" showColumnStripes="0"/></table>').format(
        ref, n_cols, "".join('<tableColumn id="{}" name="{}"/>'.format(i + 1, _x(h))
                             for i, h in enumerate(headers)))

    with zipfile.ZipFile(str(out_path), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS)
        z.writestr("xl/tables/table1.xml", table)


# ---------------- UI ----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Merge Files to Excel")
        self.geometry("640x780")
        self.minsize(560, 680)

        self.files = []          # {path, name, headers, map, rows}
        self.all_cols = []       # display names, in order of first appearance
        default_out = downloads_folder() / "Merged_{:%Y%m%d}.xlsx".format(date.today())
        self.out_var = tk.StringVar(value=str(default_out))
        self.open_var = tk.BooleanVar(value=True)
        self.dedupe_var = tk.BooleanVar(value=True)
        self.source_var = tk.BooleanVar(value=False)
        self.match_var = tk.StringVar(value="All selected columns")
        self.status_var = tk.StringVar(value="Step 1: Add one or more files (.tsv, .csv, .xlsx)")
        self._worker = None
        self._result = None

        pad = {"padx": 12, "pady": 5}
        ttk.Style(self).configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)

        # 1. Files
        f1 = ttk.LabelFrame(self, text="1. Files to merge")
        f1.pack(fill="both", **pad)
        fl = ttk.Frame(f1)
        fl.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        self.file_list = tk.Listbox(fl, height=5, font=("Segoe UI", 10), activestyle="none",
                                    selectmode="extended", exportselection=False)
        fsb = ttk.Scrollbar(fl, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=fsb.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        fsb.pack(side="right", fill="y")
        fb = ttk.Frame(f1)
        fb.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(fb, text="Add files…", command=self.add_files).pack(side="left")
        ttk.Button(fb, text="Remove selected", command=self.remove_files).pack(side="left", padx=6)
        ttk.Button(fb, text="Clear all", command=self.clear_files).pack(side="left")

        # 2. Columns
        f2 = ttk.LabelFrame(self, text="2. Columns to keep (click to select / unselect)")
        f2.pack(fill="both", expand=True, **pad)
        cb = ttk.Frame(f2)
        cb.pack(fill="x", padx=8, pady=(6, 0))
        ttk.Button(cb, text="Select all", command=lambda: self.set_all_cols(True)).pack(side="left")
        ttk.Button(cb, text="Clear", command=lambda: self.set_all_cols(False)).pack(side="left", padx=6)
        self.col_count = ttk.Label(cb, text="")
        self.col_count.pack(side="right")
        cl = ttk.Frame(f2)
        cl.pack(fill="both", expand=True, padx=8, pady=8)
        self.cols = tk.Listbox(cl, selectmode="multiple", exportselection=False,
                               font=("Segoe UI", 10), activestyle="none")
        csb = ttk.Scrollbar(cl, orient="vertical", command=self.cols.yview)
        self.cols.configure(yscrollcommand=csb.set)
        self.cols.pack(side="left", fill="both", expand=True)
        csb.pack(side="right", fill="y")
        self.cols.bind("<<ListboxSelect>>", lambda e: self.refresh_match_options())

        # 3. Duplicates
        f3 = ttk.LabelFrame(self, text="3. Duplicates")
        f3.pack(fill="x", **pad)
        ttk.Checkbutton(f3, text="Show duplicate rows only once", variable=self.dedupe_var,
                        command=self.toggle_match).pack(anchor="w", padx=8, pady=(6, 2))
        mrow = ttk.Frame(f3)
        mrow.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(mrow, text="A row is a duplicate when this matches:").pack(side="left")
        self.match_box = ttk.Combobox(mrow, textvariable=self.match_var, state="readonly", width=30)
        self.match_box.pack(side="left", padx=6)
        ttk.Checkbutton(f3, text="Add a \"Source file\" column", variable=self.source_var).pack(
            anchor="w", padx=8, pady=(0, 6))

        # 4. Output
        f4 = ttk.LabelFrame(self, text="4. Save merged file as")
        f4.pack(fill="x", **pad)
        ttk.Entry(f4, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(f4, text="Browse…", command=self.pick_out).pack(side="right", padx=8)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        ttk.Checkbutton(bottom, text="Open in Excel when done", variable=self.open_var).pack(side="left")
        self.run_btn = ttk.Button(bottom, text="Merge to Excel", style="Big.TButton",
                                  command=self.run, state="disabled")
        self.run_btn.pack(side="right")

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12)
        ttk.Label(self, textvariable=self.status_var, foreground="#444", wraplength=600).pack(
            anchor="w", padx=12, pady=(4, 10))

        self.refresh_match_options()
        self.after(200, self.add_files)

    # ----- files -----
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Add files to merge",
            initialdir=str(downloads_folder()),
            filetypes=[("Supported files", "*.tsv *.csv *.xlsx *.xlsm *.txt"),
                       ("Excel", "*.xlsx *.xlsm"), ("Bookings export", "*.tsv"),
                       ("CSV", "*.csv"), ("All files", "*.*")],
        )
        errors = []
        already = {f["path"] for f in self.files}
        for p in paths:
            if p in already:
                continue
            try:
                headers, rows = read_any(p)
            except Exception as e:
                errors.append("{}: {}".format(Path(p).name, e))
                continue
            self.files.append({"path": p, "name": Path(p).name, "headers": headers,
                               "map": {h.strip().lower(): h for h in headers}, "rows": rows})
        if errors:
            messagebox.showerror("Some files couldn't be read", "\n\n".join(errors))
        self.rebuild()

    def remove_files(self):
        for i in sorted(self.file_list.curselection(), reverse=True):
            del self.files[i]
        self.rebuild()

    def clear_files(self):
        self.files = []
        self.rebuild()

    def rebuild(self):
        selected_before = {self.all_cols[i] for i in self.cols.curselection()}
        old_cols = set(self.all_cols)

        self.file_list.delete(0, "end")
        for f in self.files:
            self.file_list.insert("end", "{}   ({} rows, {} columns)".format(
                f["name"], len(f["rows"]), len(f["headers"])))

        cols, seen = [], set()
        for f in self.files:
            for h in f["headers"]:
                k = h.strip().lower()
                if k not in seen:
                    seen.add(k)
                    cols.append(h.strip())
        self.all_cols = cols

        self.cols.delete(0, "end")
        n = len(self.files)
        for i, c in enumerate(cols):
            count = sum(1 for f in self.files if c.lower() in f["map"])
            label = c if n < 2 else "{}    [in {} of {} files]".format(c, count, n)
            self.cols.insert("end", label)
            if c in selected_before or c not in old_cols:
                self.cols.select_set(i)

        total = sum(len(f["rows"]) for f in self.files)
        self.col_count.config(text="{} columns".format(len(cols)) if cols else "")
        self.run_btn.config(state="normal" if self.files else "disabled")
        self.refresh_match_options()
        if self.files:
            self.status_var.set("{} file(s), {} rows loaded. Pick columns, then click Merge to Excel."
                                .format(len(self.files), total))
        else:
            self.status_var.set("Step 1: Add one or more files (.tsv, .csv, .xlsx)")

    # ----- columns / options -----
    def set_all_cols(self, on):
        if on:
            self.cols.select_set(0, "end")
        else:
            self.cols.select_clear(0, "end")
        self.refresh_match_options()

    def selected_cols(self):
        return [self.all_cols[i] for i in self.cols.curselection()]

    def refresh_match_options(self):
        options = ["All selected columns"] + self.selected_cols()
        self.match_box["values"] = options
        if self.match_var.get() not in options:
            self.match_var.set("All selected columns")
        self.toggle_match()

    def toggle_match(self):
        self.match_box.config(state="readonly" if self.dedupe_var.get() else "disabled")

    def pick_out(self):
        current = Path(self.out_var.get())
        path = filedialog.asksaveasfilename(
            title="Save merged file as", defaultextension=".xlsx",
            initialdir=str(current.parent), initialfile=current.name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    # ----- run -----
    def run(self):
        columns = self.selected_cols()
        if not columns:
            messagebox.showwarning("No columns", "Select at least one column to keep.")
            return
        out_path = Path(self.out_var.get()).resolve()
        if out_path.suffix.lower() != ".xlsx":
            out_path = out_path.with_suffix(".xlsx")
        try:
            out_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Bad save location", str(e))
            return

        match = self.match_var.get()
        match_col = None if match == "All selected columns" else match
        dedupe, add_source = self.dedupe_var.get(), self.source_var.get()
        files = list(self.files)

        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self.status_var.set("Merging…")

        def work():
            try:
                headers, data, date_cols, total, dupes = merge(files, columns, dedupe, match_col, add_source)
                if not data:
                    raise RuntimeError("No rows to save — the selected columns are empty in every file.")
                try:
                    write_xlsx(headers, data, date_cols, out_path)
                except PermissionError:
                    raise RuntimeError("Can't save — the file is probably open in Excel.\n"
                                       "Close it (or pick a different name) and try again.")
                if not out_path.exists():
                    raise RuntimeError("The file could not be created at:\n{}".format(out_path))
                self._result = ("ok", (len(data), dupes, total), out_path)
            except Exception as e:
                self._result = ("err", str(e), out_path)

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()
        self.after(200, self.check_done)

    def check_done(self):
        if self._worker.is_alive():
            self.after(200, self.check_done)
            return
        self.progress.stop()
        self.run_btn.config(state="normal")
        kind, value, out_path = self._result
        if kind == "ok":
            kept, dupes, total = value
            msg = "Done! {} rows saved ({} duplicates removed from {} total) to {}".format(
                kept, dupes, total, out_path)
            self.status_var.set(msg)
            if self.open_var.get():
                os.startfile(str(out_path))
            else:
                messagebox.showinfo("Done", msg)
        else:
            self.status_var.set("Something went wrong — see the error message.")
            messagebox.showerror("Couldn't create file", value[-1500:])


if __name__ == "__main__":
    App().mainloop()

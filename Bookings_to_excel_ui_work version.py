"""
Bookings_to_excel_ui.py
Desktop app: pick a Microsoft Bookings export (.tsv), choose columns,
and build a formatted Excel table, then open it in Excel.

NO INSTALLS: Python standard library only. The .xlsx is written directly,
so no PowerShell or Excel automation is needed (nothing for security tools to block).
Run it with the VS Code Run button, or double-click it.
"""

import csv
import os
import re
import threading
import tkinter as tk
import zipfile
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape
from tkinter import filedialog, messagebox, ttk

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]
EXCEL_EPOCH = datetime(1899, 12, 30)


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


# ---------------- Data logic ----------------

def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def read_tsv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter="\t"))
    if not rows:
        raise ValueError("No rows found in this file.")
    return list(rows[0].keys()), rows


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


def build_excel(rows, headers, out_path):
    data = [[(r.get(h) or "").strip() for h in headers] for r in rows]

    date_cols = set()
    for i in range(len(headers)):
        values = [row[i] for row in data if row[i]]
        if values and all(parse_date(v) for v in values):
            date_cols.add(i)
    for row in data:
        for i in date_cols:
            if row[i]:
                row[i] = (parse_date(row[i]) - EXCEL_EPOCH).total_seconds() / 86400

    try:
        write_xlsx(headers, data, date_cols, out_path)
    except PermissionError:
        raise RuntimeError("Can't save — the file is probably open in Excel.\n"
                           "Close it (or pick a different name) and try again.")
    if not Path(out_path).exists():
        raise RuntimeError("The file could not be created at:\n{}".format(out_path))
    return len(data)


# ---------------- UI ----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bookings to Excel")
        self.geometry("560x520")
        self.minsize(480, 460)

        self.rows, self.headers = [], []
        self.tsv_var = tk.StringVar()
        default_out = downloads_folder() / "Bookings_{:%Y%m%d}.xlsx".format(date.today())
        self.out_var = tk.StringVar(value=str(default_out))
        self.open_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Step 1: Choose your Bookings export (.tsv)")
        self._worker = None
        self._result = None

        pad = {"padx": 12, "pady": 6}
        style = ttk.Style(self)
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)

        # Step 1: file
        f1 = ttk.LabelFrame(self, text="1. Bookings export")
        f1.pack(fill="x", **pad)
        ttk.Entry(f1, textvariable=self.tsv_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(f1, text="Upload .tsv…", command=self.pick_tsv).pack(side="right", padx=8)

        # Step 2: columns
        f2 = ttk.LabelFrame(self, text="2. Columns to include")
        f2.pack(fill="both", expand=True, **pad)
        btns = ttk.Frame(f2)
        btns.pack(fill="x", padx=8, pady=(6, 0))
        ttk.Button(btns, text="Select all", command=lambda: self.cols.select_set(0, "end")).pack(side="left")
        ttk.Button(btns, text="Clear", command=lambda: self.cols.select_clear(0, "end")).pack(side="left", padx=6)
        self.count_lbl = ttk.Label(btns, text="")
        self.count_lbl.pack(side="right")

        lf = ttk.Frame(f2)
        lf.pack(fill="both", expand=True, padx=8, pady=8)
        self.cols = tk.Listbox(lf, selectmode="multiple", exportselection=False,
                               font=("Segoe UI", 10), activestyle="none")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.cols.yview)
        self.cols.configure(yscrollcommand=sb.set)
        self.cols.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Step 3: output
        f3 = ttk.LabelFrame(self, text="3. Save Excel file as")
        f3.pack(fill="x", **pad)
        ttk.Entry(f3, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(f3, text="Browse…", command=self.pick_out).pack(side="right", padx=8)

        # Run
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        ttk.Checkbutton(bottom, text="Open in Excel when done", variable=self.open_var).pack(side="left")
        self.run_btn = ttk.Button(bottom, text="Create Excel File", style="Big.TButton",
                                  command=self.run, state="disabled")
        self.run_btn.pack(side="right")

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12)
        ttk.Label(self, textvariable=self.status_var, foreground="#444").pack(anchor="w", padx=12, pady=(4, 10))

        self.after(200, self.pick_tsv)   # prompt for the file right away

    def pick_tsv(self):
        path = filedialog.askopenfilename(
            title="Select Bookings export",
            initialdir=str(downloads_folder()),
            filetypes=[("Bookings export", "*.tsv"), ("Text files", "*.txt *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self.headers, self.rows = read_tsv(path)
        except Exception as e:
            messagebox.showerror("Couldn't read file", str(e))
            return

        self.tsv_var.set(path)
        self.cols.delete(0, "end")
        for h in self.headers:
            self.cols.insert("end", h)
        self.cols.select_set(0, "end")
        self.count_lbl.config(text="{} rows found".format(len(self.rows)))
        self.run_btn.config(state="normal")
        self.status_var.set("Step 2: Pick columns, then click Create Excel File")

    def pick_out(self):
        current = Path(self.out_var.get())
        path = filedialog.asksaveasfilename(
            title="Save Excel file as", defaultextension=".xlsx",
            initialdir=str(current.parent), initialfile=current.name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    def run(self):
        chosen = [self.headers[i] for i in self.cols.curselection()]
        if not chosen:
            messagebox.showwarning("No columns", "Select at least one column.")
            return
        out_path = Path(self.out_var.get()).resolve()
        try:
            out_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Bad save location", str(e))
            return

        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self.status_var.set("Building Excel file…")
        self._result = None

        def work():
            try:
                self._result = ("ok", build_excel(self.rows, chosen, out_path), out_path)
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
            self.status_var.set("Done! Saved {} appointments to {}".format(value, out_path))
            if self.open_var.get():
                os.startfile(str(out_path))
            else:
                messagebox.showinfo("Done", "Saved {} appointments to:\n{}".format(value, out_path))
        else:
            self.status_var.set("Something went wrong — see the error message.")
            messagebox.showerror("Couldn't create file", value[-1500:])


if __name__ == "__main__":
    App().mainloop()

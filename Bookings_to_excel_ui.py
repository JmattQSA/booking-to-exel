"""
Bookings_to_excel_ui.py
Desktop app: pick a Microsoft Bookings export (.tsv), choose columns,
and build a formatted Excel table using the Excel installed on your PC.

NO INSTALLS: Python standard library only (tkinter + PowerShell for Excel).
Run it with the VS Code Run button, or double-click it.
"""

import csv
import json
import os
import subprocess
import tempfile
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]
EXCEL_EPOCH = datetime(1899, 12, 30)
CREATE_NO_WINDOW = 0x08000000   # hides the PowerShell console window

PS_SCRIPT = r"""
param([string]$JsonPath, [string]$OutFile)
$ErrorActionPreference = 'Stop'

$d        = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
$headers  = @($d.headers)
$rows     = @($d.rows)
$dateCols = @($d.date_cols)
$r = $rows.Count + 1
$c = $headers.Count

$arr = New-Object 'object[,]' $r, $c
for ($j = 0; $j -lt $c; $j++) { $arr[0, $j] = $headers[$j] }
for ($i = 0; $i -lt $rows.Count; $i++) {
    $row = @($rows[$i])
    for ($j = 0; $j -lt $c; $j++) { $arr[($i + 1), $j] = $row[$j] }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Appointments'

    for ($j = 0; $j -lt $c; $j++) {
        if ($dateCols -notcontains $j) { $ws.Columns.Item($j + 1).NumberFormat = '@' }
    }

    $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($r, $c))
    $rng.Value2 = $arr

    if ($r -gt 1) {
        foreach ($j in $dateCols) {
            $ws.Range($ws.Cells.Item(2, $j + 1), $ws.Cells.Item($r, $j + 1)).NumberFormat = 'mm/dd/yyyy h:mm AM/PM'
        }
    }

    $lo = $ws.ListObjects.Add(1, $rng, $null, 1)
    $lo.Name = 'Bookings'
    $lo.TableStyle = 'TableStyleMedium2'
    $rng.Font.Name = 'Arial'
    $rng.Font.Size = 10
    $ws.Columns.AutoFit() | Out-Null

    try {
        $win = $wb.Windows.Item(1)
        $win.SplitColumn = 0
        $win.SplitRow = 1
        $win.FreezePanes = $true
    } catch { }

    $wb.SaveAs($OutFile, 51)
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
"""


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


def build_excel(rows, headers, out_path):
    data = [[(r.get(h) or "").strip() for h in headers] for r in rows]

    date_cols = []
    for i in range(len(headers)):
        values = [row[i] for row in data if row[i]]
        if values and all(parse_date(v) for v in values):
            date_cols.append(i)
    for row in data:
        for i in date_cols:
            if row[i]:
                row[i] = (parse_date(row[i]) - EXCEL_EPOCH).total_seconds() / 86400

    tmp = Path(tempfile.mkdtemp())
    json_path, ps_path = tmp / "bookings.json", tmp / "build_excel.ps1"
    try:
        json_path.write_text(json.dumps({"headers": headers, "rows": data, "date_cols": date_cols}),
                             encoding="utf-8")
        ps_path.write_text(PS_SCRIPT, encoding="utf-8-sig")
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_path),
             "-JsonPath", str(json_path), "-OutFile", str(out_path)],
            capture_output=True, text=True, creationflags=CREATE_NO_WINDOW,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout or "Unknown Excel error")
    finally:
        for f in (json_path, ps_path):
            if f.exists():
                f.unlink()
        tmp.rmdir()
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
        default_out = Path.home() / "Documents" / "Bookings_{:%Y%m%d}.xlsx".format(date.today())
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
            initialdir=str(Path.home() / "Downloads"),
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
        self.status_var.set("Building Excel file… (Excel is working in the background)")
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
            messagebox.showerror("Excel step failed", value[-1500:])


if __name__ == "__main__":
    App().mainloop()

"""
bookings_to_excel_nowin32.py
Converts a Microsoft Bookings report export (.tsv) into a formatted Excel table
using the Excel installed on your PC.

NO INSTALLS NEEDED: uses only Python's standard library. Excel is driven through
Windows' built-in PowerShell, so pywin32 is not required.

Usage:
    py bookings_to_excel_nowin32.py "C:/Users/you/Downloads/BookingsReportingData.tsv"
    py bookings_to_excel_nowin32.py report.tsv -o "C:/Reports/appointments.xlsx"
    py bookings_to_excel_nowin32.py report.tsv -c "Customer name" "Service name" "Start time"
    py bookings_to_excel_nowin32.py report.tsv --open
"""

import argparse
import csv
import json
import os
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]
EXCEL_EPOCH = datetime(1899, 12, 30)

# PowerShell that does the Excel work (PowerShell can talk to Excel natively)
PS_SCRIPT = r"""
param([string]$JsonPath, [string]$OutFile)
$ErrorActionPreference = 'Stop'

$d        = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
$headers  = @($d.headers)
$rows     = @($d.rows)
$dateCols = @($d.date_cols)
$r = $rows.Count + 1
$c = $headers.Count

# Build a 2D array so Excel can take all the data in one write
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

    # Text columns stay text (keeps phone numbers / IDs intact)
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

    $lo = $ws.ListObjects.Add(1, $rng, $null, 1)     # Insert > Table
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

    $wb.SaveAs($OutFile, 51)                        # 51 = .xlsx
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
"""


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def to_excel_serial(dt):
    return (dt - EXCEL_EPOCH).total_seconds() / 86400


def read_tsv(path, columns=None):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter="\t"))
    if not rows:
        sys.exit("No rows found in {}".format(path))

    headers = list(rows[0].keys())
    if columns:
        missing = [c for c in columns if c not in headers]
        if missing:
            sys.exit("Column(s) not found: {}\nAvailable: {}".format(missing, headers))
        headers = columns

    data = [[(r.get(h) or "").strip() for h in headers] for r in rows]
    return headers, data


def prepare(headers, data):
    """Find date columns and convert them to Excel date numbers."""
    date_cols = []
    for i in range(len(headers)):
        values = [row[i] for row in data if row[i]]
        if values and all(parse_date(v) for v in values):
            date_cols.append(i)

    for row in data:
        for i in date_cols:
            if row[i]:
                row[i] = to_excel_serial(parse_date(row[i]))
    return date_cols


def run_excel(headers, data, date_cols, out_path):
    tmp = Path(tempfile.mkdtemp())
    json_path = tmp / "bookings.json"
    ps_path = tmp / "build_excel.ps1"
    try:
        json_path.write_text(json.dumps({"headers": headers, "rows": data, "date_cols": date_cols}),
                             encoding="utf-8")
        ps_path.write_text(PS_SCRIPT, encoding="utf-8-sig")

        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(ps_path), "-JsonPath", str(json_path), "-OutFile", str(out_path)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            sys.exit("Excel step failed:\n" + (result.stderr or result.stdout))
    finally:
        for f in (json_path, ps_path):
            if f.exists():
                f.unlink()
        tmp.rmdir()


def main():
    p = argparse.ArgumentParser(description="Convert a Bookings TSV export to Excel (no pywin32).")
    p.add_argument("tsv", help="Path to the Bookings .tsv export")
    p.add_argument("-o", "--output",
                   default=str(Path.home() / "Documents" / "Bookings_{:%Y%m%d}.xlsx".format(date.today())))
    p.add_argument("-c", "--columns", nargs="+", help="Only keep these columns")
    p.add_argument("--open", action="store_true", help="Open the file in Excel when done")
    args = p.parse_args()

    out_path = Path(args.output).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    headers, data = read_tsv(args.tsv, args.columns)
    date_cols = prepare(headers, data)
    run_excel(headers, data, date_cols, out_path)
    print("Saved {} appointments to {}".format(len(data), out_path))

    if args.open:
        os.startfile(str(out_path))


if __name__ == "__main__":
    main()

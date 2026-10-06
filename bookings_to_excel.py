"""
bookings_to_excel_com.py
Converts a Microsoft Bookings report export (.tsv) into a formatted Excel table,
using the Excel installed on your PC (Windows only).

Setup (one time):
    pip install pywin32

Usage:
    python bookings_to_excel_com.py "C:/Users/you/Downloads/BookingsReportingData.tsv"
    python bookings_to_excel_com.py report.tsv -o "C:/Reports/appointments.xlsx"
    python bookings_to_excel_com.py report.tsv -c "Customer name" "Service name" "Start time"
    python bookings_to_excel_com.py report.tsv --open      # open the file when done
"""

import argparse
import csv
import os
import sys
from datetime import date, datetime
from pathlib import Path

import win32com.client as win32

# Excel constants
XL_SRC_RANGE = 1
XL_YES = 1
XL_XLSX = 51

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]
EXCEL_EPOCH = datetime(1899, 12, 30)


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def to_excel_serial(dt):
    """Excel stores dates as days since 12/30/1899; this avoids pywin32 timezone quirks."""
    return (dt - EXCEL_EPOCH).total_seconds() / 86400


def read_tsv(path, columns=None):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter="\t"))
    if not rows:
        sys.exit(f"No rows found in {path}")

    headers = list(rows[0].keys())
    if columns:
        missing = [c for c in columns if c not in headers]
        if missing:
            sys.exit(f"Column(s) not found: {missing}\nAvailable: {headers}")
        headers = columns

    data = [[(r.get(h) or "").strip() for h in headers] for r in rows]
    return headers, data


def find_date_columns(headers, data):
    """A column counts as a date column if every non-blank value parses as a date."""
    date_cols = set()
    for i in range(len(headers)):
        values = [row[i] for row in data if row[i]]
        if values and all(parse_date(v) for v in values):
            date_cols.add(i)
    return date_cols


def write_with_excel(headers, data, out_path):
    date_cols = find_date_columns(headers, data)

    # Convert date columns to Excel serial numbers
    for row in data:
        for i in date_cols:
            if row[i]:
                row[i] = to_excel_serial(parse_date(row[i]))

    n_rows, n_cols = len(data) + 1, len(headers)

    xl = win32.DispatchEx("Excel.Application")   # separate, hidden Excel instance
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Appointments"

        # Text columns as Text so phone numbers/IDs aren't turned into numbers
        for i in range(n_cols):
            if i not in date_cols:
                ws.Columns(i + 1).NumberFormat = "@"

        # Write everything in one shot (much faster than cell by cell)
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols))
        rng.Value = tuple(tuple(r) for r in [headers] + data)

        for i in date_cols:
            ws.Range(ws.Cells(2, i + 1), ws.Cells(n_rows, i + 1)).NumberFormat = "mm/dd/yyyy h:mm AM/PM"

        # Insert > Table
        table = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
        table.Name = "Bookings"
        table.TableStyle = "TableStyleMedium2"

        rng.Font.Name = "Arial"
        rng.Font.Size = 10
        ws.Columns.AutoFit()

        # Freeze header row
        try:
            win = wb.Windows(1)
            win.SplitColumn = 0
            win.SplitRow = 1
            win.FreezePanes = True
        except Exception:
            pass  # cosmetic only

        wb.SaveAs(str(out_path), XL_XLSX)
    finally:
        if wb is not None:
            wb.Close(False)
        xl.Quit()


def main():
    p = argparse.ArgumentParser(description="Convert a Bookings TSV export to Excel using Excel.")
    p.add_argument("tsv", help="Path to the Bookings .tsv export")
    p.add_argument("-o", "--output",
                   default=str(Path.home() / "Documents" / f"Bookings_{date.today():%Y%m%d}.xlsx"))
    p.add_argument("-c", "--columns", nargs="+", help="Only keep these columns")
    p.add_argument("--open", action="store_true", help="Open the file in Excel when done")
    args = p.parse_args()

    out_path = Path(args.output).resolve()     # Excel needs a full path
    out_path.parent.mkdir(parents=True, exist_ok=True)

    headers, data = read_tsv(args.tsv, args.columns)
    write_with_excel(headers, data, out_path)
    print(f"Saved {len(data)} appointments to {out_path}")

    if args.open:
        os.startfile(out_path)


if __name__ == "__main__":
    main()

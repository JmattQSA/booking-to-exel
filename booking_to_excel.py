DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d",
]


def to_value(text):
    """Turn date-looking text into real datetimes so Excel can sort/filter them."""
    text = (text or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return text


def read_tsv(path, columns=None):
    # utf-8-sig strips the invisible BOM some exports start with
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
    return headers, rows


def write_excel(headers, rows, out_path):
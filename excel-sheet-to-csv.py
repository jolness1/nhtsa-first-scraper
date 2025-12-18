from pathlib import Path
import csv
import re
from typing import Optional

import openpyxl

WORKDIR = Path(__file__).parent
SCRAPED = WORKDIR / "scraped"
OUTDIR = WORKDIR / "output"
OUTDIR.mkdir(exist_ok=True)

MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

NUM_RE = re.compile(r"^-?\(?\$?([0-9,]+(?:\.[0-9]+)?)\)?$")


def parse_number(v) -> Optional[str]:
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        # preserve integer-like as int
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if not s:
        return ""
    # remove commas and currency
    s2 = s.replace(',', '').replace('$', '')
    # handle parentheses as negative
    if s2.startswith('(') and s2.endswith(')'):
        s2 = '-' + s2[1:-1]
    # try numeric conversion
    try:
        if '.' in s2:
            f = float(s2)
            if f.is_integer():
                return str(int(f))
            return str(f)
        return str(int(s2))
    except Exception:
        # fallback: extract digits
        m = NUM_RE.match(s)
        if m:
            return m.group(1)
        return s2


def find_month_header_row(rows):
    """Return index of row that contains month names (returns -1 if not found)."""
    for i, row in enumerate(rows):
        cells = [ (c or "").strip() for c in row ]
        found = 0
        for m in MONTHS:
            if any(cell.lower() == m.lower() for cell in cells):
                found += 1
        # choose row with at least 6 month names present
        if found >= 6:
            return i
    return -1


def process_file(path: Path):
    print(f"Processing {path}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header_idx = find_month_header_row(rows)
    if header_idx == -1:
        print(f"  Could not locate month header in {path}; skipping")
        return

    header_row = [ (c or "") for c in rows[header_idx] ]
    # map month name to col index
    month_cols = {}
    total_col = None
    for idx, cell in enumerate(header_row):
        if not cell:
            continue
        cell_s = str(cell).strip()
        for m in MONTHS:
            if cell_s.lower() == m.lower():
                month_cols[m] = idx
        if cell_s.lower() == 'total':
            total_col = idx

    # fallback: if total_col not found, assume it's the last non-empty column
    if total_col is None:
        for idx in range(len(header_row)-1, -1, -1):
            if header_row[idx] not in (None, ""):
                # if this column is not one of the months, consider as total
                if all(header_row[idx].strip().lower() != m.lower() for m in MONTHS):
                    total_col = idx
                    break

    out_rows = []
    # data starts after header_idx
    for r in rows[header_idx+1:]:
        first = r[0]
        if first is None:
            # stop on blank leading cell
            break
        sfirst = str(first).strip()
        if not sfirst:
            break
        # skip summary 'Total' row
        if sfirst.lower().startswith('total'):
            break
        # expect year
        try:
            year = int(float(sfirst))
        except Exception:
            # not a year row
            continue

        row_values = [str(year)]
        for m in MONTHS:
            col = month_cols.get(m)
            v = r[col] if (col is not None and col < len(r)) else None
            row_values.append(parse_number(v))
        # total value
        tot = ''
        if total_col is not None and total_col < len(r):
            tot = parse_number(r[total_col])
        row_values.append(tot)
        out_rows.append(row_values)

    if not out_rows:
        print(f"  No data rows found in {path}")
        return

    outpath = OUTDIR / (path.stem + '.csv')
    with outpath.open('w', newline='', encoding='utf-8') as fh:
        writer = csv.writer(fh)
        writer.writerow(['year'] + MONTHS + ['Total'])
        for orow in out_rows:
            writer.writerow(orow)
    print(f"  Wrote {outpath} ({len(out_rows)} rows)")


if __name__ == '__main__':
    files = sorted(SCRAPED.glob('*.xlsx'))
    if not files:
        print('No .xlsx files found in ./scraped')
    for f in files:
        process_file(f)

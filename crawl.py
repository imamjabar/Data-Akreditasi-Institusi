"""Crawl BAN-PT institution accreditation directory into CSV and Excel.

The directory page renders a client-side DataTable: every row is delivered by a
single JSON request and the 40 pages (100 entries each) are paginated in the
browser. One request therefore captures the full table.
"""

import argparse
import csv
import json
import random
import time
import urllib.request
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

PAGE_URL = "https://www.banpt.or.id/direktori/institusi/pencarian_institusi.php"
DATA_URL = "https://www.banpt.or.id/direktori/model/dir_aipt/get_data_institusi.php"
SOURCE_COLUMNS = [
    "Perguruan Tinggi",
    "Peringkat",
    "No. SK",
    "Tahun SK",
    "Wilayah",
    "Tanggal Kedaluwarsa",
    "Status Kedaluwarsa",
]
COLUMNS = ["No", *SOURCE_COLUMNS]
PAGE_SIZE = 100


def fetch_rows() -> list[list]:
    time.sleep(random.uniform(2, 5))
    req = urllib.request.Request(
        DATA_URL,
        headers={
            "User-Agent": "Mozilla/5.0 (research crawler; polite single request)",
            "Referer": PAGE_URL,
            "X-Requested-With": "XMLHttpRequest",
        },
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        payload = json.load(resp)
    rows = payload["data"]
    bad = [r for r in rows if len(r) != len(SOURCE_COLUMNS)]
    if not rows or bad:
        raise RuntimeError(f"unexpected payload: {len(rows)} rows, {len(bad)} malformed")
    return [[i, *("" if c is None else str(c).strip() for c in r)] for i, r in enumerate(rows, 1)]


def write_csv(path: str, rows: list[list]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
        w.writerows(rows)


def write_xlsx(path: str, rows: list[list], snapshot: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Institusi"
    ws.append(COLUMNS)
    tahun = COLUMNS.index("Tahun SK")
    for r in rows:
        ws.append([int(c) if j == tahun and c.isdigit() else c for j, c in enumerate(r)])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(COLUMNS):
        width = max(len(col), *(len(str(r[i])) for r in rows))
        ws.column_dimensions[get_column_letter(i + 1)].width = min(width + 2, 80)

    meta = wb.create_sheet("Info")
    for k, v in (
        ("Sumber", PAGE_URL),
        ("Tanggal data", snapshot),
        ("Jumlah institusi", len(rows)),
    ):
        meta.append([k, v])
    wb.save(path)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=date.today().isoformat(), help="snapshot date YYYY-MM-DD")
    args = ap.parse_args()

    rows = fetch_rows()
    base = f"data_akreditasi_institusi_{args.date}"
    write_csv(f"{base}.csv", rows)
    write_xlsx(f"{base}.xlsx", rows, args.date)
    pages = -(-len(rows) // PAGE_SIZE)
    print(f"{len(rows)} rows ({pages} pages x {PAGE_SIZE}) -> {base}.csv, {base}.xlsx")


if __name__ == "__main__":
    main()

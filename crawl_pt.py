"""Crawl the BAN-PT perguruan tinggi list via the bianglala autocomplete endpoint.

Queries every two-letter term aa..zz, one request per term with a random 2-5 s
delay. Raw responses are cached in .cache/pt_terms.jsonl so an interrupted run
resumes without re-requesting finished terms.
"""

import argparse
import itertools
import json
import os
import random
import re
import string
import time
import urllib.parse
import urllib.request
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font

from crawl import write_csv

SEARCH_URL = "https://service.banpt.or.id/bianglala/searchPT.php"
CACHE = ".cache/pt_terms.jsonl"
HEADER = ["Kode PT", "Nama PT"]
TERMS = ["".join(p) for p in itertools.product(string.ascii_lowercase, repeat=2)]
ITEM = re.compile(r"^(\S*)\s*-\s+(.+)$")
NO_CODE = {"", "-"}


def load_cache() -> dict[str, list[str]]:
    if not os.path.exists(CACHE):
        return {}
    with open(CACHE, encoding="utf-8") as f:
        return {e["term"]: e["results"] for e in map(json.loads, f)}


def fetch_term(term: str) -> list[str] | None:
    """Return results, or None when the server answers 200 with an empty body."""
    url = f"{SEARCH_URL}?{urllib.parse.urlencode({'term': term})}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (research crawler; polite)"})
    empty = 0
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                body = resp.read()
            if not body.strip():
                empty += 1
                if empty == 2:
                    return None
                time.sleep(random.uniform(2, 5))
                continue
            results = json.loads(body)
            if not isinstance(results, list):
                raise ValueError(f"unexpected payload type {type(results).__name__}")
            return results
        except Exception as e:
            if attempt == 2:
                raise RuntimeError(f"term {term!r} failed: {e}") from e
            time.sleep(30 * (attempt + 1))
    raise RuntimeError(f"term {term!r} failed")


def crawl() -> dict[str, list[str] | None]:
    cache = load_cache()
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    todo = [t for t in TERMS if t not in cache]
    print(f"{len(cache)} cached, {len(todo)} to fetch", flush=True)
    with open(CACHE, "a", encoding="utf-8") as f:
        for n, term in enumerate(todo, 1):
            time.sleep(random.uniform(2, 5))
            results = fetch_term(term)
            cache[term] = results
            f.write(json.dumps({"term": term, "results": results}, ensure_ascii=False) + "\n")
            f.flush()
            print(f"[{n}/{len(todo)}] {term}: {'EMPTY BODY' if results is None else len(results)}", flush=True)
    return cache


def parse(cache: dict[str, list[str] | None]) -> list[list[str]]:
    """Distinct (Kode PT, Nama PT) pairs; missing codes ("- X", "- - X") become ""."""
    pairs, bad = set(), set()
    for results in cache.values():
        for item in results or ():
            item = " ".join(item.split())
            if not item:
                continue
            m = ITEM.match(item)
            if not m:
                bad.add(item)
                continue
            kode = "" if m[1] in NO_CODE else m[1]
            pairs.add((kode, m[2]))
    if bad:
        raise RuntimeError(f"{len(bad)} unparseable entries, e.g. {sorted(bad)[:5]}")
    return [list(p) for p in sorted(pairs, key=lambda p: (p[0] == "", p[0], p[1].casefold()))]


def write_xlsx(path: str, rows: list[list[str]], snapshot: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Perguruan Tinggi"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for (cell,) in ws.iter_rows(min_row=2, max_col=1):
        cell.number_format = "@"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = min(max(len(r[1]) for r in rows) + 2, 100)

    meta = wb.create_sheet("Info")
    for k, v in (
        ("Sumber", f"{SEARCH_URL}?term=aa .. ?term=zz"),
        ("Tanggal data", snapshot),
        ("Jumlah baris", len(rows)),
    ):
        meta.append([k, v])
    wb.save(path)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=date.today().isoformat(), help="snapshot date YYYY-MM-DD")
    args = ap.parse_args()

    cache = crawl()
    missing = [t for t in TERMS if t not in cache]
    if missing:
        raise RuntimeError(f"{len(missing)} terms missing: {missing[:10]}")
    rows = parse(cache)
    base = f"data_perguruan_tinggi_{args.date}"
    write_csv(f"{base}.csv", HEADER, rows)
    write_xlsx(f"{base}.xlsx", rows, args.date)
    empty = [t for t in TERMS if cache[t] is None]
    print(f"{len(TERMS)} terms ({len(empty)} empty-body: {' '.join(empty) or '-'}) "
          f"-> {len(rows)} unique rows -> {base}.csv, {base}.xlsx")


if __name__ == "__main__":
    main()

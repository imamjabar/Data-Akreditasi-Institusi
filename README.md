# Data Akreditasi Institusi

Accreditation data for Indonesian higher-education institutions (institusi) and
study programs (program studi), from the BAN-PT directory, plus the list of
perguruan tinggi codes and names:

- Institusi: <https://www.banpt.or.id/direktori/institusi/pencarian_institusi.php>
- Program Studi: <https://www.banpt.or.id/direktori/prodi/pencarian_prodi.php>
- Perguruan Tinggi: <https://service.banpt.or.id/bianglala/searchPT.php?term=aa> … `?term=zz`

**Tanggal data: 15 Oktober 2026**

## Files

| File | Rows | Source |
|---|---|---|
| `data_akreditasi_institusi_2026-10-15.csv` / `.xlsx` | 3,926 | 40 pages × 100 entries |
| `data_akreditasi_prodi_2026-10-15.csv` / `.xlsx` | 36,734 | 368 pages × 100 entries |
| `data_perguruan_tinggi_2026-10-15.csv` / `.xlsx` | 11,054 | 676 searches, `aa` … `zz` |
| `crawl.py` | | Crawler for institusi and prodi |
| `crawl_pt.py` | | Crawler for the perguruan tinggi list |

- CSV: UTF-8 with BOM, opens directly in Excel.
- Excel: one data sheet (`Institusi`, `Prodi` or `Perguruan Tinggi`) and one `Info` sheet (source, date, row count).

## Columns

### data_akreditasi_institusi

| Column | Description |
|---|---|
| No | Row number |
| Perguruan Tinggi | Institution name |
| Peringkat | Accreditation rank (Unggul, Baik Sekali, Baik, Terakreditasi, A, B, C, ...) |
| No. SK | Accreditation decree number |
| Tahun SK | Decree year |
| Wilayah | Region code (01–17), or institution group (PTN, PTAS, PTKL, ...) |
| Tanggal Kedaluwarsa | Expiry date (YYYY-MM-DD); empty when the source has no date |
| Status Kedaluwarsa | Expiry status on the snapshot date |

### data_akreditasi_prodi

| Column | Description |
|---|---|
| No | Row number |
| Perguruan Tinggi | Institution name |
| Program Studi | Study program name |
| Strata | Level (S1, S2, S3, D-III, D-IV, Profesi, Spesialis, ...) |
| Wilayah | Region code (01–17), or institution group (PTN, PTAS, PTAN, PTKL, ...) |
| No. SK | Accreditation decree number |
| Tahun SK | Decree year; empty for 1 row in the source |
| Peringkat | Accreditation rank, as written in the source (for example `Unggul` and `UNGGUL` both occur) |
| Tanggal Kedaluwarsa | Expiry date (YYYY-MM-DD); `-` or empty when the source has no date |
| Status Kedaluwarsa | Expiry status on the snapshot date |

### data_perguruan_tinggi

| Column | Description |
|---|---|
| Kode PT | Institution code, stored as text (for example `001001`); empty when the source has no code |
| Nama PT | Institution name |

Each source entry has the form `094143 - Akademi Keperawatan Mappa Oudang Makassar`. It is
split into `Kode PT` = `094143` and `Nama PT` = `Akademi Keperawatan Mappa Oudang Makassar`.

## Notes

- Values are kept as published by BAN-PT. Only leading and trailing spaces are removed.
- `No` follows the row order of the BAN-PT data feed. The website sorts the table again
  for display (institusi by name, prodi by Tahun SK, newest first).
- `Status Kedaluwarsa` values such as "74 hari lagi kedaluwarsa" count from the snapshot date.
- The prodi source contains 1,090 rows that are exact duplicates of other rows. They are
  kept, so the files match the source. Remove them before you count programs.

### Perguruan tinggi list

- Rows are distinct (Kode PT, Nama PT) pairs, sorted by Kode PT. Repeated spaces are reduced to one
  space before duplicates are removed.
- One code can have several names, for example `001001 - Universitas Gadjah Mada` and
  `001001 - Universitas Gadjah Mada, Yogyakarta`. Both rows are kept, because they are
  different entries in the source. There are 6,321 unique codes.
- 124 rows have no code: the source shows `- Nama` or `- - Nama`.
- 11 codes have only 4 or 5 digits, as in the source (a leading zero is probably missing, for
  example `64010` and `064010`). They are not changed.
- Excel removes leading zeros when it opens the CSV directly. Use the `.xlsx` file, or import
  the CSV with `Kode PT` as text.
- **Coverage gap:** for 56 terms the server returns HTTP 200 with an empty body every time:
  `ah al ar as ba be bi ek en er fh gg gi gk gl ik il im in io is it iv ka kn ko la li lm lo ma
  mf mi mu na ng ni no nt ob ol om on po ri rn rs se si sy ta te ti un ve ya`. This is a
  server-side error for short terms with many matches, not rate limiting: other terms worked
  between them, and full-name searches for the same institutions work. Most institutions are
  still found through other terms. An institution whose name has only these letter pairs (for
  example "Universitas Nasional") cannot be found with any two-letter term, so the list is not
  complete.

## Method

Each directory page is a client-side DataTable. The browser loads all rows with
one JSON request, then splits them into pages. The crawler sends that single
request per dataset, after a random 2–5 s delay. This collects all pages with one
request each and does not flood the server.

The perguruan tinggi list comes from the bianglala autocomplete endpoint. `crawl_pt.py`
sends one request for each two-letter term `aa` … `zz` (676 requests), with a random 2–5 s
delay between requests. Responses are cached in `.cache/pt_terms.jsonl`, so an interrupted
run continues where it stopped.

## Reproduce

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python crawl.py institusi --date 2026-10-15
.venv/bin/python crawl.py prodi --date 2026-10-15
.venv/bin/python crawl_pt.py --date 2026-10-15
```

## Source and license

Data source: Badan Akreditasi Nasional Perguruan Tinggi (BAN-PT). Before you
use or publish this data, check BAN-PT's terms of use.

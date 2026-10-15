# Data Akreditasi Institusi

Accreditation data for Indonesian higher-education institutions (institusi) and
study programs (program studi), from the BAN-PT directory:

- Institusi: <https://www.banpt.or.id/direktori/institusi/pencarian_institusi.php>
- Program Studi: <https://www.banpt.or.id/direktori/prodi/pencarian_prodi.php>

**Tanggal data: 15 Oktober 2026**

## Files

| File | Rows | Source pages (100 entries) |
|---|---|---|
| `data_akreditasi_institusi_2026-10-15.csv` / `.xlsx` | 3,926 | 40 |
| `data_akreditasi_prodi_2026-10-15.csv` / `.xlsx` | 36,734 | 368 |
| `crawl.py` | Crawler script used to collect the data | |

- CSV: UTF-8 with BOM, opens directly in Excel.
- Excel: one data sheet (`Institusi` or `Prodi`) and one `Info` sheet (source, date, row count).

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

## Notes

- Values are kept as published by BAN-PT. Only leading and trailing spaces are removed.
- `No` follows the row order of the BAN-PT data feed. The website sorts the table again
  for display (institusi by name, prodi by Tahun SK, newest first).
- `Status Kedaluwarsa` values such as "74 hari lagi kedaluwarsa" count from the snapshot date.
- The prodi source contains 1,090 rows that are exact duplicates of other rows. They are
  kept, so the files match the source. Remove them before you count programs.

## Method

Each directory page is a client-side DataTable. The browser loads all rows with
one JSON request, then splits them into pages. The crawler sends that single
request per dataset, after a random 2–5 s delay. This collects all pages with one
request each and does not flood the server.

## Reproduce

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python crawl.py institusi --date 2026-10-15
.venv/bin/python crawl.py prodi --date 2026-10-15
```

## Source and license

Data source: Badan Akreditasi Nasional Perguruan Tinggi (BAN-PT). Before you
use or publish this data, check BAN-PT's terms of use.

# Data Akreditasi Institusi

Data akreditasi institusi perguruan tinggi di Indonesia dari direktori BAN-PT:
<https://www.banpt.or.id/direktori/institusi/pencarian_institusi.php>

**Tanggal data: 15 Oktober 2026**

## File

| File | Format |
|---|---|
| `data_akreditasi_institusi_2026-10-15.csv` | CSV (UTF-8 with BOM, opens directly in Excel) |
| `data_akreditasi_institusi_2026-10-15.xlsx` | Excel; sheet `Institusi` (data) and `Info` (source, date, row count) |
| `crawl.py` | Crawler script used to collect the data |

Total: **3,926 institutions** (40 pages × 100 entries on the source site).

## Columns

| Column | Description |
|---|---|
| No | Row number (1–3926), same order as the source table |
| Perguruan Tinggi | Institution name |
| Peringkat | Accreditation rank (Unggul, Baik Sekali, Baik, Terakreditasi, A, B, C, ...) |
| No. SK | Accreditation decree number |
| Tahun SK | Decree year |
| Wilayah | Region code (01–17), or institution group (PTN, PTAS, PTKL, ...) |
| Tanggal Kedaluwarsa | Expiry date (YYYY-MM-DD); empty when the source has no date |
| Status Kedaluwarsa | Expiry status, as shown on the source site on the snapshot date |

`Status Kedaluwarsa` values such as "74 hari lagi kedaluwarsa" are relative to the snapshot date.

## Method

The source page is a client-side DataTable. The browser loads all rows with one
JSON request (`/direktori/model/dir_aipt/get_data_institusi.php`), then shows
them in pages. The crawler sends that single request after a random 2–5 s delay,
so it collects all 40 pages with one request and does not flood the server.

## Reproduce

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python crawl.py --date 2026-10-15
```

## Source and license

Data source: Badan Akreditasi Nasional Perguruan Tinggi (BAN-PT). Before you
use or publish this data, check BAN-PT's terms of use.

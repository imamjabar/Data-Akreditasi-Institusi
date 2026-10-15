# Data Akreditasi Institusi

Accreditation data for Indonesian higher-education institutions (institusi) and
study programs (program studi), from the BAN-PT directory, plus the list of
perguruan tinggi codes and names, and the Kode PT for every institution in the two datasets:

- Institusi: <https://www.banpt.or.id/direktori/institusi/pencarian_institusi.php>
- Program Studi: <https://www.banpt.or.id/direktori/prodi/pencarian_prodi.php>
- Perguruan Tinggi: <https://service.banpt.or.id/bianglala/searchPT.php?term=aa> … `?term=zz`, plus name searches

**Tanggal data: 15 Oktober 2026**

## Files

| File | Rows | Source |
|---|---|---|
| `data_akreditasi_institusi_2026-10-15.csv` / `.xlsx` | 3,926 | 40 pages × 100 entries |
| `data_akreditasi_prodi_2026-10-15.csv` / `.xlsx` | 36,734 | 368 pages × 100 entries |
| `data_perguruan_tinggi_2026-10-15.csv` / `.xlsx` | 11,103 | 676 searches `aa` … `zz`, plus 151 name searches |
| `data_kode_pt_2026-10-15.csv` / `.xlsx` | 6,968 | Kode PT for each (Perguruan Tinggi, Wilayah) in the two datasets |
| `kode_pt_manual.csv` | 12 | Matches reviewed by hand (input for `crawl_pt.py`) |
| `crawl.py` | | Crawler for institusi and prodi |
| `crawl_pt.py` | | Crawler for the perguruan tinggi list and the Kode PT mapping |

- CSV: UTF-8 with BOM, opens directly in Excel.
- Excel: one data sheet (`Institusi`, `Prodi`, `Perguruan Tinggi` or `Kode PT`) and one `Info` sheet (source, date, row count).

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
| Original | The source entry exactly as BAN-PT returns it, not parsed (outer spaces removed) |

Each source entry has the form `094143 - Akademi Keperawatan Mappa Oudang Makassar`. It is
split into `Kode PT` = `094143` and `Nama PT` = `Akademi Keperawatan Mappa Oudang Makassar`.

### data_kode_pt

One row for each distinct (Perguruan Tinggi, Wilayah) pair in `data_akreditasi_institusi` and
`data_akreditasi_prodi`. Join on both columns to add Kode PT to those files.

| Column | Description |
|---|---|
| Perguruan Tinggi | Name exactly as in the accreditation datasets |
| Wilayah | Wilayah exactly as in the accreditation datasets |
| Kode PT | The selected code; empty when the code cannot be decided (see `Kandidat Kode PT`) |
| Kandidat Kode PT | All codes found for this name, separated by `; ` |
| Nama PT (BAN-PT) | The matching name(s) in the perguruan tinggi list |
| Metode | `exact`: same name after normalisation; `fuzzy 0.xx`: similar name, checked by hand; `manual`: from `kode_pt_manual.csv` |
| Dasar Kode | How `Kode PT` was selected: `satu kandidat`, `jenis PT`, `wilayah`, `manual` |
| Catatan | Note for manual and unresolved rows |
| Sumber | `institusi`, `prodi`, or both |

## Notes

- Values are kept as published by BAN-PT. Only leading and trailing spaces are removed.
- `No` follows the row order of the BAN-PT data feed. The website sorts the table again
  for display (institusi by name, prodi by Tahun SK, newest first).
- `Status Kedaluwarsa` values such as "74 hari lagi kedaluwarsa" count from the snapshot date.
- The prodi source contains 1,090 rows that are exact duplicates of other rows. They are
  kept, so the files match the source. Remove them before you count programs.

### Perguruan tinggi list

- Rows are distinct `Original` entries, sorted by Kode PT. One code can have several names, for
  example `001001 - Universitas Gadjah Mada` and `001001 - Universitas Gadjah Mada, Yogyakarta`.
  There are 6,326 unique codes.
- In `Nama PT`, repeated spaces are reduced to one space. `Original` keeps them.
- 124 rows have no code: the source shows `- Nama` or `- - Nama`.
- 11 codes have only 4 or 5 digits, as in the source (a leading zero is probably missing, for
  example `64010` and `064010`). They are not changed.
- Excel removes leading zeros when it opens the CSV directly. Use the `.xlsx` file, or import
  the CSV with `Kode PT` as text.
- For 56 two-letter terms the server returns HTTP 200 with an empty body every time:
  `ah al ar as ba be bi ek en er fh gg gi gk gl ik il im in io is it iv ka kn ko la li lm lo ma
  mf mi mu na ng ni no nt ob ol om on po ri rn rs se si sy ta te ti un ve ya`. This is a
  server-side error for short terms with many matches, not rate limiting. To close this gap,
  every dataset name that was not in the list was also searched by name (see Method).
- Terms that contain an apostrophe (`'`) cause HTTP 500 on the server, so search terms never
  contain one.

### Kode PT mapping

| Result | (Perguruan Tinggi, Wilayah) pairs | Institusi rows | Prodi rows |
|---|---|---|---|
| One Kode PT selected | 6,921 of 6,968 | 3,896 of 3,926 | 36,600 of 36,734 |
| At least one candidate code | 6,966 of 6,968 | 3,924 of 3,926 | 36,734 of 36,734 |

- Names are compared after normalisation: HTML entities (`&#039;`), broken characters (`Â`),
  parentheses, punctuation, case, typos (`Univeritas`, `llmu`, `Insitut`, `Politenik`) and
  abbreviations (STIKES, STAI, STIT, STKIP, STIE, STMIK, IAIN, UIN, AKPER, Akbid, Poltekkes
  Kemenkes, AMIK, ...) are normalised. A trailing `, City` is ignored.
- When a name has several codes, the code is selected by:
  - `jenis PT`: the third digit of Kode PT is the institution type (1 Universitas, 2 Institut,
    3 Sekolah Tinggi, 4 Akademi, 5 Politeknik, 6 Akademi Komunitas). This holds for 98.6% of
    the list.
  - `wilayah`: the first two digits match the dataset Wilayah (LLDIKTI number; `00` for PTN;
    `2` + even digit for PTAN; `2` + odd digit for PTAS). Example: Universitas Tamansiswa is
    `021011` in Wilayah 02 and `101005` in Wilayah 10.
- 45 pairs keep several candidates and an empty `Kode PT`, because the candidates have the same
  type and region (for example `Politeknik Aceh`: `135001; 135003`; and `Sekolah Tinggi Agama
  Islam Miftahul Ulum`, which has no city, while BAN-PT lists six of them).
- The 35 fuzzy matches and 12 manual decisions were checked by hand. `Akademi Kebidanan
  Palapa Husada` has only an unconfirmed candidate (`014198`, Paluta Husada).
- Two institutions are not in the BAN-PT perguruan tinggi list under any searched name, so they
  have no code: `STKIP Asy-Syafi iyah Internasional Medan` and `Sekolah Tinggi Ilmu Tarbiyah
  Bolaang mongondow Kotamobagu Sulawesi Utara`.

## Method

Each directory page is a client-side DataTable. The browser loads all rows with
one JSON request, then splits them into pages. The crawler sends that single
request per dataset, after a random 2–5 s delay. This collects all pages with one
request each and does not flood the server.

The perguruan tinggi list comes from the bianglala autocomplete endpoint. `crawl_pt.py`
sends one request for each two-letter term `aa` … `zz` (676 requests). Then it searches each
dataset name that is not in the list yet: first the cleaned full name, then its first two
distinctive words, then its longest distinctive word. It stops at the first exact match.
Every request waits a random 2–5 s. Responses are cached in `.cache/`, so an interrupted run
continues where it stopped. All results are added to the list.

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

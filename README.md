# IQC Barcode Scanner

A single-file, configurable **Incoming Quality Control (IQC) barcode validation and traceability tool** for electronics / PCBA manufacturing, built with Python and Tkinter.

Scan a barcode, get an instant, full-screen **PASS / FAIL**, and have every scan recorded in SQL Server with who scanned it, when, where and why it failed.

![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)
![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)
![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)

<!-- Add screenshots (use dummy data only) to docs/ and uncomment:
![Main window](docs/main-window.png)
![Panel scan](docs/panel-scan.png)
-->

## Features

- **Configurable validation rules** per project: any barcode, prefix/suffix, fixed length, regular expression, or *structured* (prefix + year + work-week + running number, with configurable positions)
- **Large PASS / FAIL feedback** and a live scan feed, designed for use on a production floor with a handheld scanner (keyboard-wedge)
- **Panel + array mode**: scan a panel serial number, then scan each array on the panel (visual grid up to 16 arrays, sequential mode above that)
- **Duplicate detection**: per project or global, and within a single panel
- **Full traceability**: project, GR ID, barcode, panel SN, array index, timestamp, operator, station, result, failure reason
- **Records viewer**: filter by barcode / GR ID / operator / project / result / date range, paginated, with CSV export
- **Excel-driven configuration** (`config.xlsx`): no code changes to add a new PCB project
- Session pass/fail counters and all-time yield per project

## Requirements

| Item | Notes |
|---|---|
| Python | 3.9 or newer (with Tkinter) |
| Python packages | `pyodbc`, `openpyxl` (see `requirements.txt`) |
| Database | Microsoft SQL Server (the free **Express** edition is fine) |
| ODBC driver | [Microsoft ODBC Driver 17 or 18 for SQL Server](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server) |
| OS | Developed and used on Windows. Other OSes may run it but are untested. |

## Quick start

```bash
# 1. Get the code
git clone https://github.com/dineshnaidu10/iqc-barcode-scanner.git
cd iqc-barcode-scanner

# 2. (Recommended) virtual environment
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # macOS / Linux

# 3. Install dependencies
pip install -r requirements.txt
```

**4. Create the database.** Open `schema.sql` in SQL Server Management Studio (or Azure Data Studio) and run it. It creates the `IQCScanner` database and the `scan_records` table. *(Optional: run `sample_data.sql` to load a few dummy rows.)*

**5. Configure the connection.** Copy `db_config.ini.example` to `db_config.ini` and edit it:

```ini
[database]
driver   = ODBC Driver 17 for SQL Server
server   = localhost\SQLEXPRESS
database = IQCScanner
trusted  = yes        ; yes = Windows auth, no = SQL auth
username =
password =
```

> `db_config.ini` is git-ignored. Never commit credentials.
>
> If you use ODBC Driver 18 you may also need encryption settings; set `driver = ODBC Driver 18 for SQL Server` and see [Troubleshooting](#troubleshooting).

**6. Configure your projects.** Copy `config.sample.xlsx` to `config.xlsx` (same folder as `iqc_scanner.py`). It contains four demo projects. You can also start the app and use **Config → Edit Configuration…** to generate a blank template.

**7. Run it.**

```bash
python iqc_scanner.py
```

Try the dummy barcodes in [`samples/sample_barcodes.txt`](samples/sample_barcodes.txt).

## Using the app

1. Enter **Operator** and **GR ID** (goods-receipt / lot reference) in the top bar.
2. Choose a **Project**.
3. Scan or type a barcode and press **Enter** (a scanner trigger sends Enter automatically).
4. The result panel turns green (PASS) or red (FAIL) and the scan is saved.

In **panel mode**, the first scan is the panel serial number. If it is valid, a dialog opens where you scan each array. The panel is saved only when **all** arrays have passed and you press *Submit Panel*. Failed array scans are recorded immediately.

| Key | Action |
|---|---|
| `Enter` | Submit scan |
| `F5` | Open Records & Search |
| `F2` | Focus the scan field |
| `Esc` | Clear the scan field |

## Configuration reference (`config.xlsx`)

The app reads columns **by position**, so keep the header row and column order unchanged.

### Sheet `App` (row 2)

| Col | Field | Description |
|---|---|---|
| A | `station` | Station name stored with every record |
| B | `operator_required` | `yes` / `no` |
| C | `scan_timeout_seconds` | Clears an untouched scan field after this many seconds |

### Sheet `Projects` (one row per PCB, from row 2)

| Cols | Field | Description |
|---|---|---|
| A | `project_id` | Unique ID (stored in the database) |
| B | `project_name` | Name shown in the dropdown |
| C | `mode` | `single` or `panel` |
| D | `duplicate_check` | `yes` / `no`: reject barcodes that already PASSED |
| E | `duplicate_scope` | `project` or `global` |
| F | `format` | `none`, `prefix_suffix`, `fixed_length`, `regex`, `structured` |
| G, H | `prefix`, `suffix` | Required start / end text |
| I | `total_length` | Exact barcode length (enforced as min = max) |
| J | `running_number_length` | Digits of the running number (taken from the end, before any suffix) |
| K, L | `year_section_start`, `_length` | 1-based position of the **YY** section; must equal the current year |
| M, N | `workweek_section_start`, `_length` | 1-based position of the **WW** section; must be 01-53 |
| O | `regex_pattern` | Used when `format = regex` (full match) |
| P | `description` | Human-readable rule shown in the UI and in failure messages |
| Q-T | `array_count`, `array_layout`, `array_duplicate_check`, `array_duplicate_scope` | Panel mode only |
| U-AE | `array_*` | Same fields as F-P, applied to each **array** serial number |

In panel mode, columns F-P describe the **panel** serial number and U-AE the **array** serial numbers.

### Example: structured barcode

`MA2641000001` → `MA` (prefix) + `26` (year, pos 3) + `41` (week, pos 5) + `000001` (6-digit running number) = 12 characters.

| prefix | total_length | running_number_length | year start/len | week start/len |
|---|---|---|---|---|
| `MA` | 12 | 6 | 3 / 2 | 5 / 2 |

> **Heads-up:** the year check compares against the *current* year, so barcodes from last year are rejected. The sample barcodes use `26`; adjust them if you test in another year.

## Database

A single table, `scan_records` (see [`schema.sql`](schema.sql)):

`id, project_id, project_name, gr_id, barcode, panel_sn, array_index, timestamp, operator, station, result, failure_reason`

Failed scans are stored too (with `failure_reason`); duplicate detection only considers previous **PASS** records. The column order matters because the Records viewer reads rows by position.

## Packaging as an executable (optional)

The project ships as a single `.py` file. If you want a standalone `.exe`, you can build one yourself with [PyInstaller](https://pyinstaller.org):

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --name IQCScanner ^
  --add-data "db_config.ini;." ^
  --add-data "config.xlsx;." ^
  iqc_scanner.py
```

Notes:

- The app looks for `db_config.ini` and `config.xlsx` next to the script, or inside the PyInstaller bundle when frozen, so bundle them as above.
- Prefer the default one-folder build (no `--onefile`): the config files stay on disk and are easy to edit. With `--onefile` they are unpacked to a temp folder on each launch, so changes require a rebuild.
- An optional window icon is loaded from `barcode_verifier.ico` if present (add `--icon barcode_verifier.ico --add-data "barcode_verifier.ico;."`).
- Do **not** publish or share a build that has your real `db_config.ini` bundled inside.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `db_config.ini not found` on start | Copy `db_config.ini.example` to `db_config.ini` next to the script |
| `Cannot connect to SQL Server` | Check `server`, that SQL Server accepts TCP/IP connections, and your authentication mode |
| `Data source name not found` | The ODBC driver named in `driver` is not installed; check with `pyodbc.drivers()` |
| SSL / certificate error with Driver 18 | Add `TrustServerCertificate=yes;` to the connection string in `_load_db_connection_string()` for a dev/test server (or install a proper certificate) |
| `openpyxl is not installed` | `pip install openpyxl` |
| Every scan fails with *Year mismatch* | The year section is compared with the current year; check `year_section_start/length` and your barcodes |
| `No module named '_tkinter'` (Linux) | `sudo apt install python3-tk` |
| Config changes not showing | Save `config.xlsx` and restart the app |

## Roadmap / ideas

- Optional SQLite backend for zero-setup trials
- Configurable config / database file locations (environment variable or CLI flag)
- Unit tests for `BarcodeValidator`
- Optional work-week-must-match-current-week check
- Localisation of UI text

Contributions welcome, see [CONTRIBUTING.md](CONTRIBUTING.md).

## Disclaimer

This software is provided as-is under the MIT license. It is a validation aid, not a certified quality system. Test your configuration thoroughly before using it on a production line, and back up your database.

## License

[MIT](LICENSE)

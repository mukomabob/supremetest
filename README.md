# supremetest

Copy Supreme Brands data from Acumatica into a SQLite database on your Mac. Then check the data quality.

**Status:** Step 1 – copy the data. Step 2 – first data quality report. (Plain-English questions come later.)

## How it works

`sync.py` reads the OData service list at `https://supremebrands.acumatica.com/odata/Supreme`.
This list shows every Generic Inquiry that is published to OData (40 at the start).
The script copies each inquiry into one SQLite table.

Example: `AP-Bills and Adjustments` → table `ap_bills_and_adjustments`.

## Set up (one time)

```bash
cd supremetest
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Open `.env` and type your Acumatica username and password.
**Do not commit `.env`.** Git ignores it.

## Use

```bash
python sync.py --list                    # show the inquiries
python sync.py "AR-Customers"            # copy one inquiry (good first test)
python sync.py                           # copy all inquiries
```

Each run replaces the old data in each table.
Table `_sync_log` shows the number of rows, the time, and errors for each inquiry.

Look at the data:

```bash
sqlite3 data/supreme.db "SELECT * FROM _sync_log;"
```

Or open `data/supreme.db` with [DB Browser for SQLite](https://sqlitebrowser.org/).

## Data quality report

After `sync.py`, run:

```bash
python quality_report.py
```

It makes:

- `reports/quality_report.md` – summary of possible issues, with a possible fix for each issue
- `reports/quality_report.xlsx` – Summary sheet plus one evidence sheet for each issue
- `reports/*.csv` – the rows for each issue

Checks now: missing values and duplicate rows in all tables; duplicate supplier bills (bills only, not adjustments), invoice numbers with only words, the same supplier under
different IDs, and bills with no invoice number or description; duplicate customers and missing or incorrect
contact details.

The script only reads the database. The `reports/` folder is not sent to GitHub, because it contains business data.

## Audit tests

```bash
python audit_tests.py
```

Runs 25 audit-style tests on the full population of data (GL journals, November movement TB, suppliers,
customers, sales orders, inventory) and makes `reports/audit_tests.xlsx` with a Summary sheet, a Notes sheet
and one sheet for each test. Each result is an exception to investigate, not a proven error.
Thresholds are at the top of `audit_tests.py`.

## If an inquiry fails

- **Timeout:** set a smaller `PAGE_SIZE` in `.env` (for example 1000).
- **Too large or not needed:** add its name to `SKIP_INQUIRIES` in `.env`.
- **401 Unauthorized:** check the username, password and tenant in `.env`.

## Notes

- The script only reads from Acumatica. It does not change data.
- The data stays on your Mac. The `data/` folder is not sent to GitHub.
- Now the script uses your own login. Later, change to a read-only API user.

"""Copy all Acumatica Generic Inquiries that are published to OData into SQLite.

Usage:
    python sync.py --list                 # show the inquiries that your login can read
    python sync.py                        # copy all inquiries
    python sync.py "AR-Customers" GL-TBs  # copy only these inquiries
"""
import os
import re
import sys
import json
import sqlite3
import datetime as dt
import xml.etree.ElementTree as ET
from urllib.parse import quote

import requests
from dotenv import load_dotenv

load_dotenv()

BASE = os.getenv("ACUMATICA_URL", "").rstrip("/")
TENANT = os.getenv("ACUMATICA_TENANT", "").strip()
USER = os.getenv("ACUMATICA_USER", "")
PASSWORD = os.getenv("ACUMATICA_PASSWORD", "")
PAGE_SIZE = int(os.getenv("PAGE_SIZE", "5000"))
DB_PATH = os.getenv("DB_PATH", "data/supreme.db")
SKIP = {s.strip() for s in os.getenv("SKIP_INQUIRIES", "").split(",") if s.strip()}

ODATA_ROOT = f"{BASE}/odata/{quote(TENANT)}" if TENANT else f"{BASE}/odata"


def session() -> requests.Session:
    s = requests.Session()
    s.auth = (USER, PASSWORD)
    return s


def check(r: requests.Response):
    if r.status_code == 401:
        sys.exit("401 Unauthorized: check ACUMATICA_USER, ACUMATICA_PASSWORD and ACUMATICA_TENANT in .env.")
    r.raise_for_status()


def list_inquiries(s) -> list[str]:
    """Read the OData service document. It lists every published inquiry."""
    r = s.get(ODATA_ROOT, timeout=60)
    check(r)
    ns = {"app": "http://www.w3.org/2007/app", "atom": "http://www.w3.org/2005/Atom"}
    root = ET.fromstring(r.content)
    return [c.find("atom:title", ns).text for c in root.iterfind(".//app:collection", ns)]


def fetch_rows(s, name: str):
    """Yield all rows of one inquiry, one page at a time."""
    url = f"{ODATA_ROOT}/{quote(name)}"
    skip = 0
    while True:
        r = s.get(url, params={"$format": "json", "$top": PAGE_SIZE, "$skip": skip}, timeout=300)
        check(r)
        data = r.json()
        rows = data.get("value", [])
        yield from rows
        if len(rows) < PAGE_SIZE:
            return
        skip += PAGE_SIZE


def table_name(name: str) -> str:
    """'AP-Bills and Adjustments' -> 'ap_bills_and_adjustments'"""
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def clean(v):
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, str):
        if re.fullmatch(r"-?\d+\.\d+", v):  # OData v3 sends decimals as text, e.g. "12.5000"
            return float(v)
        m = re.fullmatch(r"/Date\((-?\d+)[+-]?\d*\)/", v)  # OData v3 date format
        if m:
            return dt.datetime.fromtimestamp(int(m.group(1)) / 1000, dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    return v


def save(con, name: str, rows: list):
    """Replace the table with the new rows, in one transaction."""
    table = table_name(name)
    rows = [{k: clean(v) for k, v in row.items() if not k.startswith(("odata.", "@odata", "__"))} for row in rows]
    cols = list(dict.fromkeys(c for row in rows for c in row))
    q = lambda c: '"' + c.replace('"', '""') + '"'
    with con:
        if cols:
            con.execute(f"DROP TABLE IF EXISTS {q(table + '__new')}")
            con.execute(f"CREATE TABLE {q(table + '__new')} ({', '.join(q(c) for c in cols)})")
            con.executemany(
                f"INSERT INTO {q(table + '__new')} VALUES ({', '.join('?' for _ in cols)})",
                [[row.get(c) for c in cols] for row in rows],
            )
            con.execute(f"DROP TABLE IF EXISTS {q(table)}")
            con.execute(f"ALTER TABLE {q(table + '__new')} RENAME TO {q(table)}")
        con.execute(
            "INSERT OR REPLACE INTO _sync_log VALUES (?, ?, ?, ?, NULL)",
            (table, name, len(rows), dt.datetime.now().isoformat(timespec="seconds")),
        )
    return table, len(rows), len(cols)


def main():
    if not (BASE and USER and PASSWORD):
        sys.exit("Copy .env.example to .env and fill in ACUMATICA_URL, ACUMATICA_USER and ACUMATICA_PASSWORD.")
    s = session()
    names = list_inquiries(s)

    if "--list" in sys.argv:
        for n in names:
            print(f"{n:45} -> {table_name(n)}{'   (skipped)' if n in SKIP else ''}")
        print(f"\n{len(names)} inquiries")
        return

    wanted = [a for a in sys.argv[1:] if not a.startswith("--")]
    if wanted:
        missing = [w for w in wanted if w not in names]
        if missing:
            sys.exit(f"Not found: {missing}. Run: python sync.py --list")
        names = wanted
    else:
        names = [n for n in names if n not in SKIP]

    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS _sync_log(
        table_name TEXT PRIMARY KEY, inquiry TEXT, rows INTEGER, synced_at TEXT, error TEXT)""")

    failed = 0
    for i, name in enumerate(names, 1):
        print(f"[{i}/{len(names)}] {name} ... ", end="", flush=True)
        try:
            table, n_rows, n_cols = save(con, name, list(fetch_rows(s, name)))
            print(f"{n_rows} rows, {n_cols} columns -> {table}")
        except (requests.RequestException, ValueError) as e:
            failed += 1
            print(f"FAILED: {e}")
            with con:
                con.execute(
                    "INSERT INTO _sync_log(table_name, inquiry, synced_at, error) VALUES (?, ?, ?, ?) "
                    "ON CONFLICT(table_name) DO UPDATE SET synced_at=excluded.synced_at, error=excluded.error",
                    (table_name(name), name, dt.datetime.now().isoformat(timespec="seconds"), str(e)[:500]),
                )
    con.close()
    print(f"\nDone. {len(names) - failed} copied, {failed} failed. Database: {DB_PATH}")


if __name__ == "__main__":
    main()

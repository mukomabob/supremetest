"""Make a first data quality report from the tables in the SQLite database.

Usage:
    python quality_report.py

Output (in the reports/ folder, which is not sent to GitHub):
    reports/quality_report.md   - the summary report (open it in any text editor or VS Code)
    reports/*.csv               - the detail rows for each issue (open in Excel)

The script only reads the database. It does not change any data.
"""
import os
import re
import csv
import sqlite3
import datetime as dt
from collections import defaultdict

from dotenv import load_dotenv

load_dotenv()
DB_PATH = os.getenv("DB_PATH", "data/supreme.db")
OUT = "reports"
# Email domains that belong to Supreme Brands itself (comma-separated in .env)
OWN_EMAIL_DOMAINS = [d.strip().lower() for d in os.getenv("OWN_EMAIL_DOMAINS", "supremebrands").split(",") if d.strip()]

# Words removed before names are compared ("ABC (Pvt) Ltd" == "abc private limited")
NAME_NOISE = r"\b(pvt|private|ltd|limited|t/a|ta|inc|co|company|the|and|&)\b"
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", re.I)


# ---------- helpers ----------

def q(c):
    return '"' + c.replace('"', '""') + '"'


def tables(con):
    return [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE '\\_%' ESCAPE '\\' ORDER BY name")]


def columns(con, t):
    return [r[1] for r in con.execute(f"PRAGMA table_info({q(t)})")]


def find_col(cols, *patterns, exclude=()):
    """Find the first column whose name contains one of the patterns (case-insensitive)."""
    for p in patterns:
        for c in cols:
            lc = c.lower()
            if p in lc and not any(x in lc for x in exclude):
                return c
    return None


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm_name(s):
    s = re.sub(r"[^\w\s&/]", " ", str(s).lower())
    s = re.sub(NAME_NOISE, " ", s)
    return re.sub(r"\s+", " ", s).strip()


def split_emails(v):
    """A field can hold more than one address, separated by ; or ,"""
    return [a.strip() for a in re.split(r"[;,]", str(v)) if a.strip()]


def is_own_email(addr):
    domain = addr.split("@")[-1].lower()
    return any(d in domain for d in OWN_EMAIL_DOMAINS)


def write_csv(name, header, rows):
    path = os.path.join(OUT, name)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    return path


class Report:
    def __init__(self):
        self.lines = []
        self.findings = []  # (area, issue, count, fix)

    def add(self, *lines):
        self.lines.extend(lines)

    def finding(self, area, issue, count, fix, detail_file=None):
        if count:
            self.findings.append((area, issue, count, fix, detail_file))


# ---------- 1. profile every table ----------

def profile(con, rep):
    rep.add("## 1. Table overview", "",
            "| Table | Rows | Columns | Empty columns | Exact duplicate rows |",
            "|---|---:|---:|---:|---:|")
    details = []
    for t in tables(con):
        cols = columns(con, t)
        n = con.execute(f"SELECT COUNT(*) FROM {q(t)}").fetchone()[0]
        if not cols or n == 0:
            rep.add(f"| {t} | {n} | {len(cols)} | – | – |")
            continue
        blanks = {}
        for c in cols:
            b = con.execute(f"SELECT COUNT(*) FROM {q(t)} WHERE {q(c)} IS NULL OR TRIM(CAST({q(c)} AS TEXT))=''").fetchone()[0]
            blanks[c] = b
        empty_cols = [c for c, b in blanks.items() if b == n]
        collist = ", ".join(q(c) for c in cols)
        dupes = con.execute(
            f"SELECT COALESCE(SUM(k-1),0) FROM (SELECT COUNT(*) k FROM {q(t)} GROUP BY {collist} HAVING k>1)").fetchone()[0]
        rep.add(f"| {t} | {n:,} | {len(cols)} | {len(empty_cols)} | {dupes:,} |")
        details.append((t, n, blanks, empty_cols))
        rep.finding(t, "Exact duplicate rows (all columns the same)", dupes,
                    "Check if the inquiry joins tables in a way that repeats rows, or if records were entered twice.")
    rep.add("")

    rep.add("### Missing values per table", "",
            "Only columns with some (not all) values missing are shown.", "")
    for t, n, blanks, empty_cols in details:
        partial = sorted(((c, b) for c, b in blanks.items() if 0 < b < n), key=lambda x: -x[1])
        rep.add(f"**{t}** ({n:,} rows)", "")
        if empty_cols:
            rep.add(f"- Always empty ({len(empty_cols)}): {', '.join(empty_cols)}")
        if partial:
            for c, b in partial[:15]:
                rep.add(f"- {c}: {b:,} empty ({b / n:.0%})")
            if len(partial) > 15:
                rep.add(f"- … and {len(partial) - 15} more columns")
        if not empty_cols and not partial:
            rep.add("- No missing values")
        rep.add("")


# ---------- 2. supplier bills ----------

def supplier_checks(con, rep):
    t = "ap_bills_and_adjustments"
    rep.add("## 2. Supplier checks (AP Bills and Adjustments)", "")
    if t not in tables(con):
        rep.add("Table not found. Run: python sync.py \"AP-Bills and Adjustments\"", "")
        return
    cols = columns(con, t)
    vid = find_col(cols, "vendorid", "vendor_id", "vendorcd", "vendor", exclude=("name", "ref", "class"))
    vname = find_col(cols, "vendorname", "vendor_name", "acctname", "name", exclude=("branch", "user", "class"))
    amount = find_col(cols, "origdocamt", "amount", "amt", "total", exclude=("tax", "disc", "balance"))
    date = find_col(cols, "docdate", "date", exclude=("due", "disc", "created", "modified"))
    vref = find_col(cols, "invoicenbr", "vendorref", "suppliersref", "extref", "vendor ref", "invoice")
    desc = find_col(cols, "docdesc", "description", "desc")
    status = find_col(cols, "status")
    rep.add(f"Columns used: vendor ID = `{vid}`, vendor name = `{vname}`, amount = `{amount}`, "
            f"date = `{date}`, vendor reference = `{vref}`, description = `{desc}`, status = `{status}`.",
            "If a column is wrong or `None`, tell Claude the correct column name.", "")

    rows = [dict(zip(cols, r)) for r in con.execute(f"SELECT * FROM {q(t)}")]

    # 2a. Same vendor reference used on more than one bill for the same vendor
    if vid and vref:
        groups = defaultdict(list)
        for r in rows:
            if not is_blank(r[vref]):
                groups[(r[vid], str(r[vref]).strip().upper())].append(r)
        dup = [g for g in groups.values() if len(g) > 1]
        n = sum(len(g) for g in dup)
        f = write_csv("supplier_duplicate_vendor_refs.csv", cols, [[r[c] for c in cols] for g in dup for r in g]) if dup else None
        rep.add(f"- **Same supplier invoice number used more than once:** {len(dup)} groups, {n} bills")
        rep.finding("Suppliers", "Same supplier invoice number entered on more than one bill", n,
                    "Check each group. If it is a real double entry, void or reverse one bill. "
                    "Turn on 'Raise an error on duplicate vendor reference number' in AP Preferences (AP101000).", f)

    # 2b. Same vendor + amount + date
    if vid and amount and date:
        groups = defaultdict(list)
        for r in rows:
            if r[amount] not in (None, 0, "0"):
                groups[(r[vid], r[amount], str(r[date])[:10])].append(r)
        dup = [g for g in groups.values() if len(g) > 1]
        n = sum(len(g) for g in dup)
        f = write_csv("supplier_possible_duplicate_bills.csv", cols, [[r[c] for c in cols] for g in dup for r in g]) if dup else None
        rep.add(f"- **Same supplier, amount and date:** {len(dup)} groups, {n} bills")
        rep.finding("Suppliers", "Possible duplicate bills (same supplier, amount and date)", n,
                    "Review with the AP team. Some can be correct (for example, two deliveries on one day).", f)

    # 2c. Same vendor ID with different names / different IDs with the same name
    if vid and vname and vid != vname:
        names_per_id = defaultdict(set)
        ids_per_name = defaultdict(set)
        for r in rows:
            if not is_blank(r[vname]):
                names_per_id[r[vid]].add(str(r[vname]).strip())
                ids_per_name[norm_name(r[vname])].add(r[vid])
        multi_ids = {k: v for k, v in ids_per_name.items() if len(v) > 1 and k}
        out = [[k, len(v), "; ".join(map(str, sorted(v, key=str)))] for k, v in sorted(multi_ids.items())]
        f = write_csv("supplier_same_name_different_ids.csv", ["normalised name", "number of IDs", "vendor IDs"], out) if out else None
        rep.add(f"- **Same supplier name under different vendor IDs:** {len(multi_ids)} names")
        rep.finding("Suppliers", "One supplier set up more than once (same name, different vendor IDs)", len(multi_ids),
                    "Confirm with AP which ID is correct. Merge or deactivate the extra IDs, so that purchases and "
                    "balances for one supplier are in one place.", f)

    # 2d. Missing reference / description
    for col, label, fix in ((vref, "Bills with no supplier invoice number",
                             "Make the supplier invoice number a required field, so that duplicates can be found."),
                            (desc, "Bills with no description",
                             "Ask the team to add a short description (what was bought)."),
                            (vid, "Bills with no vendor ID", "Investigate; every bill must have a supplier.")):
        if col:
            miss = [r for r in rows if is_blank(r[col])]
            safe = re.sub(r"\W+", "_", col.lower())
            f = write_csv(f"supplier_missing_{safe}.csv", cols,
                          [[r[c] for c in cols] for r in miss]) if miss else None
            rep.add(f"- **{label}:** {len(miss)}")
            rep.finding("Suppliers", label, len(miss), fix, f)
    rep.add("")


# ---------- 3. customers ----------

def customer_checks(con, rep):
    rep.add("## 3. Customer checks", "")
    for t in ("bi_customersummary", "customersip"):
        if t not in tables(con):
            continue
        cols = columns(con, t)
        cid = find_col(cols, "customerid", "customer_id", "customercd", "customer", "acctcd", exclude=("name", "class"))
        cname = find_col(cols, "customername", "acctname", "name", exclude=("branch", "user", "class", "sales"))
        email = find_col(cols, "email")
        phone = find_col(cols, "phone", "tel", "mobile")
        addr = find_col(cols, "addressline1", "address", "street", "city")
        rep.add(f"**{t}** – columns used: ID = `{cid}`, name = `{cname}`, email = `{email}`, "
                f"phone = `{phone}`, address = `{addr}`", "")
        rows = [dict(zip(cols, r)) for r in con.execute(f"SELECT * FROM {q(t)}")]

        # one row per customer for the checks below
        if cid:
            seen, uniq = set(), []
            for r in rows:
                if r[cid] not in seen:
                    seen.add(r[cid])
                    uniq.append(r)
            if len(uniq) < len(rows):
                rep.add(f"- Table has {len(rows):,} rows for {len(uniq):,} customers (more than one row per customer).")
            rows = uniq

        if cid and cname and cid != cname:
            ids_per_name = defaultdict(set)
            for r in rows:
                if not is_blank(r[cname]):
                    ids_per_name[norm_name(r[cname])].add(r[cid])
            multi = {k: v for k, v in ids_per_name.items() if len(v) > 1 and k}
            out = [[k, len(v), "; ".join(map(str, sorted(v, key=str)))] for k, v in sorted(multi.items())]
            f = write_csv(f"{t}_same_name_different_ids.csv", ["normalised name", "number of IDs", "customer IDs"], out) if out else None
            rep.add(f"- **Same customer name under different IDs:** {len(multi)} names")
            rep.finding(f"Customers ({t})", "One customer set up more than once (same name, different IDs)", len(multi),
                        "Confirm with AR which ID is correct. Merge or deactivate the extra accounts.", f)

        for col, label in ((email, "no email"), (phone, "no phone number"), (addr, "no address")):
            if col:
                miss = [r for r in rows if is_blank(r[col])]
                rep.add(f"- **Customers with {label}:** {len(miss):,} of {len(rows):,}")
                rep.finding(f"Customers ({t})", f"Customers with {label}", len(miss),
                            "Collect the missing contact details, starting with active customers.")
        if email:
            bad, own = [], []
            for r in rows:
                if is_blank(r[email]):
                    continue
                addrs = split_emails(r[email])
                if any(not EMAIL_RE.match(a) for a in addrs):
                    bad.append(r)
                if any(is_own_email(a) for a in addrs):
                    own.append(r)
            f = write_csv(f"{t}_invalid_emails.csv", cols, [[r[c] for c in cols] for r in bad]) if bad else None
            rep.add(f"- **Customers with an email in an incorrect format:** {len(bad)}")
            rep.finding(f"Customers ({t})", "Email addresses in an incorrect format", len(bad),
                        "Correct the email addresses. Invoices and statements sent by email will fail.", f)
            f = write_csv(f"{t}_own_company_emails.csv", cols, [[r[c] for c in cols] for r in own]) if own else None
            rep.add(f"- **Customer records with a Supreme Brands email address:** {len(own)}")
            rep.finding(f"Customers ({t})", "Customer record uses a Supreme Brands email address", len(own),
                        "Get the customer's own email address. Now, statements and invoices for these customers "
                        "go to Supreme Brands, not to the customer.", f)
        rep.add("")


# ---------- main ----------

def main():
    if not os.path.exists(DB_PATH):
        raise SystemExit(f"No database at {DB_PATH}. Run: python sync.py")
    os.makedirs(OUT, exist_ok=True)
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    rep = Report()

    sync = con.execute("SELECT inquiry, rows, synced_at FROM _sync_log WHERE error IS NULL ORDER BY inquiry").fetchall()

    profile(con, rep)
    supplier_checks(con, rep)
    customer_checks(con, rep)

    head = [
        "# Supreme Brands – first data quality report",
        "",
        f"Made: {dt.datetime.now():%d %b %Y %H:%M}. Source: {len(sync)} Acumatica inquiries copied with OData.",
        "",
        "**Important:** These are possible issues, found by automatic checks. Check each one before you change data.",
        "",
        "## Summary of issues and possible fixes",
        "",
        "| Area | Issue | Count | Possible fix | Detail file |",
        "|---|---|---:|---|---|",
    ]
    for area, issue, count, fix, f in sorted(rep.findings, key=lambda x: (x[0], -x[2])):
        head.append(f"| {area} | {issue} | {count:,} | {fix} | {os.path.basename(f) if f else ''} |")
    if not rep.findings:
        head.append("| – | No issues found | – | – | – |")
    head += ["", "### Data used", ""] + [f"- {i}: {r:,} rows (copied {s})" for i, r, s in sync] + [""]

    path = os.path.join(OUT, "quality_report.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(head + rep.lines))
    con.close()
    print(f"Report: {path}")
    print(f"Detail files: {OUT}/*.csv")
    print(f"{len(rep.findings)} types of issue found.")


if __name__ == "__main__":
    main()

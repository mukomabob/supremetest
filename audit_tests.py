"""Audit tests on the full population of Acumatica data in the SQLite database.

Usage:
    python audit_tests.py

Output (in reports/, which is not sent to GitHub):
    reports/audit_tests.xlsx  - Summary sheet, Notes sheet, and one sheet for each test

The script only reads the database. Each result is an exception to investigate, not a proven error.
"""
import os
import sqlite3
import datetime as dt

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

DB_PATH = os.getenv("DB_PATH", "data/supreme.db")
OUT = "reports"
# Hours to add to Acumatica "CreatedOn" times to get Harare time (if the API gives UTC, set 2).
TIME_OFFSET = int(os.getenv("CREATED_ON_UTC_OFFSET", "0"))
AFTER_HOURS_START, AFTER_HOURS_END = 20, 6       # 20:00 to 05:59 counts as "after hours"
ROUND_GL = 1000                                   # manual journal totals that are a multiple of this
ROUND_CASH = 100                                  # cash purchases that are a multiple of this (and >= it)
SPLIT_MIN_DOCS = 5                                # same supplier, same day: this many or more cash purchases
CREDIT_MEMO_DAYS = 5                              # credit memos in the first N days of a month
OPEN_ORDER_DAYS = 30                              # sales orders still open after N days
MAX_SHEET_ROWS = 100_000

TESTS = []


def test(test_id, area, title, why, follow_up):
    def wrap(fn):
        TESTS.append(dict(id=test_id, area=area, title=title, why=why, follow_up=follow_up, fn=fn))
        return fn
    return wrap


def local(col):
    """SQL expression: CreatedOn moved to Harare time."""
    return f"datetime({col}, '{TIME_OFFSET:+d} hours')"


# ---------------- GL journals ----------------

@test("GL-01", "GL journals", "Manual GL journal batches (all)",
      "Manual journals bypass the normal sub-ledger controls. This is the population for the next GL tests.",
      "Check that large or unusual manual journals have support and approval.")
def gl_manual(c):
    return c.execute("""
        SELECT BatchNumber, Status, substr(TransactionDate,1,10) AS TranDate, PostPeriod, Description,
               Currency, ControlTotal, Type, ReversingEntry, CreatedBy, CreatedOn, LastModifiedBy
        FROM gl_journal_transactions WHERE Module='GL' ORDER BY CreatedOn""")


@test("GL-02", "GL journals", "Manual GL batches made on a weekend or after hours",
      f"Journals made outside normal hours ({AFTER_HOURS_START}:00-0{AFTER_HOURS_END}:00) or on weekends get less review.",
      "Ask the person who made each one why it was made at that time. Check the support.")
def gl_after_hours(c):
    t = local("CreatedOn")
    return c.execute(f"""
        SELECT BatchNumber, Status, substr(TransactionDate,1,10) AS TranDate, Description, Currency, ControlTotal,
               CreatedBy, {t} AS CreatedLocal,
               CASE strftime('%w',{t}) WHEN '0' THEN 'Sunday' WHEN '6' THEN 'Saturday' ELSE 'Weekday' END AS Day
        FROM gl_journal_transactions
        WHERE Module='GL' AND (strftime('%w',{t}) IN ('0','6')
              OR CAST(strftime('%H',{t}) AS INT) >= {AFTER_HOURS_START}
              OR CAST(strftime('%H',{t}) AS INT) < {AFTER_HOURS_END})
        ORDER BY {t}""")


@test("GL-03", "GL journals", f"Manual GL batches with a round total (multiple of {ROUND_GL:,})",
      "Round amounts often are estimates or adjustments, not amounts from real documents.",
      "Check the calculation and the support for each journal.")
def gl_round(c):
    return c.execute(f"""
        SELECT BatchNumber, Status, substr(TransactionDate,1,10) AS TranDate, Description, Currency, ControlTotal,
               CreatedBy, CreatedOn
        FROM gl_journal_transactions
        WHERE Module='GL' AND ABS(ControlTotal) >= {ROUND_GL} AND ABS(ControlTotal) % {ROUND_GL} = 0
        ORDER BY ABS(ControlTotal) DESC""")


@test("GL-04", "GL journals", "Batches posted to a period different from the transaction date",
      "A posting period that does not agree with the document date can move income or costs between months.",
      "Review the large ones, especially at month and year end.")
def gl_period_mismatch(c):
    return c.execute("""
        SELECT Module, BatchNumber, substr(TransactionDate,1,10) AS TranDate, PostPeriod,
               strftime('%m%Y', TransactionDate) AS DatePeriod, Description, Currency, ControlTotal, CreatedBy, CreatedOn
        FROM gl_journal_transactions
        WHERE PostPeriod IS NOT NULL AND PostPeriod <> strftime('%m%Y', TransactionDate)
        ORDER BY Module, TransactionDate""")


@test("GL-05", "GL journals", "Batches made more than 30 days after the transaction date (back-dated)",
      "Back-dated entries can change results of a period that is already reported.",
      "Confirm that each back-dated entry was approved and that the period was still open for a good reason.")
def gl_backdated(c):
    return c.execute("""
        SELECT Module, BatchNumber, substr(TransactionDate,1,10) AS TranDate, PostPeriod, substr(CreatedOn,1,10) AS CreatedDate,
               CAST(julianday(CreatedOn) - julianday(TransactionDate) AS INT) AS DaysLate,
               Description, Currency, ControlTotal, CreatedBy
        FROM gl_journal_transactions
        WHERE julianday(CreatedOn) - julianday(TransactionDate) > 30
        ORDER BY DaysLate DESC""")


@test("GL-06", "GL journals", "Batches per user and module (information)",
      "Shows who posts what. A user who posts in many modules can be a segregation-of-duties risk.",
      "Compare with job roles. Ask why one user posts in AP, AR and GL.")
def gl_users(c):
    return c.execute("""
        SELECT CreatedBy, Module, COUNT(*) AS Batches, MIN(substr(TransactionDate,1,10)) AS FirstDate,
               MAX(substr(TransactionDate,1,10)) AS LastDate
        FROM gl_journal_transactions GROUP BY CreatedBy, Module ORDER BY CreatedBy, Batches DESC""")


# ---------------- November movement TB ----------------

@test("TB-01", "Trial balance", "November 2026 movement per account (information)",
      "The account-level data covers only November 2026. This is the movement, not the closing balance.",
      "Compare with the trial balance from Acumatica when access to GL-TBs is given.")
def tb_movement(c):
    return c.execute("""
        SELECT Account, Module, ROUND(SUM(DebitAmountinBaseCurrency),2) AS Debit,
               ROUND(SUM(CreditAmountinBaseCurrency),2) AS Credit,
               ROUND(SUM(DebitAmountinBaseCurrency)-SUM(CreditAmountinBaseCurrency),2) AS Net, COUNT(*) AS Lines
        FROM api_accountdetails GROUP BY Account, Module ORDER BY Account, Module""")


@test("TB-02", "Trial balance", "November batches where debits do not equal credits",
      "Every batch must balance. A difference shows a data or system problem.",
      "Investigate each batch in Acumatica.")
def tb_unbalanced(c):
    return c.execute("""
        SELECT Module, BatchNumber, ROUND(SUM(DebitAmountinBaseCurrency),2) AS Debit,
               ROUND(SUM(CreditAmountinBaseCurrency),2) AS Credit,
               ROUND(SUM(DebitAmountinBaseCurrency)-SUM(CreditAmountinBaseCurrency),2) AS Difference
        FROM api_accountdetails GROUP BY Module, BatchNumber
        HAVING ABS(SUM(DebitAmountinBaseCurrency)-SUM(CreditAmountinBaseCurrency)) > 0.01""")


# ---------------- Suppliers ----------------

@test("AP-01", "Suppliers", "Possible duplicate bills (same supplier, amount, date and currency)",
      "Duplicate bills can lead to paying a supplier twice.",
      "Compare the documents. If a bill is a duplicate, reverse it and check if it was paid.")
def ap_dup_bills(c):
    return c.execute("""
        WITH b AS (SELECT *, trim(Vendor) AS V FROM ap_bills_and_adjustments WHERE Type='Bill' AND Amount<>0)
        SELECT b.ReferenceNbr, b.V AS Vendor, b.VendorName, substr(b.Date,1,10) AS Date, b.Currency, b.Amount,
               b.VendorRef, b.Description, b.Status, b.CreatedBy, g.n AS InGroup
        FROM b JOIN (SELECT V, Amount, substr(Date,1,10) d, Currency, COUNT(*) n FROM b
                     GROUP BY 1,2,3,4 HAVING n>1) g
          ON b.V=g.V AND b.Amount=g.Amount AND substr(b.Date,1,10)=g.d AND b.Currency=g.Currency
        ORDER BY b.V, b.Date, b.Amount""")


@test("AP-02", "Suppliers", "Possible duplicate payments (same supplier, amount, date and currency)",
      "Two payments of the same amount to one supplier on one day can be a double payment.",
      "Check the bank statement and the supplier statement.")
def ap_dup_payments(c):
    return c.execute("""
        WITH p AS (SELECT *, trim(Vendor) AS V FROM ap_checks_and_payments
                   WHERE Type IN ('Payment','Prepayment') AND Status <> 'Voided' AND PaymentAmount<>0)
        SELECT p.ReferenceNbr, p.Type, p.V AS Vendor, p.VendorName, substr(p.PaymentDate,1,10) AS Date, p.Currency,
               p.PaymentAmount, p.PaymentRef, p.Description, p.CashAccount, p.Status, p.CreatedBy, g.n AS InGroup
        FROM p JOIN (SELECT V, PaymentAmount, substr(PaymentDate,1,10) d, Currency, COUNT(*) n FROM p
                     GROUP BY 1,2,3,4 HAVING n>1) g
          ON p.V=g.V AND p.PaymentAmount=g.PaymentAmount AND substr(p.PaymentDate,1,10)=g.d AND p.Currency=g.Currency
        ORDER BY p.V, p.PaymentDate""")


@test("AP-03", "Suppliers", "Suppliers paid more than they billed (per currency)",
      "Payments above the total of bills can show overpayments, unrecorded bills or prepayments that were never used.",
      "Get a supplier statement and reconcile.")
def ap_overpaid(c):
    return c.execute("""
        WITH bills AS (SELECT trim(Vendor) V, Currency,
                         SUM(CASE WHEN Type='Debit Adj.' THEN -Amount ELSE Amount END) AS Billed,
                         MAX(VendorName) AS VendorName
                       FROM ap_bills_and_adjustments GROUP BY 1,2),
             pays AS (SELECT trim(Vendor) V, Currency, SUM(PaymentAmount) AS Paid, MAX(VendorName) AS VendorName
                      FROM ap_checks_and_payments
                      WHERE Type IN ('Payment','Prepayment') AND Status <> 'Voided' GROUP BY 1,2)
        SELECT p.V AS Vendor, COALESCE(b.VendorName, p.VendorName) AS VendorName, p.Currency,
               ROUND(COALESCE(b.Billed,0),2) AS Billed, ROUND(p.Paid,2) AS Paid,
               ROUND(p.Paid-COALESCE(b.Billed,0),2) AS PaidMoreBy
        FROM pays p LEFT JOIN bills b ON p.V=b.V AND p.Currency=b.Currency
        WHERE p.Paid > COALESCE(b.Billed,0) * 1.01 + 1
        ORDER BY PaidMoreBy DESC""")


@test("AP-04", "Suppliers", f"Cash purchases with a round amount (multiple of {ROUND_CASH})",
      "Round cash amounts are often estimates or advances, not real receipts.",
      "Check that each one has a receipt for the exact amount.")
def ap_cash_round(c):
    return c.execute(f"""
        SELECT ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(Date,1,10) AS Date, Currency, PaymentAmount,
               Description, CashAccount, Branch, Status, CreatedBy
        FROM ap_cash_purchases
        WHERE Type='Cash Purchase' AND Status <> 'Voided' AND PaymentAmount >= {ROUND_CASH}
              AND PaymentAmount % {ROUND_CASH} = 0
        ORDER BY PaymentAmount DESC""")


@test("AP-05", "Suppliers", f"Possible split cash purchases ({SPLIT_MIN_DOCS}+ to one supplier on one day)",
      "Splitting one purchase into small amounts can avoid approval limits.",
      "Check if the purchases are for the same item or order. Compare with the approval limit.")
def ap_cash_split(c):
    return c.execute(f"""
        WITH p AS (SELECT *, trim(Vendor) V, substr(Date,1,10) D FROM ap_cash_purchases
                   WHERE Type='Cash Purchase' AND Status <> 'Voided')
        SELECT p.V AS Vendor, p.VendorName, p.D AS Date, p.Currency, p.Branch, p.CreatedBy, g.n AS DocsThatDay,
               ROUND(g.total,2) AS TotalThatDay, p.ReferenceNbr, p.PaymentAmount, p.Description
        FROM p JOIN (SELECT V, D, Currency, CreatedBy, COUNT(*) n, SUM(PaymentAmount) total FROM p
                     GROUP BY 1,2,3,4 HAVING n >= {SPLIT_MIN_DOCS}) g
          ON p.V=g.V AND p.D=g.D AND p.Currency=g.Currency AND p.CreatedBy=g.CreatedBy
        ORDER BY g.total DESC, p.D, p.V""")


@test("AP-06", "Suppliers", "Cash purchases by user and supplier (information)",
      "Most cash purchases use a general 'Cash Supplier'. Then the real supplier is not recorded.",
      "Ask if large or frequent suppliers must be set up as real vendors.")
def ap_cash_summary(c):
    return c.execute("""
        SELECT trim(Vendor) AS Vendor, VendorName, Currency, CreatedBy, COUNT(*) AS Docs,
               ROUND(SUM(PaymentAmount),2) AS Total, ROUND(MAX(PaymentAmount),2) AS Largest
        FROM ap_cash_purchases WHERE Type='Cash Purchase' AND Status <> 'Voided'
        GROUP BY 1,2,3,4 ORDER BY Total DESC""")


@test("AP-07", "Suppliers", "Supplier documents changed by a person different from the one who made them",
      "A change by another person can be a normal approval, or a change after review. Shows who edits whose work.",
      "Look at the people with the most changes. Confirm that the changes are part of their role.")
def ap_modified(c):
    return c.execute("""
        SELECT 'Bill/Adjustment' AS Source, CreatedBy, LastModifiedBy, COUNT(*) AS Docs
          FROM ap_bills_and_adjustments WHERE CreatedBy <> LastModifiedBy GROUP BY 2,3
        UNION ALL
        SELECT 'Payment', CreatedBy, LastModifiedBy, COUNT(*) FROM ap_checks_and_payments
          WHERE CreatedBy <> LastModifiedBy GROUP BY 2,3
        UNION ALL
        SELECT 'Cash purchase', CreatedBy, LastModifiedBy, COUNT(*) FROM ap_cash_purchases
          WHERE CreatedBy <> LastModifiedBy GROUP BY 2,3
        ORDER BY Docs DESC""")


# ---------------- Customers ----------------

@test("AR-01", "Customers", f"Credit memos in the first {CREDIT_MEMO_DAYS} days of a month",
      "Credit memos just after a month end can reverse sales that were booked to make the last month look better.",
      "Match each large one to the invoice it reverses. Check the reason and the approval.")
def ar_cm_after_month_end(c):
    return c.execute(f"""
        SELECT ReferenceNbr, trim(Customer) AS Customer, CustomerName, substr(Date,1,10) AS Date, PostPeriod, Currency,
               Amount, Description, OrigRefNbr, Status, CreatedBy
        FROM ar_invoices_and_memos
        WHERE Type='Credit Memo' AND CAST(strftime('%d', Date) AS INT) <= {CREDIT_MEMO_DAYS}
        ORDER BY Amount DESC""")


@test("AR-02", "Customers", "Possible duplicate invoices (same customer, amount, date and currency)",
      "Duplicate invoices overstate sales and can upset customers.",
      "Compare with the sales orders and the delivery notes.")
def ar_dup_invoices(c):
    return c.execute("""
        WITH i AS (SELECT *, trim(Customer) C FROM ar_invoices_and_memos
                   WHERE Type='Invoice' AND Status <> 'Canceled' AND Amount<>0)
        SELECT i.ReferenceNbr, i.C AS Customer, i.CustomerName, substr(i.Date,1,10) AS Date, i.Currency, i.Amount,
               i.Description, i.Status, i.CreatedBy, g.n AS InGroup
        FROM i JOIN (SELECT C, Amount, substr(Date,1,10) d, Currency, COUNT(*) n FROM i
                     GROUP BY 1,2,3,4 HAVING n>1) g
          ON i.C=g.C AND i.Amount=g.Amount AND substr(i.Date,1,10)=g.d AND i.Currency=g.Currency
        ORDER BY i.C, i.Date, i.Amount""")


@test("AR-03", "Customers", "Credit memos by user (information)",
      "Shows who issues credit notes and how much. Large totals from one person need review.",
      "Confirm that credit memos need approval, and who approves them.")
def ar_cm_users(c):
    return c.execute("""
        SELECT CreatedBy, Currency, COUNT(*) AS CreditMemos, ROUND(SUM(Amount),2) AS Total, ROUND(MAX(Amount),2) AS Largest
        FROM ar_invoices_and_memos WHERE Type='Credit Memo' GROUP BY 1,2 ORDER BY Total DESC""")


@test("AR-04", "Customers", "Customer documents changed by a person different from the one who made them",
      "Shows who edits invoices after they are made (including the 'apiuser' integration).",
      "Confirm what 'apiuser' changes, and that other changes are part of the person's role.")
def ar_modified(c):
    return c.execute("""
        SELECT Type, CreatedBy, LastModifiedBy, COUNT(*) AS Docs FROM ar_invoices_and_memos
        WHERE CreatedBy <> LastModifiedBy GROUP BY 1,2,3 ORDER BY Docs DESC""")


# ---------------- Sales orders ----------------

@test("SO-01", "Sales orders", "Sales order lines with a manual price",
      "A price typed by hand overrides the price list. It can give unapproved discounts.",
      "Compare the manual prices with the price list. Find who sets them and why.")
def so_manual_price(c):
    return c.execute("""
        SELECT OrderType, OrderNbr, LineNbr, substr(Date,1,10) AS Date, Status, trim(Customer) AS Customer, CustomerName,
               InventoryID, Description_2 AS Item, LineOrderQTY AS Qty, Amount, Currency, DefaultSalesperson, CreatedBy
        FROM so_salesorderdetails WHERE ManualPrice=1 AND Behavior='Sales Order'
        ORDER BY Date DESC""")


@test("SO-02", "Sales orders", "Manual prices by user and salesperson (information)",
      "Shows who uses manual prices most.",
      "Discuss the top users with the sales manager.")
def so_manual_summary(c):
    return c.execute("""
        SELECT CreatedBy, DefaultSalesperson, COUNT(*) AS Lines,
               SUM(ManualPrice) AS ManualPriceLines, ROUND(100.0*SUM(ManualPrice)/COUNT(*),1) AS PctManual,
               ROUND(SUM(CASE WHEN ManualPrice=1 THEN Amount ELSE 0 END),2) AS ManualPriceAmount
        FROM so_salesorderdetails WHERE Behavior='Sales Order' GROUP BY 1,2 ORDER BY ManualPriceLines DESC""")


@test("SO-03", "Sales orders", "Lines invoiced for more than the ordered quantity",
      "Invoicing more than was ordered can overcharge customers or show data errors.",
      "Check the order, the shipment and the invoice.")
def so_over_invoiced(c):
    return c.execute("""
        SELECT OrderType, OrderNbr, LineNbr, substr(Date,1,10) AS Date, Status, trim(Customer) AS Customer, CustomerName,
               InventoryID, LineOrderQTY, InvoicedQTY, Amount, InvoicedAmount, CreatedBy
        FROM so_salesorderdetails WHERE InvoicedQTY > LineOrderQTY + 0.0001 ORDER BY Date DESC""")


@test("SO-04", "Sales orders", f"Sales orders still open after {OPEN_ORDER_DAYS} days and not invoiced",
      "Old open orders can be lost sales, or orders that were delivered but not invoiced.",
      "Ask sales and dispatch if the goods were delivered. Close the orders that are no longer needed.")
def so_old_open(c):
    return c.execute(f"""
        SELECT OrderType, OrderNbr, substr(Date,1,10) AS Date, Status, trim(Customer) AS Customer, CustomerName,
               Currency, MAX(OrderTotal) AS OrderTotal, SUM(InvoicedQTY) AS InvoicedQty,
               CAST(julianday('now') - julianday(MAX(Date)) AS INT) AS DaysOpen, MAX(CreatedBy) AS CreatedBy
        FROM so_salesorderdetails
        WHERE Behavior='Sales Order' AND Status IN ('Open','Back Order','Shipping','Awaiting Payment','On Hold','Credit Hold')
          AND julianday('now') - julianday(Date) > {OPEN_ORDER_DAYS}
        GROUP BY OrderType, OrderNbr HAVING SUM(InvoicedQTY)=0 ORDER BY Date""")


# ---------------- Inventory ----------------

@test("IN-01", "Inventory", "Stock on hand with zero cost",
      "Stock with no cost understates inventory and cost of sales.",
      "Find the receipt or production cost and correct it.")
def in_zero_cost(c):
    return c.execute("""
        SELECT InventoryID, Description, Warehouse, WarehouseName, QuantityonHand, InventoryCost, ItemStatus
        FROM in_inventorybalance WHERE QuantityonHand > 0 AND COALESCE(InventoryCost,0) = 0""")


@test("IN-02", "Inventory", "Same item with very different unit cost in different warehouses (2x or more)",
      "A large difference in unit cost for one item can show costing errors.",
      "Check the receipts and transfers for these items.")
def in_cost_spread(c):
    return c.execute("""
        WITH u AS (SELECT InventoryID, Description, Warehouse, QuantityonHand, InventoryCost,
                          InventoryCost/QuantityonHand AS UnitCost
                   FROM in_inventorybalance WHERE QuantityonHand > 0 AND InventoryCost > 0)
        SELECT u.InventoryID, u.Description, u.Warehouse, u.QuantityonHand, u.InventoryCost, ROUND(u.UnitCost,4) AS UnitCost,
               ROUND(g.mn,4) AS MinUnitCost, ROUND(g.mx,4) AS MaxUnitCost
        FROM u JOIN (SELECT InventoryID, MIN(UnitCost) mn, MAX(UnitCost) mx FROM u GROUP BY 1
                     HAVING MAX(UnitCost) >= 2*MIN(UnitCost)) g ON u.InventoryID=g.InventoryID
        ORDER BY g.mx/g.mn DESC, u.InventoryID""")


# ---------------- run and write ----------------

def load_numeric_copy(path):
    """Copy the database into memory and store decimal text ("12.5000") as numbers.

    OData v3 sends Edm.Decimal values as text. Without this, SQLite compares them as text ('10.0' < '5.0').
    Only columns where every value is a decimal number with a '.' are changed, so codes like account
    numbers ("500150") and periods ("112026") stay as text. The database file is not changed.
    """
    src = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    con = sqlite3.connect(":memory:")
    src.backup(con)
    src.close()
    tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    for t in tables:
        for col in [r[1] for r in con.execute(f'PRAGMA table_info("{t}")')]:
            qc = '"' + col.replace('"', '""') + '"'
            texts, numeric = con.execute(
                f"SELECT COUNT(*), SUM({qc} GLOB '*.*' AND {qc} NOT GLOB '*[^0-9.-]*' AND {qc} GLOB '*[0-9]*') "
                f'FROM "{t}" WHERE typeof({qc})=\'text\'').fetchone()
            if texts and texts == numeric:
                con.execute(f'UPDATE "{t}" SET {qc} = CAST({qc} AS REAL) WHERE typeof({qc})=\'text\'')
    con.commit()
    return con


INFO_TESTS = {"GL-01", "GL-06", "TB-01", "AP-06", "AP-07", "AR-03", "AR-04", "SO-02"}


def write_workbook(results, notes, path, title="Supreme Brands – audit tests on Acumatica data", info=None):
    info = INFO_TESTS if info is None else info
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    def safe(v):  # remove control characters that Excel does not accept
        return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v

    f, b = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    fill = PatternFill("solid", start_color="D9E1F2")
    link = Font(name="Arial", size=10, color="0563C1", underline="single")
    wrap = Alignment(wrap_text=True, vertical="top")

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = (f"Made {dt.datetime.now():%d %b %Y %H:%M}. Each result is an exception to investigate, "
                "not a proven error. 'Information' sheets are analysis, not exceptions.")
    ws["A2"].font = f
    hdr = ["Test", "Area", "Test", "Type", "Rows", "Why it matters", "Follow-up", "Sheet"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.fill = b, fill
    for r, t in enumerate(results, 5):
        name = t["id"]
        vals = [t["id"], t["area"], t["title"], "Information" if name in info else "Exception",
                f"=MAX(COUNTA('{name}'!A:A)-3,0)", t["why"], t["follow_up"], name]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = f
            if i in (3, 6, 7):
                c.alignment = wrap
        ws.cell(row=r, column=5).number_format = "#,##0"
        s = ws.cell(row=r, column=8)
        s.hyperlink, s.font = f"#'{name}'!A1", link
    for col, w in zip("ABCDEFGH", (8, 14, 45, 12, 9, 50, 50, 8)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    n = wb.create_sheet("Notes")
    n["A1"] = "Notes on the data and the tests"
    n["A1"].font = Font(name="Arial", size=12, bold=True)
    for i, line in enumerate(notes, 3):
        n.cell(row=i, column=1, value=line).font = f
    n.column_dimensions["A"].width = 140

    for t in results:
        d = wb.create_sheet(t["id"])
        d["A1"] = f"{t['id']} – {t['title']}"
        d["A1"].font = Font(name="Arial", size=12, bold=True)
        d["A2"] = t["why"]
        d["A2"].font = Font(name="Arial", size=10, italic=True)
        for ci, h in enumerate(t["cols"], 1):
            c = d.cell(row=3, column=ci, value=h)
            c.font, c.fill = b, fill
        for ri, row in enumerate(t["rows"][:MAX_SHEET_ROWS], 4):
            for ci, v in enumerate(row, 1):
                d.cell(row=ri, column=ci, value=safe(v)).font = f
        if len(t["rows"]) > MAX_SHEET_ROWS:
            d.cell(row=MAX_SHEET_ROWS + 5, column=1,
                   value=f"Only the first {MAX_SHEET_ROWS:,} of {len(t['rows']):,} rows are shown.").font = b
        for ci, h in enumerate(t["cols"], 1):
            w = max([len(str(h))] + [len(str(r[ci - 1])) for r in t["rows"][:200]])
            d.column_dimensions[get_column_letter(ci)].width = min(max(w + 2, 8), 40)
        if t["cols"]:
            d.auto_filter.ref = f"A3:{get_column_letter(len(t['cols']))}{min(len(t['rows']), MAX_SHEET_ROWS) + 3}"
        d.freeze_panes = "A4"
    wb.save(path)


def main():
    if not os.path.exists(DB_PATH):
        raise SystemExit(f"No database at {DB_PATH}. Run: python sync.py")
    os.makedirs(OUT, exist_ok=True)
    print("Loading data ...")
    con = load_numeric_copy(DB_PATH)
    tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}

    results = []
    for t in TESTS:
        try:
            cur = t["fn"](con)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
            print(f"{t['id']:6} {len(rows):>8,}  {t['title']}")
        except sqlite3.OperationalError as e:  # table or column missing
            cols, rows = ["Not run"], [[f"Test not run: {e}"]]
            print(f"{t['id']:6}  NOT RUN  {t['title']}  ({e})")
        results.append(dict(t, cols=cols, rows=rows))

    sync = con.execute("SELECT inquiry, rows, synced_at FROM _sync_log WHERE error IS NULL ORDER BY inquiry").fetchall() \
        if "_sync_log" in tables else []
    span = con.execute("SELECT MIN(substr(TransactionDate,1,10)), MAX(substr(TransactionDate,1,10)) "
                       "FROM gl_journal_transactions").fetchone() if "gl_journal_transactions" in tables else (None, None)
    con.close()

    notes = [
        f"Database: {DB_PATH}. GL batches cover {span[0]} to {span[1]}.",
        "Account-level data (API-AccountDetails) covers November 2026 only, so a full trial balance is not possible yet (no access to GL-TBs).",
        "GL-Journal Transactions has one row per batch (batch total only, no accounts).",
        "Amounts are in the document currency (USD, ZAR, ZWG, ZWL, EUR). Do not add amounts of different currencies.",
        f"'After hours' = {AFTER_HOURS_START}:00 to 0{AFTER_HOURS_END}:00 and weekends, using CreatedOn + {TIME_OFFSET} hours. "
        "If Acumatica gives CreatedOn in UTC, set CREATED_ON_UTC_OFFSET=2 in .env and run again.",
        "Thresholds (round amounts, days, split count) are set at the top of audit_tests.py. Change them to suit the business.",
        "",
        "Data used:",
    ] + [f"  {i}: {r:,} rows (copied {s})" for i, r, s in sync]

    path = os.path.join(OUT, "audit_tests.xlsx")
    write_workbook(results, notes, path)
    print(f"\nExcel: {path}")


if __name__ == "__main__":
    main()

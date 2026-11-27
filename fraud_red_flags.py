"""Red-flag tests for ways money could be taken, on the full population of Acumatica data.

Usage:
    python fraud_red_flags.py

Output: reports/fraud_red_flags.xlsx (not sent to GitHub). It contains staff names: share only with
the people who must see it.

A red flag is a reason to look closer. It is NOT proof that anyone did something wrong.
"""
import os
import math
import sqlite3
import statistics
from collections import Counter, defaultdict

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from audit_tests import load_numeric_copy, write_workbook, local, TIME_OFFSET

DB_PATH = os.getenv("DB_PATH", "data/supreme.db")
START = os.getenv("PERIOD_START", "2023-07-01")
END = os.getenv("PERIOD_END", "2026-09-30")
OUT = "reports"

LATE_DAYS = 7              # receipt recorded more than N days after the payment date
OPEN_PREPAY_DAYS = 60      # supplier prepayment still open after N days
REISSUE_DAYS = 30          # voided payment, then a new payment of the same amount within N days
ROUND_PAYMENT = 1000       # supplier payments that are a multiple of this
VAGUE_MIN = 100            # cash purchases with a vague description, this amount or more
LOW_PRICE = 0.70           # manual price below 70% of the item's median unit price
SELF_MIN = 500             # cash purchases made and last changed by one person, this amount or more
KEYWORDS = ["personal", "loan", "advance", "shareholder", "director", "staff", "salary", "allowance",
            "gift", "donation", "sundry", "misc", "write off", "write-off", "cash out", "test", "refund"]
VAGUE_WORDS = {"", "payment", "cash", "purchase", "purchases", "misc", "sundry", "sundries", "expenses",
               "expense", "various", "goods", "items", "stock", "cash purchase", "petty cash", "na", "n/a"}

TESTS = []


def test(test_id, area, title, why, follow_up):
    def wrap(fn):
        TESTS.append(dict(id=test_id, area=area, title=title, why=why, follow_up=follow_up, fn=fn))
        return fn
    return wrap


def period(col):
    return f"substr({col},1,10) BETWEEN '{START}' AND '{END}'"


def rows_of(cur):
    return [d[0] for d in cur.description], cur.fetchall()


# ---------------- 1. Customer receipts ----------------

@test("FR-01", "Customer cash", "Voided customer receipts",
      "Cash can be received, the receipt recorded, and then voided so that the cash can be kept.",
      "Match each voided receipt to the bank or cash book. Ask why it was voided and who approved it.")
def voided_receipts(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Customer) AS Customer, CustomerName, substr(PaymentDate,1,10) AS PaymentDate,
               Currency, PaymentAmount, PaymentMethod, trim(CashAccount) AS CashAccount, PaymentRef, Description,
               Branch, CreatedBy, LastModifiedBy, substr(LastModifiedOn,1,10) AS VoidedAbout
        FROM ar_payments_and_applications
        WHERE Type='Payment' AND Status='Voided' AND {period('PaymentDate')}
        ORDER BY PaymentAmount DESC"""))


@test("FR-02", "Customer cash", "Credit memos with no link to an invoice",
      "A credit memo can hide a payment that was taken: the customer's balance is cleared without cash. "
      "Most credit memos here have no invoice link, so this control is weak.",
      "Start with the largest. Get the reason, the approval and the customer's confirmation.")
def cm_no_link(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Customer) AS Customer, CustomerName, substr(Date,1,10) AS Date, Currency, Amount,
               Description, Status, Branch, CreatedBy
        FROM ar_invoices_and_memos
        WHERE Type='Credit Memo' AND (OrigRefNbr IS NULL OR trim(OrigRefNbr)='') AND {period('Date')}
        ORDER BY Amount DESC"""))


@test("FR-03", "Customer cash", f"Customer receipts recorded more than {LATE_DAYS} days after the payment date",
      "Late recording of cash can show 'lapping': cash from one customer is used to cover an earlier stolen payment.",
      "For the largest and the most frequent users, compare with the bank or cash book dates.")
def late_receipts(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Customer) AS Customer, CustomerName, substr(PaymentDate,1,10) AS PaymentDate,
               substr(CreatedOn,1,10) AS RecordedOn, CAST(julianday(CreatedOn)-julianday(PaymentDate) AS INT) AS DaysLate,
               Currency, PaymentAmount, PaymentMethod, trim(CashAccount) AS CashAccount, Status, CreatedBy
        FROM ar_payments_and_applications
        WHERE Type='Payment' AND julianday(CreatedOn)-julianday(PaymentDate) > {LATE_DAYS} AND {period('PaymentDate')}
        ORDER BY DaysLate DESC"""))


@test("FR-04", "Customer cash", "Refunds to customers",
      "A refund sends money out. A fake refund can pay money to the wrong person.",
      "Check that each refund has a reason, an approval, and went to the customer's own account.")
def refunds(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, Type, trim(Customer) AS Customer, CustomerName, substr(PaymentDate,1,10) AS Date, Currency,
               PaymentAmount, PaymentMethod, trim(CashAccount) AS CashAccount, Description, Status, CreatedBy
        FROM ar_payments_and_applications
        WHERE Type IN ('Refund','Voided Refund') AND {period('PaymentDate')} ORDER BY PaymentAmount DESC"""))


@test("FR-05", "Customer cash", "Customer receipts on hold",
      "A receipt on hold is recorded but not applied or released. Money can be missing from the bank.",
      "Confirm that the money is in the bank. Release or correct each receipt.")
def receipts_on_hold(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Customer) AS Customer, CustomerName, substr(PaymentDate,1,10) AS Date, Currency,
               PaymentAmount, PaymentMethod, trim(CashAccount) AS CashAccount, Description, CreatedBy
        FROM ar_payments_and_applications WHERE Status='On Hold' AND {period('PaymentDate')}
        ORDER BY PaymentAmount DESC"""))


# ---------------- 2. Payments out ----------------

@test("FR-06", "Payments out", "Suppliers paid who never sent a bill",
      "A fake supplier can be set up to receive payments. A real supplier always sends bills.",
      "Confirm that each supplier exists (registration, address, bank details) and what was bought.")
def paid_no_bills(c):
    return rows_of(c.execute(f"""
        SELECT trim(p.Vendor) AS Vendor, MAX(p.VendorName) AS VendorName, p.Currency, COUNT(*) AS Payments,
               ROUND(SUM(p.PaymentAmount),2) AS Paid, MIN(substr(p.PaymentDate,1,10)) AS FirstPayment,
               MAX(substr(p.PaymentDate,1,10)) AS LastPayment, group_concat(DISTINCT p.CreatedBy) AS MadeBy
        FROM ap_checks_and_payments p
        WHERE p.Type IN ('Payment','Prepayment') AND p.Status <> 'Voided' AND {period('p.PaymentDate')}
          AND trim(p.Vendor) NOT IN (SELECT DISTINCT trim(Vendor) FROM ap_bills_and_adjustments WHERE Type='Bill')
        GROUP BY 1,3 ORDER BY Paid DESC"""))


@test("FR-07", "Payments out", "Payments to staff through supplier payments",
      "Payments to employees set up as suppliers bypass payroll controls.",
      "Check each payment against an approved claim or advance form.")
def staff_payments(c):
    return rows_of(c.execute(f"""
        SELECT 'Payment' AS Source, ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(PaymentDate,1,10) AS Date,
               Currency, PaymentAmount AS Amount, Description, Status, CreatedBy
        FROM ap_checks_and_payments WHERE trim(Vendor) LIKE 'EP%' AND {period('PaymentDate')}
        UNION ALL
        SELECT 'Cash purchase', ReferenceNbr, trim(Vendor), VendorName, substr(Date,1,10), Currency, PaymentAmount,
               Description, Status, CreatedBy
        FROM ap_cash_purchases WHERE trim(Vendor) LIKE 'EP%' AND {period('Date')}
        ORDER BY Amount DESC"""))


@test("FR-08", "Payments out", f"Supplier prepayments still open after {OPEN_PREPAY_DAYS} days",
      "Money was paid before a bill. If no bill comes, the money can be lost or taken.",
      "Ask for the bill or the goods for each one. Recover the money if nothing was received.")
def open_prepayments(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(PaymentDate,1,10) AS PaymentDate,
               CAST(julianday('{END}')-julianday(PaymentDate) AS INT) AS DaysOpenAtPeriodEnd, Currency, PaymentAmount,
               Balance, Description, trim(CashAccount) AS CashAccount, CreatedBy
        FROM ap_checks_and_payments
        WHERE Type='Prepayment' AND Status='Open' AND Balance > 0
          AND julianday('{END}')-julianday(PaymentDate) > {OPEN_PREPAY_DAYS} AND {period('PaymentDate')}
        ORDER BY Balance DESC"""))


@test("FR-09", "Payments out", f"Voided supplier payment, then paid again within {REISSUE_DAYS} days",
      "A payment can be voided in the system after the money left, then paid again: the supplier is paid twice.",
      "Check the bank statement for both payments.")
def reissued(c):
    return rows_of(c.execute(f"""
        SELECT v.ReferenceNbr AS VoidedPayment, n.ReferenceNbr AS NewPayment, trim(v.Vendor) AS Vendor, v.VendorName,
               v.Currency, v.PaymentAmount, substr(v.PaymentDate,1,10) AS VoidedDate, substr(n.PaymentDate,1,10) AS NewDate,
               trim(v.CashAccount) AS VoidedCashAccount, trim(n.CashAccount) AS NewCashAccount,
               CASE WHEN trim(v.CashAccount)<>trim(n.CashAccount) THEN 'Different account' ELSE '' END AS Flag,
               v.CreatedBy AS VoidedBy, n.CreatedBy AS NewBy
        FROM ap_checks_and_payments v JOIN ap_checks_and_payments n
          ON trim(v.Vendor)=trim(n.Vendor) AND v.PaymentAmount=n.PaymentAmount AND v.Currency=n.Currency
         AND n.ReferenceNbr<>v.ReferenceNbr AND n.Status<>'Voided' AND n.Type IN ('Payment','Prepayment')
         AND julianday(n.PaymentDate)-julianday(v.PaymentDate) BETWEEN 0 AND {REISSUE_DAYS}
        WHERE v.Status='Voided' AND v.Type IN ('Payment','Prepayment') AND {period('v.PaymentDate')}
        ORDER BY v.PaymentAmount DESC"""))


@test("FR-10", "Payments out", f"Supplier payments with a round amount (multiple of {ROUND_PAYMENT:,})",
      "Real bills rarely have round totals. Round payments can be advances or made-up amounts.",
      "Match each one to the bills it pays.")
def round_payments(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, Type, trim(Vendor) AS Vendor, VendorName, substr(PaymentDate,1,10) AS Date, Currency,
               PaymentAmount, Description, trim(CashAccount) AS CashAccount, Status, CreatedBy
        FROM ap_checks_and_payments
        WHERE Type IN ('Payment','Prepayment') AND Status<>'Voided' AND PaymentAmount >= {ROUND_PAYMENT}
          AND PaymentAmount % {ROUND_PAYMENT} = 0 AND {period('PaymentDate')}
        ORDER BY PaymentAmount DESC"""))


@test("FR-11", "Payments out", "Descriptions with risk words (personal, loan, advance, shareholder, ...)",
      "These words can show personal or related-party payments booked as business costs.",
      "Read each description. Check that the payment was for the business and approved.")
def keywords(c):
    like = " OR ".join(f"lower(COALESCE(Description,'')) LIKE '%{k}%'" for k in KEYWORDS)
    return rows_of(c.execute(f"""
        SELECT 'Supplier payment' AS Source, ReferenceNbr, trim(Vendor) AS Party, VendorName AS Name,
               substr(PaymentDate,1,10) AS Date, Currency, PaymentAmount AS Amount, Description, CreatedBy
        FROM ap_checks_and_payments WHERE ({like}) AND Status<>'Voided' AND {period('PaymentDate')}
        UNION ALL
        SELECT 'Cash purchase', ReferenceNbr, trim(Vendor), VendorName, substr(Date,1,10), Currency, PaymentAmount,
               Description, CreatedBy
        FROM ap_cash_purchases WHERE ({like}) AND Status<>'Voided' AND {period('Date')}
        UNION ALL
        SELECT 'Supplier bill', ReferenceNbr, trim(Vendor), VendorName, substr(Date,1,10), Currency, Amount,
               Description, CreatedBy
        FROM ap_bills_and_adjustments WHERE ({like}) AND {period('Date')}
        UNION ALL
        SELECT 'Credit memo', ReferenceNbr, trim(Customer), CustomerName, substr(Date,1,10), Currency, Amount,
               Description, CreatedBy
        FROM ar_invoices_and_memos WHERE Type='Credit Memo' AND ({like}) AND {period('Date')}
        ORDER BY Amount DESC"""))


# ---------------- 3. Cash purchases ----------------

@test("FR-12", "Cash purchases", f"Cash purchases of {VAGUE_MIN}+ with a vague or empty description",
      "Without a clear description, nobody can check what was bought.",
      "Ask for the receipt and what was bought. Require clear descriptions.")
def vague(c):
    cols, rows = rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(Date,1,10) AS Date, Currency, PaymentAmount,
               Description, trim(CashAccount) AS CashAccount, Branch, Status, CreatedBy
        FROM ap_cash_purchases
        WHERE Type='Cash Purchase' AND Status NOT IN ('Voided','Rejected') AND PaymentAmount >= {VAGUE_MIN}
          AND {period('Date')} ORDER BY PaymentAmount DESC"""))
    keep = [r for r in rows if (r[6] or "").strip().lower().rstrip(".") in VAGUE_WORDS
            or len((r[6] or "").split()) <= 1]
    return cols, keep


@test("FR-13", "Cash purchases", "Cash purchases made on a weekend",
      "Purchases recorded on weekends get less review.",
      "Check the receipts and why the purchase was needed on that day.")
def weekend_cash(c):
    t = local("CreatedOn")
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(Date,1,10) AS Date, {t} AS RecordedLocal,
               CASE strftime('%w',{t}) WHEN '0' THEN 'Sunday' ELSE 'Saturday' END AS Day,
               Currency, PaymentAmount, Description, Branch, CreatedBy
        FROM ap_cash_purchases
        WHERE Type='Cash Purchase' AND Status NOT IN ('Voided','Rejected') AND strftime('%w',{t}) IN ('0','6')
          AND {period('Date')} ORDER BY PaymentAmount DESC"""))


@test("FR-14", "Cash purchases", f"Cash purchases of {SELF_MIN}+ made and last changed by the same person",
      "No sign that a second person reviewed the purchase.",
      "Check that an approval exists for each one (in Acumatica or on paper).")
def self_handled(c):
    return rows_of(c.execute(f"""
        SELECT ReferenceNbr, trim(Vendor) AS Vendor, VendorName, substr(Date,1,10) AS Date, Currency, PaymentAmount,
               Description, Branch, Status, CreatedBy
        FROM ap_cash_purchases
        WHERE Type='Cash Purchase' AND Status='Closed' AND CreatedBy=LastModifiedBy AND PaymentAmount >= {SELF_MIN}
          AND {period('Date')} ORDER BY PaymentAmount DESC"""))


@test("FR-15", "Cash purchases", "Benford's Law: first digit of USD cash purchases and supplier payments",
      "In real business data, amounts start with 1 about 30% of the time and with 9 about 5%. "
      "Made-up amounts often break this pattern.",
      "If a digit is far above the expected %, look at the amounts that start with it, by person.")
def benford(c):
    out = []
    for label, sql in (("Cash purchases (USD)",
                         f"SELECT PaymentAmount FROM ap_cash_purchases WHERE Currency='USD' AND Type='Cash Purchase' "
                         f"AND Status NOT IN ('Voided','Rejected') AND PaymentAmount >= 10 AND {period('Date')}"),
                        ("Supplier payments (USD)",
                         f"SELECT PaymentAmount FROM ap_checks_and_payments WHERE Currency='USD' "
                         f"AND Type IN ('Payment','Prepayment') AND Status<>'Voided' AND PaymentAmount >= 10 "
                         f"AND {period('PaymentDate')}")):
        vals = [v for (v,) in c.execute(sql) if v]
        digits = Counter(int(str(int(abs(v)))[0]) for v in vals)
        n = sum(digits.values()) or 1
        mad = sum(abs(digits[d] / n - math.log10(1 + 1 / d)) for d in range(1, 10)) / 9
        verdict = ("close conformity" if mad < 0.006 else "acceptable" if mad < 0.012
                   else "marginal" if mad < 0.015 else "nonconformity")
        for d in range(1, 10):
            exp = math.log10(1 + 1 / d)
            act = digits[d] / n
            out.append((label, d, digits[d], round(exp * 100, 1), round(act * 100, 1), round((act - exp) * 100, 1),
                        n, round(mad, 4), verdict))
    return ["Data", "FirstDigit", "Count", "Expected %", "Actual %", "Difference (pts)", "Total amounts",
            "MAD", "Result (Nigrini)"], out


# ---------------- 4. Sales ----------------

def _low_price(c):
    rows = c.execute(f"""
        SELECT OrderType, OrderNbr, LineNbr, substr(Date,1,10) AS Date, Status, trim(Customer) AS Customer, CustomerName,
               InventoryID, Description_2, LineOrderQTY, Amount, Currency, ManualPrice, CreatedBy, DefaultSalesperson
        FROM so_salesorderdetails
        WHERE Behavior='Sales Order' AND Status NOT IN ('Canceled','Rejected') AND Currency='USD'
          AND LineOrderQTY > 0 AND Amount > 0 AND {period('Date')}""").fetchall()
    prices = defaultdict(list)
    for r in rows:
        prices[r[7]].append(r[10] / r[9])
    med = {k: statistics.median(v) for k, v in prices.items() if len(v) >= 10}
    out = []
    for r in rows:
        m = med.get(r[7])
        unit = r[10] / r[9]
        if m and r[12] == 1 and unit < LOW_PRICE * m:
            out.append(list(r[:10]) + [round(unit, 4), round(m, 4), round(unit / m - 1, 3),
                                       round((m - unit) * r[9], 2)] + [r[11], r[13], r[14]])
    cols = ["OrderType", "OrderNbr", "LineNbr", "Date", "Status", "Customer", "CustomerName", "InventoryID", "Item",
            "Qty", "UnitPrice", "MedianUnitPrice", "BelowMedian %", "ValueBelowMedian", "Currency", "CreatedBy",
            "Salesperson"]
    out.sort(key=lambda x: -x[13])
    return cols, out


@test("FR-16", "Sales", f"Manual prices below {int(LOW_PRICE * 100)}% of the item's normal (median) price (USD)",
      "Selling below the normal price to friends or for a kickback is a common way to take value.",
      "Ask who approved each low price. Look for the same customer or user many times.")
def low_price(c):
    return _low_price(c)


@test("FR-17", "Sales", "Low manual prices by user and customer (information)",
      "Shows if low prices go to the same customers or come from the same users.",
      "Discuss the top pairs with the sales manager.")
def low_price_summary(c):
    cols, rows = _low_price(c)
    agg = defaultdict(lambda: [0, 0.0])
    for r in rows:
        k = (r[15], r[16], r[5], r[6])
        agg[k][0] += 1
        agg[k][1] += r[13]
    return (["CreatedBy", "Salesperson", "Customer", "CustomerName", "Lines", "ValueBelowMedian"],
            sorted([list(k) + [v[0], round(v[1], 2)] for k, v in agg.items()], key=lambda x: -x[5]))


# ---------------- 5. Cash banking ----------------

@test("FR-18", "Cash banking", "Cash received and deposits recorded, per cash account per month (information)",
      "Cash that is received but not banked is the simplest way to take money.",
      "For each cash account, compare with the bank statements and the cash count. Explain big differences.")
def cash_vs_deposits(c):
    return rows_of(c.execute(f"""
        WITH r AS (SELECT trim(CashAccount) a, Currency cur, substr(PaymentDate,1,7) m, SUM(PaymentAmount) v, COUNT(*) n
                   FROM ar_payments_and_applications
                   WHERE Type='Payment' AND Status NOT IN ('Voided') AND PaymentMethod='CASH' AND {period('PaymentDate')}
                   GROUP BY 1,2,3),
             d AS (SELECT trim(CashAccount) a, Currency cur, substr(DepositDate,1,7) m, SUM(TotalAmount) v, COUNT(*) n
                   FROM ca_bank_deposits WHERE Status <> 'Voided' AND {period('DepositDate')} GROUP BY 1,2,3),
             k AS (SELECT a, cur, m FROM r UNION SELECT a, cur, m FROM d)
        SELECT k.a AS CashAccount, k.cur AS Currency, k.m AS Month, ROUND(COALESCE(r.v,0),2) AS CashReceipts,
               COALESCE(r.n,0) AS Receipts, ROUND(COALESCE(d.v,0),2) AS DepositsRecorded, COALESCE(d.n,0) AS Deposits
        FROM k LEFT JOIN r ON r.a=k.a AND r.cur=k.cur AND r.m=k.m LEFT JOIN d ON d.a=k.a AND d.cur=k.cur AND d.m=k.m
        ORDER BY k.a, k.cur, k.m"""))


# ---------------- run ----------------

INFO = {"FR-15", "FR-17", "FR-18", "FR-99"}
PERSON_COL = {"FR-01": "CreatedBy", "FR-02": "CreatedBy", "FR-03": "CreatedBy", "FR-04": "CreatedBy",
              "FR-05": "CreatedBy", "FR-07": "CreatedBy", "FR-08": "CreatedBy", "FR-09": "NewBy",
              "FR-10": "CreatedBy", "FR-11": "CreatedBy", "FR-12": "CreatedBy", "FR-13": "CreatedBy",
              "FR-14": "CreatedBy", "FR-16": "CreatedBy"}


def by_person(results):
    """FR-99: number of red flags per person, per test."""
    ids = [t["id"] for t in results if t["id"] in PERSON_COL]
    counts = defaultdict(Counter)
    for t in results:
        col = PERSON_COL.get(t["id"])
        if col and col in t["cols"]:
            i = t["cols"].index(col)
            for r in t["rows"]:
                counts[r[i] or "(blank)"][t["id"]] += 1
    rows = [[p, sum(c.values())] + [c.get(i, 0) for i in ids] for p, c in counts.items()]
    rows.sort(key=lambda x: -x[1])
    return ["Person", "Total flags"] + ids, rows


def main():
    if not os.path.exists(DB_PATH):
        raise SystemExit(f"No database at {DB_PATH}. Run: python sync.py")
    os.makedirs(OUT, exist_ok=True)
    print("Loading data ...")
    con = load_numeric_copy(DB_PATH)
    results = []
    for t in TESTS:
        try:
            cols, rows = t["fn"](con)
            print(f"{t['id']:6} {len(rows):>8,}  {t['title']}")
        except sqlite3.OperationalError as e:
            cols, rows = ["Not run"], [[f"Test not run: {e}"]]
            print(f"{t['id']:6}  NOT RUN  {t['title']}  ({e})")
        results.append(dict(t, cols=cols, rows=[list(r) for r in rows]))
    pc, pr = by_person(results)
    results.append(dict(id="FR-99", area="People", title="Red flags per person (information)",
                        why="Shows whose transactions appear most in the red-flag tests. Many flags can simply mean "
                            "a person processes many transactions. Compare with their workload.",
                        follow_up="Review the top names' flagged items first. Do not draw conclusions from counts only.",
                        cols=pc, rows=pr))
    con.close()
    notes = [
        "IMPORTANT: a red flag is a reason to look closer, not proof that anyone did something wrong. "
        "Many flags will have normal explanations.",
        "This file contains staff names. Share it only with the people who must see it (for example the CEO and IT/finance lead).",
        f"Period: {START} to {END}. Amounts are in the document currency. Do not add different currencies.",
        "Approval limits are not known, so the 'just below the limit' test is not included. Add it when the limits are known.",
        "No scrap is recorded in production, so there is no scrap test.",
        "98% of credit memos have no link to the original invoice (FR-02). That is a control weakness by itself.",
        f"Weekend test uses CreatedOn + {TIME_OFFSET} hours (CREATED_ON_UTC_OFFSET in .env).",
        "Thresholds (days, amounts, %) are at the top of fraud_red_flags.py.",
        "Benford's Law MAD limits (Nigrini): <0.006 close, <0.012 acceptable, <0.015 marginal, else nonconformity.",
    ]
    path = os.path.join(OUT, "fraud_red_flags.xlsx")
    write_workbook(results, notes, path, title="Supreme Brands – red flags for possible loss of money", info=INFO)
    print(f"\nExcel: {path}")


if __name__ == "__main__":
    main()

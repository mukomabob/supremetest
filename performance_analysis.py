"""Quarterly business performance from Acumatica data, one metric per sheet.

Usage:
    python performance_analysis.py

Output: reports/performance_analysis.xlsx (not sent to GitHub)

Period and currency come from .env (defaults below). Amounts are in the document currency.
The main measure uses BASE_CURRENCY documents only; other currencies are on their own sheet.
"""
import os
import datetime as dt

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from audit_tests import load_numeric_copy  # same numeric fix as the audit tests

DB_PATH = os.getenv("DB_PATH", "data/supreme.db")
START = os.getenv("PERIOD_START", "2023-07-01")
END = os.getenv("PERIOD_END", "2026-09-30")
CUR = os.getenv("BASE_CURRENCY", "USD")
OUT = "reports"
TOP_PRODUCTS = 20
# Document pairs that cancel each other (invoice + credit memo, bill + debit adjustment; same party and amount,
# reversal within 31 days), in BASE_CURRENCY and this large or larger, are left out and listed on "Excluded".
PAIR_MIN = float(os.getenv("REVERSED_PAIR_MIN", "100000"))


def qtr(col):
    """SQL: '2024-Q3' from a date column."""
    return f"strftime('%Y',{col}) || '-Q' || ((CAST(strftime('%m',{col}) AS INT)+2)/3)"


def in_period(col):
    return f"substr({col},1,10) BETWEEN '{START}' AND '{END}'"


def quarters():
    s, e = dt.date.fromisoformat(START), dt.date.fromisoformat(END)
    out, y, q = [], s.year, (s.month - 1) // 3 + 1
    while (y, q) <= (e.year, (e.month - 1) // 3 + 1):
        out.append(f"{y}-Q{q}")
        y, q = (y + 1, 1) if q == 4 else (y, q + 1)
    return out


def by_quarter(con, sql):
    """Run SQL that returns (quarter, values...) and fill missing quarters with zeros."""
    cur = con.execute(sql)
    cols = [d[0] for d in cur.description]
    data = {r[0]: list(r[1:]) for r in cur.fetchall()}
    return cols, [[q] + data.get(q, [0] * (len(cols) - 1)) for q in quarters()]


# ---------------- metrics ----------------

def metrics(con):
    Q = quarters()
    m = {}

    # Large document pairs that cancel each other (typing errors, go-live balances), base currency only:
    # customer invoice + credit memo, and supplier bill + debit adjustment.
    con.execute("DROP TABLE IF EXISTS excluded_docs")
    con.execute(f"""
        CREATE TEMP TABLE excluded_docs AS
        SELECT 'Customer' AS Side, i.ReferenceNbr AS DocNbr, m.ReferenceNbr AS ReversalNbr,
               trim(i.Customer) AS Party, i.CustomerName AS PartyName, substr(i.Date,1,10) AS DocDate,
               substr(m.Date,1,10) AS ReversalDate, i.Currency, i.Amount, i.Description, i.CreatedBy
        FROM ar_invoices_and_memos i JOIN ar_invoices_and_memos m
          ON trim(i.Customer)=trim(m.Customer) AND i.Amount=m.Amount AND i.Currency=m.Currency
         AND m.Type='Credit Memo' AND julianday(m.Date) - julianday(i.Date) BETWEEN 0 AND 31
        WHERE i.Type='Invoice' AND i.Currency='{CUR}' AND i.Amount >= {PAIR_MIN}
        UNION ALL
        SELECT 'Supplier', b.ReferenceNbr, d.ReferenceNbr, trim(b.Vendor), b.VendorName, substr(b.Date,1,10),
               substr(d.Date,1,10), b.Currency, b.Amount, b.Description, b.CreatedBy
        FROM ap_bills_and_adjustments b JOIN ap_bills_and_adjustments d
          ON trim(b.Vendor)=trim(d.Vendor) AND b.Amount=d.Amount AND b.Currency=d.Currency
         AND d.Type='Debit Adj.' AND julianday(d.Date) - julianday(b.Date) BETWEEN 0 AND 31
        WHERE b.Type='Bill' AND b.Currency='{CUR}' AND b.Amount >= {PAIR_MIN}""")
    m["excluded"] = (["Side", "DocNbr", "ReversalNbr", "Party", "PartyName", "DocDate", "ReversalDate",
                      "Currency", "Amount", "Description", "CreatedBy"],
                     con.execute("SELECT * FROM excluded_docs ORDER BY Amount DESC").fetchall())
    ex = ("AND ReferenceNbr NOT IN (SELECT DocNbr FROM excluded_docs "
          "UNION SELECT ReversalNbr FROM excluded_docs)")

    # Sales: invoices, debit memos, credit memos
    m["sales"] = by_quarter(con, f"""
        SELECT {qtr('Date')} AS Quarter,
               SUM(CASE WHEN Type='Invoice' THEN Amount ELSE 0 END) AS Invoices,
               SUM(CASE WHEN Type='Debit Memo' THEN Amount ELSE 0 END) AS DebitMemos,
               SUM(CASE WHEN Type='Credit Memo' THEN Amount ELSE 0 END) AS CreditMemos,
               SUM(CASE WHEN Type='Invoice' THEN 1 ELSE 0 END) AS InvoiceCount,
               SUM(CASE WHEN Type='Credit Memo' THEN 1 ELSE 0 END) AS CreditMemoCount,
               COUNT(DISTINCT CASE WHEN Type='Invoice' THEN trim(Customer) END) AS ActiveCustomers
        FROM ar_invoices_and_memos
        WHERE Currency='{CUR}' AND Status NOT IN ('Canceled','Rejected') AND {in_period('Date')} {ex}
        GROUP BY 1""")

    # Top 10 customers' share per quarter
    rows = con.execute(f"""
        SELECT {qtr('Date')} q, trim(Customer) c,
               SUM(CASE WHEN Type='Credit Memo' THEN -Amount ELSE Amount END) s
        FROM ar_invoices_and_memos
        WHERE Currency='{CUR}' AND Status NOT IN ('Canceled','Rejected') AND {in_period('Date')} {ex}
        GROUP BY 1,2""").fetchall()
    per_q = {}
    for q, c, s in rows:
        per_q.setdefault(q, []).append(s or 0)
    m["concentration"] = (["Quarter", "Top10Sales", "AllSales", "Customers"],
                          [[q, sum(sorted(per_q.get(q, []), reverse=True)[:10]), sum(per_q.get(q, [])),
                            len(per_q.get(q, []))] for q in Q])

    # Customer receipts
    m["receipts"] = by_quarter(con, f"""
        SELECT {qtr('PaymentDate')} AS Quarter,
               SUM(CASE WHEN Type IN ('Payment','Prepayment') THEN PaymentAmount ELSE 0 END) AS Receipts,
               SUM(CASE WHEN Type='Refund' THEN PaymentAmount ELSE 0 END) AS Refunds,
               SUM(CASE WHEN Type IN ('Payment','Prepayment') THEN 1 ELSE 0 END) AS ReceiptCount
        FROM ar_payments_and_applications
        WHERE Currency='{CUR}' AND Status NOT IN ('Voided') AND Type IN ('Payment','Prepayment','Refund')
              AND {in_period('PaymentDate')}
        GROUP BY 1""")

    # Sales orders (volume)
    m["orders"] = by_quarter(con, f"""
        SELECT {qtr('Date')} AS Quarter, COUNT(DISTINCT OrderType||OrderNbr) AS Orders, COUNT(*) AS Lines,
               SUM(LineOrderQTY) AS QtyOrdered, SUM(Amount) AS OrderValue
        FROM so_salesorderdetails
        WHERE Behavior='Sales Order' AND Status NOT IN ('Canceled','Rejected') AND Currency='{CUR}'
              AND {in_period('Date')}
        GROUP BY 1""")

    # Top products over the whole period, with value per quarter
    top = con.execute(f"""
        SELECT InventoryID, MAX(Description_2), SUM(Amount) v, SUM(LineOrderQTY)
        FROM so_salesorderdetails
        WHERE Behavior='Sales Order' AND Status NOT IN ('Canceled','Rejected') AND Currency='{CUR}'
              AND {in_period('Date')}
        GROUP BY 1 ORDER BY v DESC LIMIT {TOP_PRODUCTS}""").fetchall()
    pq = {}
    for inv, q, v in con.execute(f"""
        SELECT InventoryID, {qtr('Date')}, SUM(Amount) FROM so_salesorderdetails
        WHERE Behavior='Sales Order' AND Status NOT IN ('Canceled','Rejected') AND Currency='{CUR}'
              AND {in_period('Date')} GROUP BY 1,2"""):
        pq[(inv, q)] = v
    m["products"] = (["InventoryID", "Product", "TotalValue", "TotalQty"] + Q,
                     [[inv, d, v, qty] + [pq.get((inv, q), 0) for q in Q] for inv, d, v, qty in top])

    # Purchases: bills only (adjustments are also used for currency moves and payments) and cash purchases
    _, bills = by_quarter(con, f"""
        SELECT {qtr('Date')}, SUM(Amount), COUNT(DISTINCT trim(Vendor))
        FROM ap_bills_and_adjustments
        WHERE Type='Bill' AND Currency='{CUR}' AND Status NOT IN ('On Hold','Reserved') AND {in_period('Date')} {ex}
        GROUP BY 1""")
    _, cash = by_quarter(con, f"""
        SELECT {qtr('Date')},
               SUM(CASE WHEN Type='Cash Return' THEN -PaymentAmount ELSE PaymentAmount END), COUNT(*)
        FROM ap_cash_purchases
        WHERE Currency='{CUR}' AND Type IN ('Cash Purchase','Cash Return')
              AND Status NOT IN ('Voided','Rejected','On Hold','Pending Approval') AND {in_period('Date')}
        GROUP BY 1""")
    m["purchases"] = (["Quarter", "SupplierBills", "Suppliers", "CashPurchases", "CashPurchaseCount"],
                      [[b[0], b[1], b[2], c[1], c[2]] for b, c in zip(bills, cash)])

    # Production: one row per production order (the inquiry repeats orders per transaction)
    m["production"] = by_quarter(con, f"""
        WITH o AS (SELECT ProductionNbr, MIN(StartDate) d, MAX(QtytoProduce) qp, MAX(CompletedQty) qc,
                          MAX(ScrappedQty) qs, MAX(Canceled) cx
                   FROM am_production_orders_details GROUP BY ProductionNbr)
        SELECT {qtr('d')} AS Quarter, COUNT(*) AS Orders, SUM(qp) AS QtyToProduce, SUM(qc) AS QtyCompleted,
               SUM(qs) AS QtyScrapped
        FROM o WHERE cx=0 AND {in_period('d')} GROUP BY 1""")

    # Other currencies: net sales per currency
    others = [r[0] for r in con.execute(
        f"SELECT DISTINCT Currency FROM ar_invoices_and_memos WHERE Currency<>'{CUR}' ORDER BY 1")]
    data = {}
    for q, c, s in con.execute(f"""
        SELECT {qtr('Date')}, Currency, SUM(CASE WHEN Type='Credit Memo' THEN -Amount ELSE Amount END)
        FROM ar_invoices_and_memos
        WHERE Currency<>'{CUR}' AND Status NOT IN ('Canceled','Rejected') AND {in_period('Date')} {ex}
        GROUP BY 1,2"""):
        data[(q, c)] = s
    m["other_currencies"] = (["Quarter"] + [f"NetSales {c}" for c in others],
                             [[q] + [data.get((q, c), 0) for c in others] for q in Q])
    return m


# ---------------- workbook ----------------

def build(m, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter as L
    from openpyxl.chart import LineChart, BarChart, Reference

    F, B = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    FILL = PatternFill("solid", start_color="D9E1F2")
    MONEY, NUM, PCT = "#,##0;(#,##0);-", "#,##0;(#,##0);-", "0.0%;(0.0%);-"
    HR = 4  # header row
    n = len(quarters())

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    def sheet(name, title, note, headers, rows, formats, formulas=(), growth=(), chart=None):
        """Write one metric sheet. formulas: list of (header, template with {r}); growth: column letters."""
        ws = wb.create_sheet(name)
        ws["A1"], ws["A2"] = title, note
        ws["A1"].font = Font(name="Arial", size=12, bold=True)
        ws["A2"].font = Font(name="Arial", size=10, italic=True)
        heads = list(headers) + [h for h, _ in formulas]
        for col in growth:
            heads += [f"{ws_header(headers, formulas, col)} QoQ %", f"{ws_header(headers, formulas, col)} YoY %"]
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=HR, column=i, value=h)
            c.font, c.fill = B, FILL
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for ri, row in enumerate(rows, HR + 1):
            for ci, v in enumerate(row, 1):
                c = ws.cell(row=ri, column=ci, value=v)
                c.font = F
                if ci > 1 and isinstance(v, (int, float)):
                    c.number_format = formats.get(headers[ci - 1], MONEY)
            ci = len(headers)
            for h, tmpl in formulas:
                ci += 1
                c = ws.cell(row=ri, column=ci, value=tmpl.format(r=ri))
                c.font, c.number_format = F, formats.get(h, MONEY)
            for col in growth:
                p1, p4 = ri - 1, ri - 4
                ci += 1
                c = ws.cell(row=ri, column=ci,
                            value=f'=IF(OR({col}{p1}="",{col}{p1}=0),"",{col}{ri}/{col}{p1}-1)' if p1 > HR else "")
                c.font, c.number_format = F, PCT
                ci += 1
                c = ws.cell(row=ri, column=ci,
                            value=f'=IF(OR({col}{p4}="",{col}{p4}=0),"",{col}{ri}/{col}{p4}-1)' if p4 > HR else "")
                c.font, c.number_format = F, PCT
        for i in range(1, len(heads) + 1):
            ws.column_dimensions[L(i)].width = 14 if i > 1 else 12
        ws.row_dimensions[HR].height = 30
        ws.freeze_panes = ws.cell(row=HR + 1, column=2)
        if chart:
            col, ctitle, kind = chart
            ch = BarChart() if kind == "bar" else LineChart()
            ch.title, ch.height, ch.width = ctitle, 8, 22
            ch.add_data(Reference(ws, min_col=col, min_row=HR, max_row=HR + len(rows)), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=1, min_row=HR + 1, max_row=HR + len(rows)))
            ch.roundedCorners = False
            ch.legend = None
            ch.x_axis.delete = False
            ch.y_axis.delete = False
            ch.title.overlay = False
            ws.add_chart(ch, f"A{HR + len(rows) + 3}")
        return ws

    def ws_header(headers, formulas, col):
        idx = ord(col) - 65
        allh = list(headers) + [h for h, _ in formulas]
        return allh[idx] if idx < len(allh) else col

    c = f"({CUR})"
    h, r = m["sales"]
    rows = [[x[0], x[1], x[2], x[3], x[4], x[5], x[6]] for x in r]
    sheet("Net sales", f"Net sales per quarter {c}",
          "Net sales = invoices + debit memos - credit memos. Canceled documents are not included.",
          ["Quarter", "Invoices", "Debit memos", "Credit memos", "Invoice count", "Credit memo count", "Active customers"],
          rows, {"Invoice count": NUM, "Credit memo count": NUM, "Active customers": NUM},
          formulas=[("Net sales", "=B{r}+C{r}-D{r}")], growth=["H"], chart=(8, f"Net sales {c}", "bar"))
    sheet("Invoices", f"Number of invoices and average invoice value {c}", "Average = invoice value / invoice count.",
          ["Quarter", "Invoice count", "Invoice value"], [[x[0], x[4], x[1]] for x in r],
          {"Invoice count": NUM}, formulas=[("Average invoice", '=IF(B{r}=0,"",C{r}/B{r})')],
          growth=["B", "D"], chart=(4, f"Average invoice value {c}", "line"))
    sheet("Credit memos", f"Credit memos as % of invoiced sales {c}",
          "A rising % can show more returns, price disputes or invoicing errors.",
          ["Quarter", "Invoices", "Credit memos", "Credit memo count"], [[x[0], x[1], x[3], x[5]] for x in r],
          {"Credit memo count": NUM}, formulas=[("Credit memos %", '=IF(B{r}=0,"",C{r}/B{r})')],
          chart=(5, "Credit memos as % of invoices", "line"), growth=[])
    for ws_name in ("Credit memos",):
        for row in range(HR + 1, HR + n + 1):
            wb[ws_name].cell(row=row, column=5).number_format = PCT
    sheet("Active customers", "Active customers per quarter (at least one invoice)",
          f"Customers with at least one {CUR} invoice in the quarter.",
          ["Quarter", "Active customers"], [[x[0], x[6]] for x in r], {"Active customers": NUM},
          growth=["B"], chart=(2, "Active customers", "bar"))

    h, r = m["concentration"]
    sheet("Concentration", f"Sales from the top 10 customers {c}",
          "A high % means that the business depends on a small number of customers.",
          ["Quarter", "Top 10 customers", "All customers", "Customers"], r, {"Customers": NUM},
          formulas=[("Top 10 share %", '=IF(C{r}=0,"",B{r}/C{r})')], chart=(5, "Top 10 customers share", "line"))
    for row in range(HR + 1, HR + n + 1):
        wb["Concentration"].cell(row=row, column=5).number_format = PCT

    rh, rr = m["receipts"]
    sales_net = {x[0]: (x[1] or 0) + (x[2] or 0) - (x[3] or 0) for x in m["sales"][1]}
    sheet("Collections", f"Customer receipts and collection rate {c}",
          "Collection rate = receipts / net sales of the same quarter. Over 100% means old debts were collected.",
          ["Quarter", "Receipts", "Refunds", "Receipt count", "Net sales"],
          [[x[0], x[1], x[2], x[3], sales_net.get(x[0], 0)] for x in rr], {"Receipt count": NUM},
          formulas=[("Collection rate %", '=IF(E{r}=0,"",(B{r}-C{r})/E{r})')],
          growth=["B"], chart=(2, f"Customer receipts {c}", "bar"))
    for row in range(HR + 1, HR + n + 1):
        wb["Collections"].cell(row=row, column=6).number_format = PCT

    h, r = m["orders"]
    sheet("Sales volume", f"Sales orders and quantity ordered {c}",
          "From sales order lines (not canceled or rejected). Quantity is in each item's unit (mostly EACH).",
          ["Quarter", "Orders", "Order lines", "Qty ordered", "Order value"], r,
          {"Orders": NUM, "Order lines": NUM, "Qty ordered": NUM},
          growth=["D", "E"], chart=(4, "Quantity ordered", "bar"))

    h, r = m["products"]
    ws = sheet("Top products", f"Top {TOP_PRODUCTS} products by sales order value {c}",
               "Total for the whole period, then the value per quarter.",
               ["InventoryID", "Product", "Total value", "Total qty"] + quarters(), r, {"Total qty": NUM})
    ws.column_dimensions["B"].width = 34

    h, r = m["purchases"]
    sheet("Purchases", f"Supplier bills and cash purchases {c}",
          "Supplier bills only. Debit and credit adjustments are left out, because they are also used to move balances between USD and ZAR and to record payments. Cash purchases = cash purchases - cash returns.",
          ["Quarter", "Supplier bills", "Suppliers billed", "Cash purchases", "Cash purchase count"], r,
          {"Suppliers billed": NUM, "Cash purchase count": NUM},
          formulas=[("Total purchases", "=B{r}+D{r}")], growth=["F"], chart=(6, f"Total purchases {c}", "bar"))

    h, r = m["production"]
    sheet("Production", "Production orders and quantities",
          "One row per production order, by start date. Canceled orders are not included.",
          ["Quarter", "Orders", "Qty to produce", "Qty completed", "Qty scrapped"], r,
          {"Orders": NUM, "Qty to produce": NUM, "Qty completed": NUM, "Qty scrapped": NUM},
          formulas=[("Completion %", '=IF(C{r}=0,"",D{r}/C{r})')], growth=["D"],
          chart=(4, "Quantity completed", "bar"))
    for row in range(HR + 1, HR + n + 1):
        wb["Production"].cell(row=row, column=6).number_format = PCT

    h, r = m["other_currencies"]
    sheet("Other currencies", "Net sales in other currencies",
          "Not comparable over time: ZWL was replaced by ZWG in April 2024, and the rates changed a lot. Not converted.",
          [x.replace("NetSales ", "Net sales ") for x in h], r, {})

    h, r = m["excluded"]
    ws = sheet("Excluded", f"Large document pairs that cancel each other, left out ({CUR} >= {PAIR_MIN:,.0f})",
               "Customer invoice + credit memo, or supplier bill + debit adjustment: same party and amount, reversal "
               "within 31 days. Totals do not change, but averages and % do. Check each pair (typing error? go-live?).",
               h, [list(x) for x in r] or [["None"] + [""] * (len(h) - 1)], {})
    for col, w in zip("ABCDEFGHIJK", (10, 12, 13, 10, 30, 12, 13, 9, 16, 40, 22)):
        ws.column_dimensions[col].width = w

    # Summary: one line per quarter with the main numbers, linked to the metric sheets
    summary["A1"] = f"Supreme Brands – quarterly performance {START} to {END} ({CUR})"
    summary["A1"].font = Font(name="Arial", size=14, bold=True)
    summary["A2"] = "Values come from the metric sheets. See the Notes sheet before you use these numbers."
    summary["A2"].font = Font(name="Arial", size=10, italic=True)
    cols = [("Quarter", "'Net sales'!A{r}", None), ("Net sales", "'Net sales'!H{r}", MONEY),
            ("Net sales YoY %", "'Net sales'!J{r}", PCT), ("Invoices", "Invoices!B{r}", NUM),
            ("Avg invoice", "Invoices!D{r}", MONEY), ("Active customers", "'Active customers'!B{r}", NUM),
            ("Top 10 share %", "Concentration!E{r}", PCT), ("Credit memos %", "'Credit memos'!E{r}", PCT),
            ("Receipts", "Collections!B{r}", MONEY), ("Collection rate %", "Collections!F{r}", PCT),
            ("Qty ordered", "'Sales volume'!D{r}", NUM), ("Total purchases", "Purchases!F{r}", MONEY),
            ("Qty produced", "Production!D{r}", NUM)]
    for i, (hd, _, _) in enumerate(cols, 1):
        cc = summary.cell(row=HR, column=i, value=hd)
        cc.font, cc.fill = B, FILL
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        summary.column_dimensions[L(i)].width = 13
    for k in range(n):
        ri = HR + 1 + k
        for i, (_, ref, fmt) in enumerate(cols, 1):
            cc = summary.cell(row=ri, column=i, value="=" + ref.format(r=ri))
            cc.font = F
            if fmt:
                cc.number_format = fmt
    summary.row_dimensions[HR].height = 30
    summary.freeze_panes = "B5"

    notes = wb.create_sheet("Notes", 1)
    lines = [
        f"Period: {START} to {END} (full quarters). Main currency: {CUR} documents only.",
        "Acumatica data starts in mid-2023 and the number of documents grows fast in the first quarters. "
        "Part of the early growth is probably the move of business into Acumatica (go-live), not real growth. "
        "Compare later quarters with care, and use year-on-year (YoY) more than quarter-on-quarter (QoQ).",
        f"Large invoice and credit memo pairs that cancel each other (>= {PAIR_MIN:,.0f}) are left out of the sales "
        "metrics and listed on the Excluded sheet. Examples: a $155,250,000 invoice (March 2024) and a $19,328,050 "
        "supplier bill (December 2024), each reversed soon after. These look like typing errors.",
        "Purchases use supplier bills only. Debit and credit adjustments are left out, because the team also uses "
        "them to move supplier balances between USD and ZAR accounts and to record payments.",
        "Production data starts in January 2025 (manufacturing module go-live), so earlier quarters show zero.",
        "Other currencies (ZWL, ZWG, ZAR) are not converted to USD. ZWL was replaced by ZWG in April 2024.",
        "Gross profit and net profit are not shown: account-level data is available for November 2026 only.",
        "Sales use AR invoices by document date. Volumes use sales order lines by order date, so the two can differ.",
        "Collection rate compares receipts with sales of the same quarter. It is an indicator, not a debtors analysis.",
        "Production uses one row per production order (the inquiry repeats orders), by start date.",
        "QoQ % = change from the previous quarter. YoY % = change from the same quarter last year.",
        f"Made {dt.datetime.now():%d %b %Y %H:%M} by performance_analysis.py.",
    ]
    notes["A1"] = "Notes"
    notes["A1"].font = Font(name="Arial", size=12, bold=True)
    for i, t in enumerate(lines, 3):
        cc = notes.cell(row=i, column=1, value=t)
        cc.font, cc.alignment = F, Alignment(wrap_text=True, vertical="top")
    notes.column_dimensions["A"].width = 130
    wb.save(path)


def main():
    if not os.path.exists(DB_PATH):
        raise SystemExit(f"No database at {DB_PATH}. Run: python sync.py")
    os.makedirs(OUT, exist_ok=True)
    print("Loading data ...")
    con = load_numeric_copy(DB_PATH)
    m = metrics(con)
    con.close()
    path = os.path.join(OUT, "performance_analysis.xlsx")
    build(m, path)
    print(f"Excel: {path}  ({len(quarters())} quarters, {START} to {END}, {CUR})")


if __name__ == "__main__":
    main()

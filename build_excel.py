"""
Build a self-contained Excel workbook from your CIBC CSV exports.

Run:
    python3 build_excel.py

This reads every CSV in credit/ and debit/, classifies each transaction, and
writes finance.xlsx. After that, EVERYTHING happens inside Excel -- no Python,
no internet, no AI:

  - Dashboard   : headline numbers + charts (spending by category, by month,
                  income). They recompute automatically via formulas.
  - Transactions: every cleaned transaction. The Category column is a formula
                  that looks up the merchant/name on the sheets below.
  - Merchants   : each unique spending merchant. Pick a Category from the
                  dropdown -- it applies to every matching transaction.
  - Names       : each unique transfer name (allowance, split bills, salary...).

Re-run this script whenever you add new statements; your previous dropdown
choices are read back from the existing finance.xlsx and preserved.
"""

import glob
import io
import os

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import clean_data as cd

XLSX = os.path.join(cd.BASE, "finance.xlsx")

# Category options offered in the dropdowns (Uncategorized is always allowed).
SPENDING_OPTIONS = cd.SPENDING_CATEGORIES + ["Uncategorized"]
TRANSFER_OPTIONS = cd.TRANSFER_CATEGORIES + ["Uncategorized"]

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=20)
LABEL_FONT = Font(bold=True)
MONEY_FMT = "$#,##0"
MONEY_FMT2 = "$#,##0.00"


# ── ingest (reuses clean_data, but writes nothing) ─────────────────
DEDUP_KEY = ["date", "description", "debit", "credit"]


def _finalize(df):
    """Normalize date and derive merchant / flow / amount on a raw frame."""
    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df.dropna(subset=["date"])
    df["merchant"] = df["description"].apply(cd.normalize_merchant)
    df["flow"] = df["description"].apply(cd.classify_flow)
    df["amount"] = df["debit"] + df["credit"]
    return df.sort_values("date").reset_index(drop=True)


def ingest_frames(file_specs):
    """file_specs: list of (path-or-file-like, source) -> classified DataFrame."""
    frames = [cd.read_bank_csv(src, source) for src, source in file_specs]
    if not frames:
        return pd.DataFrame()
    return _finalize(pd.concat(frames, ignore_index=True))


def ingest():
    """CLI entry: read every CSV in credit/ and debit/ on disk."""
    specs = [(f, source)
             for source in ("credit", "debit")
             for f in glob.glob(os.path.join(cd.BASE, source, "*.csv"))]
    return ingest_frames(specs)


def merge_prior(df_new, prior_raw):
    """Fold history from a previous workbook into the freshly imported rows.

    A previous transaction is kept only when its (date, description, debit,
    credit) does NOT appear in the new files. So re-exporting an overlapping
    date range never doubles anything, while genuine same-day repeats within a
    single export are preserved.
    """
    if prior_raw is None or prior_raw.empty:
        return df_new
    p = prior_raw.copy()
    p["date"] = pd.to_datetime(p["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    p = p.dropna(subset=["date"])
    if df_new is None or df_new.empty:
        return _finalize(p)
    new_keys = set(df_new[DEDUP_KEY].itertuples(index=False, name=None))
    keep = ~p[DEDUP_KEY].apply(tuple, axis=1).isin(new_keys)
    p_keep = p[keep]
    if p_keep.empty:
        return df_new
    combined = pd.concat([df_new, _finalize(p_keep)], ignore_index=True)
    return combined.sort_values("date").reset_index(drop=True)


def seed_table(df, flow, defaults):
    """One row per unique merchant/name with a starting Category guess."""
    sub = df[df["flow"] == flow].copy()
    mapping = {k.upper(): v for k, v in defaults.items()}
    sub["label"] = sub["description"].apply(lambda d: cd.apply_label(d, mapping))
    rows = []
    for name, g in sub.groupby("merchant"):
        labelled = g["label"][g["label"] != "Uncategorized"]
        cat = labelled.mode().iloc[0] if not labelled.empty else "Uncategorized"
        rows.append({
            "name": name,
            "category": cat,
            "count": len(g),
            "total": round(float(g["amount"].sum()), 2),
            "example": g["description"].iloc[0],
        })
    if not rows:
        return pd.DataFrame(columns=["name", "category", "count", "total", "example"])
    return (pd.DataFrame(rows)
            .sort_values("total", ascending=False)
            .reset_index(drop=True))


def _labels_from_wb(wb):
    """Pull {name: category} from the Merchants and Names sheets of a workbook."""
    result = {}
    for sheet in ("Merchants", "Names"):
        d = {}
        if sheet in wb.sheetnames:
            for k, v in wb[sheet].iter_rows(min_row=2, max_col=2, values_only=True):
                if k and v and str(v).strip().lower() not in ("", "none"):
                    d[str(k).strip()] = str(v).strip()
        result[sheet] = d
    return result["Merchants"], result["Names"]


def _prior_transactions(wb):
    """Recover the raw transactions stored on a previous Transactions sheet."""
    cols = ["date", "source", "description", "debit", "credit"]
    if "Transactions" not in wb.sheetnames:
        return pd.DataFrame(columns=cols)
    rows = []
    for r in wb["Transactions"].iter_rows(min_row=2, values_only=True):
        if not r or r[0] in (None, ""):
            continue
        rows.append({
            "date": str(r[0]),
            "source": str(r[1]) if len(r) > 1 and r[1] is not None else "",
            "description": str(r[4]) if len(r) > 4 and r[4] is not None else "",
            "debit": float(r[5]) if len(r) > 5 and r[5] not in (None, "") else 0.0,
            "credit": float(r[6]) if len(r) > 6 and r[6] not in (None, "") else 0.0,
        })
    return pd.DataFrame(rows, columns=cols)


def read_prior(wb):
    """Return (merchant_labels, name_labels, prior_transactions) from a workbook."""
    m, n = _labels_from_wb(wb)
    return m, n, _prior_transactions(wb)


# ── worksheet builders ─────────────────────────────────────────────
def _write_header(ws, row, headers):
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=text)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT


def _text_categories(chart, labels):
    """Make the category axis a STRING reference with cached labels.

    openpyxl's set_categories writes a numeric reference (numRef); when the
    category cells hold text (category names, month labels) Excel treats that
    as unreadable content and strips the chart on open. A strRef with a cache
    is what Excel expects for text categories."""
    ref = chart.series[0].cat.numRef.f
    cache = StrData(pt=[StrVal(idx=i, v=str(v)) for i, v in enumerate(labels)],
                    ptCount=len(labels))
    chart.series[0].cat = AxDataSource(strRef=StrRef(f=ref, strCache=cache))


def build_lists(ws):
    ws["A1"] = "SpendingOptions"
    for i, opt in enumerate(SPENDING_OPTIONS, start=2):
        ws.cell(row=i, column=1, value=opt)
    ws["B1"] = "TransferOptions"
    for i, opt in enumerate(TRANSFER_OPTIONS, start=2):
        ws.cell(row=i, column=2, value=opt)
    ws.sheet_state = "hidden"


def build_mapping_sheet(ws, name_header, table, options, list_col):
    _write_header(ws, 1, [name_header, "Category", "Count", "Total", "Example"])
    for i, r in enumerate(table.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=r.name)
        ws.cell(row=i, column=2, value=r.category)
        ws.cell(row=i, column=3, value=int(r.count))
        amt = ws.cell(row=i, column=4, value=float(r.total))
        amt.number_format = MONEY_FMT2
        ws.cell(row=i, column=5, value=r.example)

    last = len(table) + 1
    dv = DataValidation(
        type="list",
        formula1=f"Lists!${list_col}$2:${list_col}${len(options) + 1}",
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    if last >= 2:
        dv.add(f"B2:B{max(last, 2)}")

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 60
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{max(last, 1)}"


def build_transactions(ws, df, n_merch, n_names):
    headers = ["date", "source", "flow", "merchant", "description",
               "debit", "credit", "month", "category"]
    _write_header(ws, 1, headers)
    m_end = n_merch + 1
    n_end = n_names + 1
    for i, r in enumerate(df.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=r.date)
        ws.cell(row=i, column=2, value=r.source)
        ws.cell(row=i, column=3, value=r.flow)
        ws.cell(row=i, column=4, value=r.merchant)
        ws.cell(row=i, column=5, value=r.description)
        dc = ws.cell(row=i, column=6, value=float(r.debit))
        dc.number_format = MONEY_FMT2
        cr = ws.cell(row=i, column=7, value=float(r.credit))
        cr.number_format = MONEY_FMT2
        ws.cell(row=i, column=8, value=f"=LEFT(A{i},7)")
        ws.cell(row=i, column=9, value=(
            f'=IFERROR(IF(C{i}="Spending",'
            f"VLOOKUP(D{i},Merchants!$A$2:$B${m_end},2,FALSE),"
            f"VLOOKUP(D{i},Names!$A$2:$B${n_end},2,FALSE)),\"Uncategorized\")"
        ))
    for col, width in (("A", 11), ("B", 8), ("C", 10), ("D", 30),
                       ("E", 60), ("F", 12), ("G", 12), ("H", 9), ("I", 20)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{len(df) + 1}"


def build_dashboard(ws, df):
    months = sorted(df["date"].str.slice(0, 7).unique().tolist())
    n_spend = len(cd.SPENDING_CATEGORIES)

    ws["A1"] = "My Spending"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"{len(df)} transactions  ·  {df['date'].min()} to {df['date'].max()}"
                "  ·  edit categories on the Merchants / Names sheets")
    ws["A2"].font = Font(color="6B7280")

    income_terms = "+".join(
        f'SUMIFS(Transactions!$G:$G,Transactions!$I:$I,"{lbl}")'
        for lbl in sorted(cd.INCOME_LABELS)
    )
    metrics = [
        ("Total spending", '=SUMIFS(Transactions!$F:$F,Transactions!$C:$C,"Spending")'),
        ("Total income", "=" + income_terms),
        ("Net", "=B5-B4"),
        ("Uncategorized spend",
         '=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"Uncategorized",'
         'Transactions!$C:$C,"Spending")'),
    ]
    for i, (label, formula) in enumerate(metrics, start=4):
        lc = ws.cell(row=i, column=1, value=label)
        lc.font = LABEL_FONT
        vc = ws.cell(row=i, column=2, value=formula)
        vc.number_format = MONEY_FMT

    # Month selector — drives the per-month category pie below.
    MONTH_CELL = "B8"
    sel = ws.cell(row=8, column=1, value="Pie month →")
    sel.font = LABEL_FONT
    mc = ws.cell(row=8, column=2, value=(months[-1] if months else ""))
    mc.font = Font(bold=True, color="2563EB")

    # Spending by category table
    cat_top = 10
    _write_header(ws, cat_top, ["Category", "Spent"])
    for j, cat in enumerate(cd.SPENDING_CATEGORIES):
        r = cat_top + 1 + j
        ws.cell(row=r, column=1, value=cat)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$C:$C,"Spending")'))
        v.number_format = MONEY_FMT
    cat_end = cat_top + n_spend

    # Monthly spending table
    mon_top = cat_end + 3
    _write_header(ws, mon_top, ["Month", "Spent"])
    for j, m in enumerate(months):
        r = mon_top + 1 + j
        ws.cell(row=r, column=1, value=m)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$H:$H,"{m}",'
            f'Transactions!$C:$C,"Spending")'))
        v.number_format = MONEY_FMT
    mon_end = mon_top + len(months)

    # Month dropdown validation (references the Month column just written).
    if months:
        dv = DataValidation(
            type="list",
            formula1=f"$A${mon_top + 1}:$A${mon_end}",
            allow_blank=False,
        )
        ws.add_data_validation(dv)
        dv.add(MONTH_CELL)

    # Income breakdown table
    inc_top = mon_end + 3
    _write_header(ws, inc_top, ["Income source", "Amount"])
    inc_labels = sorted(cd.INCOME_LABELS)
    for j, lbl in enumerate(inc_labels):
        r = inc_top + 1 + j
        ws.cell(row=r, column=1, value=lbl)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$G:$G,Transactions!$I:$I,"{lbl}")'))
        v.number_format = MONEY_FMT
    inc_end = inc_top + len(inc_labels)

    # Per-month spending by category (recomputes when the month cell changes)
    pm_top = inc_end + 3
    _write_header(ws, pm_top, ["Category", "Spent in selected month"])
    for j, cat in enumerate(cd.SPENDING_CATEGORIES):
        r = pm_top + 1 + j
        ws.cell(row=r, column=1, value=cat)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$H:$H,${MONTH_CELL[0]}${MONTH_CELL[1:]},'
            f'Transactions!$C:$C,"Spending")'))
        v.number_format = MONEY_FMT
    pm_end = pm_top + n_spend

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22

    # Charts (anchored in column D so they don't overlap the tables)
    cat_chart = PieChart()
    cat_chart.title = "Spending by category (all months)"
    cat_chart.height = 9
    cat_chart.width = 18
    data = Reference(ws, min_col=2, min_row=cat_top, max_row=cat_end)
    cats = Reference(ws, min_col=1, min_row=cat_top + 1, max_row=cat_end)
    cat_chart.add_data(data, titles_from_data=True)
    cat_chart.set_categories(cats)
    _text_categories(cat_chart, list(cd.SPENDING_CATEGORIES))
    cat_chart.dataLabels = DataLabelList()
    cat_chart.dataLabels.showPercent = True
    ws.add_chart(cat_chart, "D4")

    mon_chart = BarChart()
    mon_chart.type = "col"
    mon_chart.title = "Monthly spending"
    mon_chart.legend = None
    mon_chart.height = 9
    mon_chart.width = 18
    mdata = Reference(ws, min_col=2, min_row=mon_top, max_row=mon_end)
    mcats = Reference(ws, min_col=1, min_row=mon_top + 1, max_row=mon_end)
    mon_chart.add_data(mdata, titles_from_data=True)
    mon_chart.set_categories(mcats)
    _text_categories(mon_chart, months)
    mon_chart.x_axis.delete = False   # force month labels to render
    mon_chart.y_axis.delete = False
    ws.add_chart(mon_chart, "D24")

    pm_chart = PieChart()
    pm_chart.title = "Spending by category (selected month)"
    pm_chart.height = 9
    pm_chart.width = 18
    pdata = Reference(ws, min_col=2, min_row=pm_top, max_row=pm_end)
    pcats = Reference(ws, min_col=1, min_row=pm_top + 1, max_row=pm_end)
    pm_chart.add_data(pdata, titles_from_data=True)
    pm_chart.set_categories(pcats)
    _text_categories(pm_chart, list(cd.SPENDING_CATEGORIES))
    pm_chart.dataLabels = DataLabelList()
    pm_chart.dataLabels.showPercent = True
    ws.add_chart(pm_chart, "D44")


def build_month_matrix(ws, df):
    """A category × month grid: one row per spending category, one column per
    month, plus a Total. A second grid below shows each category's share of
    that month's spending (the 'portion' view)."""
    months = sorted(df["date"].str.slice(0, 7).unique().tolist())
    cats = list(cd.SPENDING_CATEGORIES)
    n = len(cats)
    total_col = len(months) + 2          # col index of the row-total column
    last_month_col = len(months) + 1
    L = get_column_letter

    ws["A1"] = "Spending by category — monthly"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Amounts on top · each category's share of the month below"
    ws["A2"].font = Font(color="6B7280")

    # ── Amounts grid ──
    amt_hdr = 4
    _write_header(ws, amt_hdr, ["Category ($)"] + months + ["Total"])
    amt_first = amt_hdr + 1
    amt_last = amt_hdr + n
    for j, cat in enumerate(cats):
        r = amt_first + j
        ws.cell(row=r, column=1, value=cat).font = LABEL_FONT
        for k, m in enumerate(months):
            c = ws.cell(row=r, column=2 + k, value=(
                f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
                f'Transactions!$H:$H,"{m}",Transactions!$C:$C,"Spending")'))
            c.number_format = MONEY_FMT
        rt = ws.cell(row=r, column=total_col,
                     value=f"=SUM(B{r}:{L(last_month_col)}{r})")
        rt.number_format = MONEY_FMT
        rt.font = LABEL_FONT
    amt_total = amt_last + 1
    ws.cell(row=amt_total, column=1, value="Total").font = HEADER_FONT
    ws.cell(row=amt_total, column=1).fill = HEADER_FILL
    for c in range(2, total_col + 1):
        col = L(c)
        cell = ws.cell(row=amt_total, column=c,
                       value=f"=SUM({col}{amt_first}:{col}{amt_last})")
        cell.number_format = MONEY_FMT
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    # ── Share-of-month grid (reads the amounts above) ──
    pct_hdr = amt_total + 3
    _write_header(ws, pct_hdr, ["Category (% of month)"] + months)
    pct_first = pct_hdr + 1
    for j, cat in enumerate(cats):
        pr = pct_first + j
        ws.cell(row=pr, column=1, value=cat).font = LABEL_FONT
        ar = amt_first + j           # matching amounts row
        for k in range(len(months)):
            col = L(2 + k)
            cell = ws.cell(row=pr, column=2 + k, value=(
                f"=IFERROR({col}{ar}/{col}${amt_total},0)"))
            cell.number_format = "0.0%"

    ws.column_dimensions["A"].width = 24
    for c in range(2, total_col + 1):
        ws.column_dimensions[L(c)].width = 10
    ws.freeze_panes = "B5"


# ── workbook assembly (shared by CLI + web) ────────────────────────
def build_workbook(df, prior_m=None, prior_n=None):
    """Build the finance workbook from a classified DataFrame. Returns a Workbook."""
    spend = seed_table(df, "Spending", cd.DEFAULT_CATEGORIES)
    names = seed_table(df, "Transfer", cd.DEFAULT_TRANSFERS)

    if prior_m:
        spend["category"] = spend["name"].map(prior_m).fillna(spend["category"])
    if prior_n:
        names["category"] = names["name"].map(prior_n).fillna(names["category"])

    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_tx = wb.create_sheet("Transactions")
    ws_m = wb.create_sheet("Merchants")
    ws_n = wb.create_sheet("Names")
    ws_month = wb.create_sheet("By Month")
    ws_lists = wb.create_sheet("Lists")

    build_lists(ws_lists)
    build_mapping_sheet(ws_m, "Merchant", spend, SPENDING_OPTIONS, "A")
    build_mapping_sheet(ws_n, "Name", names, TRANSFER_OPTIONS, "B")
    build_transactions(ws_tx, df, len(spend), len(names))
    build_dashboard(ws_dash, df)
    build_month_matrix(ws_month, df)
    return wb, spend, names


def build_bytes(file_specs, prior_xlsx_bytes=None):
    """Web entry: classify uploaded CSVs and return the xlsx as bytes.

    file_specs : list of (file-like-or-str, source) — the uploaded CSVs.
    Raises ValueError if no usable transactions are found.
    """
    df = ingest_frames(file_specs)
    prior_m, prior_n = {}, {}
    if prior_xlsx_bytes:
        try:
            prev = load_workbook(io.BytesIO(prior_xlsx_bytes), read_only=True)
            prior_m, prior_n, prior_tx = read_prior(prev)
            prev.close()
            df = merge_prior(df, prior_tx)
        except Exception:
            pass
    if df is None or df.empty:
        raise ValueError("No transactions found in the uploaded files.")
    wb, _, _ = build_workbook(df, prior_m, prior_n)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ── CLI ────────────────────────────────────────────────────────────
def build():
    df = ingest()
    if df.empty:
        print("No CSV files found in credit/ or debit/. Add your CIBC exports and rerun.")
        return

    prior_m, prior_n = {}, {}
    if os.path.exists(XLSX):
        try:
            prev = load_workbook(XLSX, read_only=True)
            prior_m, prior_n, prior_tx = read_prior(prev)
            prev.close()
            df = merge_prior(df, prior_tx)
        except Exception:
            pass
    wb, spend, names = build_workbook(df, prior_m, prior_n)

    try:
        wb.save(XLSX)
    except PermissionError:
        print(f"Could not write {os.path.basename(XLSX)} — close it in Excel and rerun.")
        return

    n_un_s = int((spend["category"] == "Uncategorized").sum())
    n_un_n = int((names["category"] == "Uncategorized").sum())
    print(f"Wrote {os.path.relpath(XLSX, cd.BASE)}")
    print(f"  {len(df)} transactions  ·  {len(spend)} merchants  ·  {len(names)} transfer names")
    print(f"  Still 'Uncategorized': {n_un_s} merchants, {n_un_n} names "
          "(label them on the Merchants / Names sheets).")


if __name__ == "__main__":
    build()

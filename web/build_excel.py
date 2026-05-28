"""
Build a self-contained Excel workbook from your CIBC CSV exports.

Run:
    python3 build_excel.py

This reads every CSV in credit/ and debit/, classifies each transaction, and
writes finance.xlsx. After that, EVERYTHING happens inside Excel -- no Python,
no internet, no AI:

  - Overview    : all-time numbers + charts (spending by category bar, monthly
                  trend, income, and transfers split into money in vs out).
  - Monthly     : pick a month and see a bulk-category pie, a detailed
                  breakdown, and that month's money in vs out.
  - Transactions: every cleaned transaction. The Category column is a formula
                  that looks up the merchant/name on the sheets below.
  - Merchants   : each unique spending merchant. Pick a Category from the
                  dropdown -- it applies to every matching transaction.
  - Names       : each unique transfer name (allowance, split bills, salary...),
                  with money received (In) and sent (Out) shown separately.

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
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties, Paragraph, ParagraphProperties)
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
DEDUP_KEY = ["date", "source", "description", "debit", "credit"]


def _finalize(df):
    """Normalize date and derive merchant / flow / amount on a raw frame."""
    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df.dropna(subset=["date"])
    df["merchant"] = df["description"].apply(cd.normalize_merchant)
    df["flow"] = [
        cd.ws_classify(r["description"])[0]
        if r["source"] == "wealthsimple"
        else cd.classify_flow(r["description"], r["debit"], r["credit"])
        for _, r in df.iterrows()
    ]
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
             for source in ("credit", "debit", "wealthsimple")
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
    sub["label"] = [
        cd.ws_classify(r["description"])[1]
        if r["source"] == "wealthsimple"
        else cd.apply_label(r["description"], mapping)
        for _, r in sub.iterrows()
    ]
    rows = []
    for name, g in sub.groupby("merchant"):
        labelled = g["label"][g["label"] != "Uncategorized"]
        cat = labelled.mode().iloc[0] if not labelled.empty else "Uncategorized"
        rows.append({
            "name": name,
            "category": cat,
            "count": len(g),
            "total": round(float(g["amount"].sum()), 2),
            "inflow": round(float(g["credit"].sum()), 2),
            "outflow": round(float(g["debit"].sum()), 2),
            "example": g["description"].iloc[0],
        })
    if not rows:
        return pd.DataFrame(columns=["name", "category", "count", "total",
                                     "inflow", "outflow", "example"])
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


def _rich_size(pt, bold=False):
    """A RichText block that forces a font size (points) on chart text."""
    cp = CharacterProperties(sz=int(pt * 100), b=bold)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),
                                 endParaRPr=cp)])


def _set_title(chart, text, pt=15):
    """Set a chart title and enlarge its font."""
    chart.title = text
    cp = CharacterProperties(sz=int(pt * 100), b=True)
    for para in chart.title.tx.rich.p:
        para.pPr = ParagraphProperties(defRPr=cp)
        for run in (para.r or []):
            run.rPr = cp


def _style_pie(chart):
    """Big readable pie: right-hand legend and larger label/legend fonts."""
    chart.legend.position = "r"
    chart.legend.txPr = _rich_size(11)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showLegendKey = False
    chart.dataLabels.showVal = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False
    chart.dataLabels.txPr = _rich_size(10, bold=True)


def _default_category_totals(df):
    """Spend per category using the SAME merchant→category resolution the
    workbook uses, so the ordering matches what the user first sees. Used only
    to ORDER the charts — every category is still plotted (see below) so the
    chart stays correct after the user re-categorizes a merchant."""
    spend_tbl = seed_table(df, "Spending", cd.DEFAULT_CATEGORIES)
    cat_by_merchant = dict(zip(spend_tbl["name"], spend_tbl["category"]))
    sub = df[df["flow"] == "Spending"].copy()
    if sub.empty:
        return {}
    sub["cat"] = sub["merchant"].map(cat_by_merchant).fillna("Uncategorized")
    return sub.groupby("cat")["debit"].sum().to_dict()


def _chart_categories(df):
    """Every spending category, ordered by current spend (largest first).

    We plot ALL categories — not just the ones with spend today — so a category
    the user later assigns a merchant to still shows up without a rebuild. The
    "Uncategorized" bucket is included whenever it has spend, so the charts tie
    out to Total spending instead of quietly dropping unlabelled money."""
    totals = _default_category_totals(df)
    cats = list(cd.SPENDING_CATEGORIES)
    if totals.get("Uncategorized", 0) > 0:
        cats.append("Uncategorized")
    cats.sort(key=lambda c: totals.get(c, 0.0), reverse=True)
    return cats


def _chart_groups(df):
    """Every bulk group (plus Uncategorized when present), largest first."""
    totals = _default_category_totals(df)
    group_total = {}
    for cat, amt in totals.items():
        group_total[cd.spending_group(cat) if cat != "Uncategorized"
                    else "Uncategorized"] = (
            group_total.get(cd.spending_group(cat) if cat != "Uncategorized"
                            else "Uncategorized", 0.0) + amt)
    groups = list(cd.SPENDING_GROUPS)
    if totals.get("Uncategorized", 0) > 0:
        groups.append("Uncategorized")
    groups.sort(key=lambda g: group_total.get(g, 0.0), reverse=True)
    return groups


def build_lists(ws):
    ws["A1"] = "SpendingOptions"
    for i, opt in enumerate(SPENDING_OPTIONS, start=2):
        ws.cell(row=i, column=1, value=opt)
    ws["B1"] = "TransferOptions"
    for i, opt in enumerate(TRANSFER_OPTIONS, start=2):
        ws.cell(row=i, column=2, value=opt)
    ws.sheet_state = "hidden"


def build_mapping_sheet(ws, name_header, table, options, list_col,
                        show_direction=False):
    """One row per merchant/transfer name with an editable Category dropdown.

    For transfer names (show_direction=True) the single "Total" is split into
    "In" (money received) and "Out" (money sent) so a person you both pay and
    get paid by isn't collapsed into one misleading number."""
    if show_direction:
        _write_header(ws, 1,
                      [name_header, "Category", "Count", "In", "Out", "Example"])
    else:
        _write_header(ws, 1, [name_header, "Category", "Count", "Total", "Example"])
    for i, r in enumerate(table.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=r.name)
        ws.cell(row=i, column=2, value=r.category)
        ws.cell(row=i, column=3, value=int(r.count))
        if show_direction:
            ci = ws.cell(row=i, column=4, value=float(r.inflow))
            ci.number_format = MONEY_FMT2
            co = ws.cell(row=i, column=5, value=float(r.outflow))
            co.number_format = MONEY_FMT2
            ws.cell(row=i, column=6, value=r.example)
        else:
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
    if show_direction:
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 60
        last_col = "F"
    else:
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 60
        last_col = "E"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{last_col}{max(last, 1)}"


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


def _mref(month_cell):
    """Absolute reference string for the month-selector cell (e.g. B4 -> $B$4)."""
    return f"${month_cell[0]}${month_cell[1:]}"


def build_overview(ws, df, ws_lists):
    """All-time view: totals, spending by category (bar), the monthly trend,
    income, and transfers split by direction (money in vs out)."""
    months = sorted(df["date"].str.slice(0, 7).unique().tolist())

    ws["A1"] = "Overview — all transactions"
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
        ("Net (income − spending)", "=B5-B4"),
        ("Uncategorized spend",
         '=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"Uncategorized",'
         'Transactions!$C:$C,"Spending")'),
    ]
    for i, (label, formula) in enumerate(metrics, start=4):
        ws.cell(row=i, column=1, value=label).font = LABEL_FONT
        vc = ws.cell(row=i, column=2, value=formula)
        vc.number_format = MONEY_FMT

    # Spending by category (all categories, for reference)
    cat_top = 9
    _write_header(ws, cat_top, ["Category", "Spent"])
    for j, cat in enumerate(cd.SPENDING_CATEGORIES):
        r = cat_top + 1 + j
        ws.cell(row=r, column=1, value=cat)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$C:$C,"Spending")'))
        v.number_format = MONEY_FMT
    cat_end = cat_top + len(cd.SPENDING_CATEGORIES)

    # Monthly trend table (drives the over-time column chart)
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

    # Income breakdown
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

    # Transfers — money IN vs OUT, per transfer category. A transfer category
    # can have both (e.g. you both pay and get paid by the same person), so
    # the two directions are shown side by side rather than netted.
    tr_top = inc_end + 3
    _write_header(ws, tr_top, ["Transfer type", "In", "Out"])
    tr_cats = cd.TRANSFER_CATEGORIES
    for j, cat in enumerate(tr_cats):
        r = tr_top + 1 + j
        internal = cat in cd.INTERNAL_TRANSFERS
        ws.cell(row=r, column=1,
                value=(cat + " (internal)") if internal else cat)
        ci = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$G:$G,Transactions!$I:$I,"{cat}",'
            f'Transactions!$C:$C,"Transfer")'))
        ci.number_format = MONEY_FMT
        co = ws.cell(row=r, column=3, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$C:$C,"Transfer")'))
        co.number_format = MONEY_FMT
    tr_end = tr_top + len(tr_cats)
    tot = ws.cell(row=tr_end + 1, column=1, value="Total")
    tot.font = LABEL_FONT
    for c in (2, 3):
        col = get_column_letter(c)
        cell = ws.cell(row=tr_end + 1, column=c,
                       value=f"=SUM({col}{tr_top + 1}:{col}{tr_end})")
        cell.number_format = MONEY_FMT
        cell.font = LABEL_FONT
    note = ws.cell(row=tr_end + 2, column=1, value=(
        "(internal) = your own money moving between accounts / card payments — "
        "it appears on both sides and nets to ~0, so it isn't real income or spend."))
    note.font = Font(italic=True, color="6B7280")

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14

    # Helper table for the bar: ALL categories, ascending so the largest lands
    # at the top and unused ones sink to the bottom. Values stay live SUMIFS.
    # Kept on the hidden Lists sheet so it never clutters (or gets edited on)
    # the Overview; charts read from a hidden sheet without issue.
    chart_cats = list(reversed(_chart_categories(df)))
    HCAT, HALL, hdr = 4, 5, 1
    ws_lists.cell(row=hdr, column=HALL, value="All months")
    for j, cat in enumerate(chart_cats):
        r = hdr + 1 + j
        ws_lists.cell(row=r, column=HCAT, value=cat)
        a = ws_lists.cell(row=r, column=HALL, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$C:$C,"Spending")'))
        a.number_format = MONEY_FMT
    h_end = hdr + len(chart_cats)
    n_cats = max(len(chart_cats), 1)

    WIDE = 28
    cat_chart = BarChart()
    cat_chart.type = "bar"
    _set_title(cat_chart, "Spending by category — all months", 16)
    cat_chart.legend = None
    cat_chart.width = WIDE
    cat_chart.height = max(10, 0.85 * n_cats + 3)
    cat_chart.add_data(Reference(ws_lists, min_col=HALL, min_row=hdr, max_row=h_end),
                       titles_from_data=True)
    cat_chart.set_categories(Reference(ws_lists, min_col=HCAT, min_row=hdr + 1, max_row=h_end))
    _text_categories(cat_chart, chart_cats)
    cat_chart.dataLabels = DataLabelList()
    cat_chart.dataLabels.showVal = True
    cat_chart.dataLabels.numFmt = MONEY_FMT
    cat_chart.dataLabels.txPr = _rich_size(10, bold=True)
    cat_chart.x_axis.delete = False
    cat_chart.y_axis.delete = False
    cat_chart.x_axis.txPr = _rich_size(10)
    cat_chart.y_axis.txPr = _rich_size(11)
    ws.add_chart(cat_chart, "D3")

    mon_chart = BarChart()
    mon_chart.type = "col"
    _set_title(mon_chart, "Monthly spending over time", 16)
    mon_chart.legend = None
    mon_chart.height = 11
    mon_chart.width = WIDE
    mon_chart.add_data(Reference(ws, min_col=2, min_row=mon_top, max_row=mon_end),
                       titles_from_data=True)
    mon_chart.set_categories(Reference(ws, min_col=1, min_row=mon_top + 1, max_row=mon_end))
    _text_categories(mon_chart, months)
    mon_chart.x_axis.delete = False
    mon_chart.y_axis.delete = False
    mon_chart.x_axis.txPr = _rich_size(10)
    mon_chart.y_axis.txPr = _rich_size(10)
    ws.add_chart(mon_chart, f"D{max(int(0.85 * n_cats + 3) + 6, 26)}")


def build_monthly(ws, df, ws_lists):
    """Single-month view: pick a month, then see a bulk-category pie, the
    detailed breakdown, and money in vs out for that month."""
    months = sorted(df["date"].str.slice(0, 7).unique().tolist())
    MONTH_CELL = "B4"
    mref = _mref(MONTH_CELL)

    ws["A1"] = "Monthly view"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Pick a month in cell B4 — every table and chart below updates."
    ws["A2"].font = Font(color="6B7280")

    ws.cell(row=4, column=1, value="Month →").font = LABEL_FONT
    mc = ws.cell(row=4, column=2, value=(months[-1] if months else ""))
    mc.font = Font(bold=True, color="2563EB")

    # Money in vs out for the selected month
    sum_top = 6
    _write_header(ws, sum_top, ["This month", "Amount"])
    spend_f = (f'SUMIFS(Transactions!$F:$F,Transactions!$H:$H,{mref},'
               f'Transactions!$C:$C,"Spending")')
    income_f = "+".join(
        f'SUMIFS(Transactions!$G:$G,Transactions!$I:$I,"{lbl}",'
        f'Transactions!$H:$H,{mref})' for lbl in sorted(cd.INCOME_LABELS))
    internal_in = "+".join(
        f'SUMIFS(Transactions!$G:$G,Transactions!$I:$I,"{cat}",'
        f'Transactions!$H:$H,{mref})' for cat in sorted(cd.INTERNAL_TRANSFERS))
    internal_out = "+".join(
        f'SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
        f'Transactions!$H:$H,{mref})' for cat in sorted(cd.INTERNAL_TRANSFERS))
    transfer_in_f = (f'SUMIFS(Transactions!$G:$G,Transactions!$H:$H,{mref},'
                     f'Transactions!$C:$C,"Transfer")-({internal_in})')
    transfer_out_f = (f'SUMIFS(Transactions!$F:$F,Transactions!$H:$H,{mref},'
                      f'Transactions!$C:$C,"Transfer")-({internal_out})')
    summary = [
        ("Spending (out)", "=" + spend_f),
        ("Income (in)", "=" + income_f),
        ("Net (income − spending)", f"=({income_f})-({spend_f})"),
        ("Transfers received (excl. internal)", "=" + transfer_in_f),
        ("Transfers sent (excl. internal)", "=" + transfer_out_f),
    ]
    for j, (label, formula) in enumerate(summary):
        r = sum_top + 1 + j
        ws.cell(row=r, column=1, value=label).font = LABEL_FONT
        v = ws.cell(row=r, column=2, value=formula)
        v.number_format = MONEY_FMT
    sum_end = sum_top + len(summary)
    snote = ws.cell(row=sum_end + 1, column=1, value=(
        "Internal moves (own-account transfers, card payments) are excluded — "
        "they'd double-count money you already see as spending."))
    snote.font = Font(italic=True, color="6B7280")
    sum_end += 1

    # Bulk-group table for the month (drives the pie). Includes Uncategorized
    # when present so the pie ties out to the month's total spending.
    members = {}
    for c in cd.SPENDING_CATEGORIES:
        members.setdefault(cd.spending_group(c), []).append(c)
    groups = _chart_groups(df)
    grp_top = sum_end + 3
    _write_header(ws, grp_top, ["Group", "Spent this month"])
    for j, g in enumerate(groups):
        r = grp_top + 1 + j
        ws.cell(row=r, column=1, value=g)
        terms = "+".join(
            f'SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{c}",'
            f'Transactions!$H:$H,{mref},Transactions!$C:$C,"Spending")'
            for c in members.get(g, [g]))   # Uncategorized: match itself
        v = ws.cell(row=r, column=2, value="=" + terms)
        v.number_format = MONEY_FMT
    grp_end = grp_top + len(groups)

    # Detailed category table for the month (drives the detail bar)
    detail_cats = list(reversed(_chart_categories(df)))
    det_top = grp_end + 3
    _write_header(ws, det_top, ["Category", "Spent this month"])
    for j, cat in enumerate(detail_cats):
        r = det_top + 1 + j
        ws.cell(row=r, column=1, value=cat)
        v = ws.cell(row=r, column=2, value=(
            f'=SUMIFS(Transactions!$F:$F,Transactions!$I:$I,"{cat}",'
            f'Transactions!$H:$H,{mref},Transactions!$C:$C,"Spending")'))
        v.number_format = MONEY_FMT
    det_end = det_top + len(detail_cats)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16

    # Month dropdown — validate against a month list kept on the hidden Lists
    # sheet so it never appears as stray data on the Monthly view.
    ML = 7
    for j, m in enumerate(months):
        ws_lists.cell(row=2 + j, column=ML, value=m)
    if months:
        col = get_column_letter(ML)
        dv = DataValidation(type="list",
                            formula1=f"Lists!${col}$2:${col}${1 + len(months)}",
                            allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(MONTH_CELL)

    WIDE = 28
    pie = PieChart()
    _set_title(pie, "Spending by group — selected month (cell B4)", 16)
    pie.height = 16
    pie.width = WIDE
    pie.add_data(Reference(ws, min_col=2, min_row=grp_top, max_row=grp_end),
                 titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=grp_top + 1, max_row=grp_end))
    _text_categories(pie, groups)
    _style_pie(pie)
    ws.add_chart(pie, "D3")

    n_det = max(len(detail_cats), 1)
    bar = BarChart()
    bar.type = "bar"
    _set_title(bar, "Spending by category — selected month", 16)
    bar.legend = None
    bar.width = WIDE
    bar.height = max(10, 0.85 * n_det + 3)
    bar.add_data(Reference(ws, min_col=2, min_row=det_top, max_row=det_end),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=det_top + 1, max_row=det_end))
    _text_categories(bar, detail_cats)
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    bar.dataLabels.numFmt = MONEY_FMT
    bar.dataLabels.txPr = _rich_size(10, bold=True)
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.x_axis.txPr = _rich_size(10)
    bar.y_axis.txPr = _rich_size(11)
    ws.add_chart(bar, "D36")


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
    ws_over = wb.active
    ws_over.title = "Overview"
    ws_monthly = wb.create_sheet("Monthly")
    ws_tx = wb.create_sheet("Transactions")
    ws_m = wb.create_sheet("Merchants")
    ws_n = wb.create_sheet("Names")
    ws_month = wb.create_sheet("By Month")
    ws_lists = wb.create_sheet("Lists")

    build_lists(ws_lists)
    build_mapping_sheet(ws_m, "Merchant", spend, SPENDING_OPTIONS, "A")
    build_mapping_sheet(ws_n, "Name", names, TRANSFER_OPTIONS, "B",
                        show_direction=True)
    build_transactions(ws_tx, df, len(spend), len(names))
    build_overview(ws_over, df, ws_lists)
    build_monthly(ws_monthly, df, ws_lists)
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
        # The user deliberately supplied a previous workbook to carry labels and
        # old transactions forward. If it can't be read, fail loudly rather than
        # silently dropping their saved categories.
        try:
            prev = load_workbook(io.BytesIO(prior_xlsx_bytes), read_only=True)
            prior_m, prior_n, prior_tx = read_prior(prev)
            prev.close()
            df = merge_prior(df, prior_tx)
        except Exception as e:
            raise ValueError(
                "Couldn't read the previous finance.xlsx you provided, so your "
                "saved categories weren't carried over. Remove it and try again, "
                f"or upload the correct file. (details: {e})")
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
        except Exception as e:
            print(f"WARNING: couldn't read the existing {os.path.basename(XLSX)} "
                  f"({e}). Your previously saved categories will NOT be carried "
                  "over. Move or fix that file if you want to keep them.")
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

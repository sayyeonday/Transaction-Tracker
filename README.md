# Expense Tracker

Turn your **CIBC or BMO** bank CSV exports into a tidy Excel workbook that shows
**where your money goes** and **how much comes in** — with click-to-categorize
dropdowns and charts that update themselves. The bank is detected automatically.

No accounts, no internet, no AI. Your data stays with you.

### ▶︎ Use it now: <https://sayyeonday.github.io/Transaction-Tracker/>

Just open the link, pick your CSV files, and download your workbook. Nothing to
install — it all runs in your browser.

---

## Two ways to use it

### A. The website (easiest, most private)

Open <https://sayyeonday.github.io/Transaction-Tracker/>, pick your CSV files, and
download `finance.xlsx`. **Everything runs inside your own browser — your
transactions are never uploaded or stored anywhere.**

To run it locally instead:

```bash
python3 -m http.server 8530 --directory web
```

Then open <http://localhost:8530>. (It must be served over HTTP — opening the
`.html` file directly won't work, because the in-browser Python engine needs to
fetch its files.)

### B. Local Python script

```bash
pip install -r requirements.txt
# put your CIBC or BMO CSV exports in:  credit/  and  debit/
python3 build_excel.py
```

This writes `finance.xlsx` next to the script.

---

## How to get your CSVs

In online banking, open an account, choose **Download transactions**, and pick
**CSV**. Do this for each card/account. The tool detects which bank a file came
from automatically — no need to tell it.

**CIBC** files have no header row and look like:

```
2026-05-26,"WAL-MART SUPERCENTER#1234 ANYTOWN, ON",4.50,,1234********5678
2026-05-26,Internet Banking E-TRANSFER 000000000000 JANE DOE,,50.00
```

**BMO** files have a header row. There are two layouts and both are detected:

*Credit card* — one signed amount column (a purchase is positive, a payment
received is negative):

```
Item #,Card #,Transaction Date,Posting Date,Transaction Amount,Description
1,1234********5678,20260526,20260527,4.50,"WAL-MART SUPERCENTER ANYTOWN, ON"
2,1234********5678,20260518,20260519,-816.19,PAYMENT RECEIVED - THANK YOU
```

*Chequing / debit* — a **Transaction Type** column (DR = money out, CR = money
in) sets the direction:

```
First Bank Card,Transaction Type,Date Posted,Transaction Amount,Description
'5100********1234',DR,20260524,-12.99,WAL-MART SUPERCENTER
'5100********1234',CR,20260520,2000.00,PAYROLL DEPOSIT
```

---

## Using the workbook

`finance.xlsx` has these tabs:

- **Overview** — all-time totals (spending, income, net), a spending-by-category
  bar, the monthly spending trend, an income breakdown, and a **Transfers**
  table that splits each transfer type into money **In** vs **Out**.
- **Monthly** — pick a month in cell **B4** and everything updates: a
  bulk-category **pie** (Food, Shopping, Transport…), a detailed category bar,
  and that month's money **in vs out** summary.
- **Transactions** — every cleaned transaction. The *Category* column is a
  formula that looks up the merchant/name from the tabs below.
- **Merchants** — each unique place you spent money. Pick a **Category** from the
  dropdown and it applies to every matching transaction.
- **Names** — each unique transfer name (allowance, split bills, salary, your
  own accounts…). Same dropdown idea, with money received (**In**) and sent
  (**Out**) shown separately so direction is never lost.

Workflow: open the **Merchants** and **Names** tabs, choose a category for each
row from the dropdown, and watch the **Overview** / **Monthly** tabs update.
That's it.

### Adding new statements (and keeping your labels)

Re-run the tool (or re-upload on the website) and **also provide your previous
`finance.xlsx`**. Two things are carried over:

- **Your category choices** — read back from the Merchants / Names tabs, so you
  only ever label the *new* merchants.
- **Your older transactions** — the previous workbook's transactions are merged
  with the new CSVs, so the workbook grows cumulatively instead of starting over.

To avoid double-counting, a transaction from the old workbook is dropped only
when the new files already contain it (same date, description, and amount). If
you happen to make two identical purchases on the same day, both are kept — they
only collapse when a re-export of the *same period* would otherwise duplicate them.

---

## How it classifies (no AI)

Every transaction is first split into a **flow**:

- **Spending** — money spent on goods/services/fees.
- **Transfer** — money *moving* (e-transfers, internal transfers, card payments,
  deposits) rather than being consumed.

The detailed category comes from a built-in keyword dictionary (e.g. `WALMART` →
`Food_Groceries`), matched against the raw transaction text. You refine the rest
with the dropdowns. There is no machine learning and nothing is sent online.

---

## Project layout

```
clean_data.py    # classification library (read / flow / merchant / dictionary)
build_excel.py   # builds finance.xlsx — CLI (build) and in-browser (build_bytes)
web/             # the static website (deployable as-is)
  index.html
  main.js
  clean_data.py     # copies used by the site
  build_excel.py
credit/  debit/  # put your CIBC or BMO CSV exports here for the local script
requirements.txt
```

> The `web/` folder keeps copies of the two Python files so it can be deployed on
> its own. If you edit the Python at the repo root, copy them back in
> (`cp clean_data.py build_excel.py web/`). The GitHub Pages workflow does this
> automatically on every deploy.

---

## Deploying the website to GitHub Pages

1. Create a GitHub repo and push this project to the `main` branch.
2. In the repo, go to **Settings → Pages** and set **Source = GitHub Actions**.
3. Every push to `main` builds and publishes the site (see
   `.github/workflows/deploy-pages.yml`). Your live URL appears in the Actions run.

---

## Privacy

The website runs entirely client-side. The only thing downloaded from the
internet is the Python engine (Pyodide) and its libraries — **your financial
data is never transmitted, logged, or stored.** Close the tab and nothing remains.

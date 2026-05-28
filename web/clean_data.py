"""
CIBC transaction classification library.

Design
  - Every transaction is first split into one of two FLOWS:
        Spending = money spent on goods / services / fees
        Transfer = money MOVING (e-transfers, internal transfers, card
                   payments, deposits) — not consumption.
    The flow is decided by matching MOVE_PATTERNS against the raw text.
  - The detailed LABEL is assigned from a built-in dictionary (keyword ->
    label, matched as a substring of the raw text). Matching is done on the
    RAW text, not the normalized name, so it stays robust even when
    normalization is imperfect.
  - normalize_merchant() only produces a readable name for grouping
    transactions (it never drives classification).

This module is a library — it reads and classifies, but writes nothing.
build_excel.py imports it to produce finance.xlsx, where the user assigns the
final categories via dropdowns. No internet or AI is used at any point.
"""

import os
import re
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))

# ── Flow detection: if any of these appears in the raw text -> "Transfer" ──
MOVE_PATTERNS = [
    "E-TRANSFER", "E-TRANSFERT", "INTERNET TRANSFER", "MEMO - TRANSFER",
    "FULFILL REQUEST", "PAYMENT THANK YOU", "PAIEMEN", "TO CARD",
    "ELECTRONIC FUNDS TRANSFER", "DEPOSIT",
]

# ── The label vocabularies (also used to populate the UI dropdowns) ──
SPENDING_CATEGORIES = [
    "Food_DiningOut", "Food_Cafe", "Food_Groceries",
    "Shopping_General", "Shopping_Beauty",
    "Education", "Subscription_Software", "Travel_Leisure",
    "Health_Medical", "Transport_Daily",
    "Fees_Banking", "Cash_Withdrawal", "Other",
]
TRANSFER_CATEGORIES = [
    "Transfer_Allowance", "Transfer_Reimbursed", "Transfer_SplitSent",
    "Transfer_Income", "Transfer_OwnAccount", "Transfer_Investment",
    "Transfer_CardPayment",
    "Income_Salary", "Income_PartTime", "Income_SideHustle",
    "Income_Rewards", "Income_Financial", "Other",
]
# Labels that represent money flowing IN to the user (for the dashboard).
INCOME_LABELS = {
    "Income_Salary", "Income_PartTime", "Income_SideHustle",
    "Income_Rewards", "Income_Financial",
    "Transfer_Allowance", "Transfer_Income", "Transfer_Reimbursed",
}

# ── Default dictionaries shipped with the code (the user overrides via CSV) ──
DEFAULT_CATEGORIES = {
    # Food_Groceries
    "WALMART": "Food_Groceries", "WAL-MART": "Food_Groceries", "FOODMART": "Food_Groceries",
    "TERRA FOOD": "Food_Groceries", "NO FRILLS": "Food_Groceries", "COSTCO": "Food_Groceries",
    "SUPERSTORE": "Food_Groceries", "TASC": "Food_Groceries",
    # Food_Cafe
    "TIM HORTONS": "Food_Cafe", "STARBUCKS": "Food_Cafe", "CAFE": "Food_Cafe",
    "CREPE": "Food_Cafe", "TEN REN": "Food_Cafe", "NAYAX": "Food_Cafe",
    # Food_DiningOut
    "MCDONALD": "Food_DiningOut", "EARLS": "Food_DiningOut", "PAK PAK": "Food_DiningOut",
    "YAMA": "Food_DiningOut", "JUSFRES": "Food_DiningOut", "KITCHEN": "Food_DiningOut",
    "PIZZA": "Food_DiningOut", "SUSHI": "Food_DiningOut", "RESTAURANT": "Food_DiningOut",
    "KIM'S KOREAN": "Food_DiningOut", "BAR BIRU": "Food_DiningOut", "ROL SAN": "Food_DiningOut",
    # Shopping_General
    "TEMU": "Shopping_General", "AMZN": "Shopping_General", "AMAZON": "Shopping_General",
    "SHEIN": "Shopping_General", "MINISO": "Shopping_General", "KIOKII": "Shopping_General",
    "DOLLARAMA": "Shopping_General", "CHAPTERS": "Shopping_General",
    "URBAN PLANET": "Shopping_General", "DICKIES": "Shopping_General",
    "ARITZIA": "Shopping_General", "COACH": "Shopping_General", "CIRCLE K": "Shopping_General",
    # Shopping_Beauty
    "SEPHORA": "Shopping_Beauty", "HAIR SHINE": "Shopping_Beauty",
    # Education
    "UTM": "Education", "U OF T": "Education", "UNIVERSITY OF TORONTO": "Education",
    "BOOKSTORE": "Education", "RECREATION AT U OF T": "Education", "LCC": "Education",
    "TRINITY": "Education",
    # Subscription_Software
    "OPENAI": "Subscription_Software", "CHATGPT": "Subscription_Software",
    "GOOGLE": "Subscription_Software", "NETFLIX": "Subscription_Software",
    "CLOUD": "Subscription_Software", "APPLE.COM": "Subscription_Software",
    "4KDOWNLOAD": "Subscription_Software", "SPOTIFY": "Subscription_Software",
    # Travel_Leisure
    "VIA RAIL": "Travel_Leisure", "UNITED": "Travel_Leisure", "ASIANA": "Travel_Leisure",
    "CINEPLEX": "Travel_Leisure", "TICKETMASTER": "Travel_Leisure",
    "EVENTBRITE": "Travel_Leisure", "PEARSON PARKING": "Travel_Leisure", "RATP": "Travel_Leisure",
    # Health_Medical
    "SHOPPERS DRUG": "Health_Medical", "ARC EYECARE": "Health_Medical",
    # Transport_Daily
    "UBER": "Transport_Daily", "UBR*": "Transport_Daily", "LYFT": "Transport_Daily",
    "PRESTO": "Transport_Daily", "GO TRANSIT": "Transport_Daily", "PARKING": "Transport_Daily",
    # Fees_Banking / Cash
    "SERVICE CHARGE": "Fees_Banking", "NETWORK TRANSACTION FEE": "Fees_Banking",
    "FX CASH": "Fees_Banking", "ATM WITHDRAWAL": "Cash_Withdrawal",
}

# Only the mechanical movements are pre-filled. Named e-transfers / deposits
# are left blank so the user picks (Allowance / SplitSent / Income / ...).
DEFAULT_TRANSFERS = {
    "PAYMENT THANK YOU": "Transfer_CardPayment", "PAIEMEN": "Transfer_CardPayment",
    "TO CARD": "Transfer_CardPayment",
    "INTERNET TRANSFER": "Transfer_OwnAccount", "MEMO - TRANSFER": "Transfer_OwnAccount",
}

# ── Display-only noise lists (never affect classification) ──
PROVINCES = {"ON", "NB", "BC", "AB", "QC", "NS", "MB", "SK", "PE", "NL", "NT", "YT", "NU"}
COMMON_CITIES = ["MISSISSAUGA", "MISS", "TORONTO", "BRAMPTON", "FREDERICTON", "OAKVILLE",
                 "SCARBOROUGH", "MONCTON", "NIAGARA", "ST CATHERINES", "MONTREAL"]
PROCESSORS = ["SQ", "TST", "CS", "PAYPAL", "PP", "SP", "SQUARE"]
PREFIX_PATTERNS = [
    r"^POINT OF SALE - INTERAC RETAIL PURCHASE",
    r"^POINT OF SALE - VISA DEBIT INTL VISA DEB RETAIL PURCHASE",
    r"^POINT OF SALE - VISA DEBIT VISA DEBIT RETAIL PURCHASE",
    r"^POINT OF SALE - VISA DEBIT CORRECTION",
    r"^POINT OF SALE - VISA DEBIT INT'?L? VISA DEB( RETAIL)?( PURCHASE)?( REVERSAL)?",
    r"^POINT OF SALE - VISA DEBIT",
    r"^INTERNET BANKING INTERNET BILL PAY",
    r"^INTERNET BANKING E-TRANSFER",
    r"^INTERNET BANKING",
    r"^AUTOMATED BANKING MACHINE INTL ATM WITHDRAWAL\*?",
    r"^AUTOMATED BANKING MACHINE",
    r"^BRANCH TRANSACTION MEMO - TRANSFER",
    r"^BRANCH TRANSACTION",
    r"^ELECTRONIC FUNDS TRANSFER( CREDIT MEMO)?",
]


# ── 1. Read CSVs ───────────────────────────────────────────────
def read_cibc_csv(path, source):
    """CIBC CSVs have no header and 4-5 ragged columns, so map them manually."""
    raw = pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
    records = []
    for row in raw.itertuples(index=False):
        cells = list(row)
        records.append({
            "date": cells[0].strip(),
            "description": cells[1].strip(),
            "debit": _to_float(cells[2]) if len(cells) > 2 else 0.0,   # money out
            "credit": _to_float(cells[3]) if len(cells) > 3 else 0.0,  # money in
            "source": source,
        })
    return pd.DataFrame(records)


def _to_float(s):
    s = (s or "").strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _mtext(desc):
    """Matching text: collapse whitespace + uppercase."""
    return re.sub(r"\s+", " ", desc).upper().strip()


# ── 2. Flow + label + display name ─────────────────────────────
def classify_flow(desc):
    """Return 'Transfer' if a move signal is present, else 'Spending'."""
    u = _mtext(desc)
    return "Transfer" if any(p in u for p in MOVE_PATTERNS) else "Spending"


def apply_label(desc, mapping):
    """Substring-match dictionary keywords against the raw text (longest wins)."""
    u = _mtext(desc)
    best, matched = "", ""
    for kw, label in mapping.items():
        if kw in u and len(kw) > len(best):
            best, matched = kw, label
    return matched if best else "Uncategorized"


def normalize_merchant(desc):
    """Display/grouping only. 'Good enough' is fine — never used for matching."""
    s = desc.upper()
    for pat in PREFIX_PATTERNS:
        new = re.sub(pat, "", s)
        if new != s:
            s = new
            break
    s = re.sub(r"^\s*(" + "|".join(PROCESSORS) + r")\s*\*\s*", "", s)
    s = re.sub(r"\b(" + "|".join(PROCESSORS) + r")\s*\*\s*", "", s)
    s = re.sub(r"\bSTORE\b", " ", s)
    s = s.replace("#", " ")
    s = re.sub(r"\b\d{3,}\b", " ", s)
    s = re.sub(r"\b\d+\.\d{2}\s*(US|USD|EU|EUR|CAD)?\b", " ", s)
    s = re.sub(r"@\s*[\d.]+", " ", s)
    s = re.sub(r"\$[\d.]*", " ", s)
    s = re.sub(r",?\s*\b(" + "|".join(PROVINCES) + r")\b\.?\s*", " ", s)
    s = re.sub(r"\b(" + "|".join(COMMON_CITIES) + r")\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,.-*/")
    if not s:
        s = _mtext(desc)[:40] or "Unknown"
    return s

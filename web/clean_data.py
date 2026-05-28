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

import csv
import os
import re
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))

# ── Flow detection: if any of these appears in the raw text -> "Transfer" ──
MOVE_PATTERNS = [
    "E-TRANSFER", "E-TRANSFERT", "INTERNET TRANSFER", "MEMO - TRANSFER",
    "FULFILL REQUEST", "PAYMENT THANK YOU", "PAIEMEN", "TO CARD",
    "ELECTRONIC FUNDS TRANSFER", "DEPOSIT", "FEE REBATE",
]

# ── The label vocabularies (also used to populate the UI dropdowns) ──
SPENDING_CATEGORIES = [
    # Food & drink
    "Food_Groceries", "Food_DiningOut", "Food_Cafe", "Food_Delivery", "Food_AlcoholBars",
    # Shopping
    "Shopping_General", "Shopping_Clothing", "Shopping_Beauty",
    "Shopping_Electronics", "Shopping_HomeGoods",
    # Transport & travel
    "Transport_Daily", "Transport_Rideshare", "Transport_Fuel", "Transport_Parking",
    "Travel_Flights", "Travel_Accommodation", "Travel_Leisure",
    # Bills & housing
    "Bills_Rent", "Bills_Utilities", "Bills_PhoneInternet", "Bills_Insurance",
    # Health
    "Health_Medical", "Health_Pharmacy", "Health_Fitness",
    # Lifestyle
    "Entertainment", "Subscription_Software", "Subscription_Streaming",
    "Education", "Gifts_Donations", "Pets", "Kids_Family",
    # Misc
    "Fees_Banking", "Fees_AnnualCard", "Cash_Withdrawal", "Other",
]
TRANSFER_CATEGORIES = [
    "Transfer_Allowance", "Transfer_Reimbursed", "Transfer_SplitSent",
    "Transfer_Income", "Transfer_OwnAccount", "Transfer_Investment",
    "Transfer_CardPayment",
    "Income_Salary", "Income_PartTime", "Income_SideHustle",
    "Income_Rewards", "Income_Financial", "Rebate_CardFee", "Other",
]
# Labels that represent money flowing IN to the user (for the dashboard).
INCOME_LABELS = {
    "Income_Salary", "Income_PartTime", "Income_SideHustle",
    "Income_Rewards", "Income_Financial", "Rebate_CardFee",
    "Transfer_Allowance", "Transfer_Income", "Transfer_Reimbursed",
}

# ── Default dictionaries shipped with the code (the user overrides via CSV) ──
DEFAULT_CATEGORIES = {
    # Food_Groceries
    "WALMART": "Food_Groceries", "WAL-MART": "Food_Groceries", "FOODMART": "Food_Groceries",
    "TERRA FOOD": "Food_Groceries", "NO FRILLS": "Food_Groceries", "COSTCO": "Food_Groceries",
    "SUPERSTORE": "Food_Groceries", "TASC": "Food_Groceries", "LOBLAW": "Food_Groceries",
    "METRO": "Food_Groceries", "SOBEYS": "Food_Groceries", "FRESHCO": "Food_Groceries",
    "FOOD BASICS": "Food_Groceries", "T&T": "Food_Groceries", "H MART": "Food_Groceries",
    "GALLERIA": "Food_Groceries", "FARM BOY": "Food_Groceries", "SAFEWAY": "Food_Groceries",
    "SUPER C": "Food_Groceries", "SAVE-ON-FOODS": "Food_Groceries", "SAVE ON FOODS": "Food_Groceries",
    "IGA": "Food_Groceries", "FOODLAND": "Food_Groceries", "LONGO": "Food_Groceries",
    "GIANT TIGER": "Food_Groceries", "ZEHRS": "Food_Groceries",
    "NOFRILLS": "Food_Groceries", "'S NF": "Food_Groceries",  # No Frills franchises: "<owner>'S NF"
    # Food_Cafe
    "TIM HORTONS": "Food_Cafe", "STARBUCKS": "Food_Cafe", "CAFE": "Food_Cafe",
    "CREPE": "Food_Cafe", "TEN REN": "Food_Cafe", "NAYAX": "Food_Cafe",
    "SECOND CUP": "Food_Cafe", "COFFEE": "Food_Cafe", "BUBBLE TEA": "Food_Cafe",
    "CHATIME": "Food_Cafe", "GONG CHA": "Food_Cafe",
    # Food_DiningOut
    "MCDONALD": "Food_DiningOut", "EARLS": "Food_DiningOut", "PAK PAK": "Food_DiningOut",
    "YAMA": "Food_DiningOut", "JUSFRES": "Food_DiningOut", "KITCHEN": "Food_DiningOut",
    "PIZZA": "Food_DiningOut", "SUSHI": "Food_DiningOut", "RESTAURANT": "Food_DiningOut",
    "KIM'S KOREAN": "Food_DiningOut", "ROL SAN": "Food_DiningOut", "BURGER": "Food_DiningOut",
    "SUBWAY": "Food_DiningOut", "A&W": "Food_DiningOut", "KFC": "Food_DiningOut",
    "WENDY": "Food_DiningOut", "POPEYE": "Food_DiningOut", "CHIPOTLE": "Food_DiningOut",
    "OSMOW": "Food_DiningOut", "RAMEN": "Food_DiningOut",
    # Food_Delivery
    "UBER EATS": "Food_Delivery", "UBR* EATS": "Food_Delivery", "DOORDASH": "Food_Delivery",
    "SKIPTHEDISHES": "Food_Delivery", "SKIP THE": "Food_Delivery", "FANTUAN": "Food_Delivery",
    "INSTACART": "Food_Delivery",
    # Food_AlcoholBars
    "LCBO": "Food_AlcoholBars", "BEER STORE": "Food_AlcoholBars", "BAR BIRU": "Food_AlcoholBars",
    "BREWERY": "Food_AlcoholBars", "WINE": "Food_AlcoholBars",
    # Shopping_General
    "TEMU": "Shopping_General", "AMZN": "Shopping_General", "AMAZON": "Shopping_General",
    "MINISO": "Shopping_General", "KIOKII": "Shopping_General", "DOLLARAMA": "Shopping_General",
    "CHAPTERS": "Shopping_General", "INDIGO": "Shopping_General", "CIRCLE K": "Shopping_General",
    "WINNERS": "Shopping_General",
    # Shopping_Clothing
    "SHEIN": "Shopping_Clothing", "URBAN PLANET": "Shopping_Clothing", "DICKIES": "Shopping_Clothing",
    "ARITZIA": "Shopping_Clothing", "COACH": "Shopping_Clothing", "UNIQLO": "Shopping_Clothing",
    "H&M": "Shopping_Clothing", "ZARA": "Shopping_Clothing", "LULULEMON": "Shopping_Clothing",
    "OLD NAVY": "Shopping_Clothing", "NIKE": "Shopping_Clothing", "ADIDAS": "Shopping_Clothing",
    "GAP": "Shopping_Clothing", "ABERCROMBIE": "Shopping_Clothing", "GARAGE": "Shopping_Clothing",
    "HOLLISTER": "Shopping_Clothing", "FOREVER 21": "Shopping_Clothing", "MARK'S": "Shopping_Clothing",
    "SPORTING LIFE": "Shopping_Clothing", "SIMONS": "Shopping_Clothing", "BANANA REPUBLIC": "Shopping_Clothing",
    "ROOTS": "Shopping_Clothing", "BROWNS SHOES": "Shopping_Clothing", "ALDO": "Shopping_Clothing",
    # Shopping_Beauty
    "SEPHORA": "Shopping_Beauty", "HAIR SHINE": "Shopping_Beauty", "MAC COSMETICS": "Shopping_Beauty",
    "ULTA": "Shopping_Beauty", "SALON": "Shopping_Beauty", "NAILS": "Shopping_Beauty",
    "BARBER": "Shopping_Beauty",
    # Shopping_Electronics
    "BEST BUY": "Shopping_Electronics", "APPLE STORE": "Shopping_Electronics",
    "CANADA COMPUTERS": "Shopping_Electronics", "THE SOURCE": "Shopping_Electronics",
    "NEWEGG": "Shopping_Electronics",
    # Shopping_HomeGoods
    "IKEA": "Shopping_HomeGoods", "HOME DEPOT": "Shopping_HomeGoods",
    "CANADIAN TIRE": "Shopping_HomeGoods", "HOMESENSE": "Shopping_HomeGoods",
    "BED BATH": "Shopping_HomeGoods", "STRUCTUBE": "Shopping_HomeGoods",
    # Education
    "UTM": "Education", "U OF T": "Education", "UNIVERSITY OF TORONTO": "Education",
    "BOOKSTORE": "Education", "RECREATION AT U OF T": "Education", "LCC": "Education",
    "TRINITY": "Education", "TUITION": "Education", "UDEMY": "Education", "COURSERA": "Education",
    # Subscription_Software
    "OPENAI": "Subscription_Software", "CHATGPT": "Subscription_Software",
    "GOOGLE": "Subscription_Software", "CLAUDE.AI": "Subscription_Software",
    "ANTHROPIC": "Subscription_Software", "CLOUD": "Subscription_Software",
    "APPLE.COM": "Subscription_Software", "4KDOWNLOAD": "Subscription_Software",
    "MICROSOFT": "Subscription_Software", "ADOBE": "Subscription_Software",
    "GITHUB": "Subscription_Software", "NOTION": "Subscription_Software",
    "DROPBOX": "Subscription_Software",
    # Subscription_Streaming
    "NETFLIX": "Subscription_Streaming", "SPOTIFY": "Subscription_Streaming",
    "DISNEY": "Subscription_Streaming", "CRAVE": "Subscription_Streaming",
    "YOUTUBE PREMIUM": "Subscription_Streaming", "PRIME VIDEO": "Subscription_Streaming",
    "APPLE MUSIC": "Subscription_Streaming", "HBO": "Subscription_Streaming",
    # Entertainment
    "CINEPLEX": "Entertainment", "TICKETMASTER": "Entertainment", "EVENTBRITE": "Entertainment",
    "STEAM": "Entertainment", "NINTENDO": "Entertainment", "PLAYSTATION": "Entertainment",
    "XBOX": "Entertainment",
    # Travel_Flights
    "UNITED": "Travel_Flights", "ASIANA": "Travel_Flights", "AIR CANADA": "Travel_Flights",
    "WESTJET": "Travel_Flights", "FLAIR": "Travel_Flights", "KOREAN AIR": "Travel_Flights",
    "EXPEDIA": "Travel_Flights",
    # Travel_Accommodation
    "AIRBNB": "Travel_Accommodation", "HOTEL": "Travel_Accommodation",
    "BOOKING.COM": "Travel_Accommodation", "MARRIOTT": "Travel_Accommodation",
    "HILTON": "Travel_Accommodation",
    # Travel_Leisure
    "VIA RAIL": "Travel_Leisure", "RATP": "Travel_Leisure", "MUSEUM": "Travel_Leisure",
    # Transport_Daily
    "PRESTO": "Transport_Daily", "GO TRANSIT": "Transport_Daily", "TTC": "Transport_Daily",
    "TRANSIT": "Transport_Daily",
    # Transport_Rideshare
    "UBER": "Transport_Rideshare", "UBR*": "Transport_Rideshare", "LYFT": "Transport_Rideshare",
    # Transport_Fuel
    "PETRO": "Transport_Fuel", "ESSO": "Transport_Fuel", "SHELL": "Transport_Fuel",
    "PETRO-CANADA": "Transport_Fuel", "ULTRAMAR": "Transport_Fuel", "HUSKY": "Transport_Fuel",
    # Transport_Parking
    "PEARSON PARKING": "Transport_Parking", "PARKING": "Transport_Parking",
    "GREEN P": "Transport_Parking", "IMPARK": "Transport_Parking",
    # Bills_Utilities
    "HYDRO": "Bills_Utilities", "ENBRIDGE": "Bills_Utilities", "TORONTO HYDRO": "Bills_Utilities",
    "UTILITY": "Bills_Utilities", "ALECTRA": "Bills_Utilities", "HYDRO ONE": "Bills_Utilities",
    "EPCOR": "Bills_Utilities", "FORTISBC": "Bills_Utilities", "UNION GAS": "Bills_Utilities",
    "REGION OF": "Bills_Utilities", "CITY OF": "Bills_Utilities", "ELECTRIC": "Bills_Utilities",
    # Bills_PhoneInternet
    "ROGERS": "Bills_PhoneInternet", "BELL CANADA": "Bills_PhoneInternet",
    "BELL MOBILITY": "Bills_PhoneInternet", "TELUS": "Bills_PhoneInternet",
    "FIDO": "Bills_PhoneInternet", "FREEDOM MOBILE": "Bills_PhoneInternet",
    "KOODO": "Bills_PhoneInternet", "VIRGIN": "Bills_PhoneInternet", "CHATR": "Bills_PhoneInternet",
    "PUBLIC MOBILE": "Bills_PhoneInternet", "LUCKY MOBILE": "Bills_PhoneInternet",
    "TEKSAVVY": "Bills_PhoneInternet", "DISTRIBUTEL": "Bills_PhoneInternet",
    "VIDEOTRON": "Bills_PhoneInternet", "SHAW": "Bills_PhoneInternet", "EASTLINK": "Bills_PhoneInternet",
    # Bills_Insurance
    "INSURANCE": "Bills_Insurance", "INTACT": "Bills_Insurance", "SUN LIFE": "Bills_Insurance",
    "MANULIFE": "Bills_Insurance", "TD INSURANCE": "Bills_Insurance", "BELAIRDIRECT": "Bills_Insurance",
    "BELAIR DIRECT": "Bills_Insurance", "ALLSTATE": "Bills_Insurance", "AVIVA": "Bills_Insurance",
    "DESJARDINS INSUR": "Bills_Insurance", "CAA INSUR": "Bills_Insurance", "SONNET": "Bills_Insurance",
    "ECONOMICAL": "Bills_Insurance", "WAWANESA": "Bills_Insurance", "COOPERATORS": "Bills_Insurance",
    # Bills_Rent
    "RENT": "Bills_Rent", "PROPERTY MANAGEMENT": "Bills_Rent",
    # Health_Pharmacy
    "SHOPPERS DRUG": "Health_Pharmacy", "REXALL": "Health_Pharmacy", "PHARMACY": "Health_Pharmacy",
    "PHARMA": "Health_Pharmacy",
    # Health_Medical
    "ARC EYECARE": "Health_Medical", "DENTAL": "Health_Medical", "DENTIST": "Health_Medical",
    "CLINIC": "Health_Medical", "PHYSIO": "Health_Medical", "OPTICAL": "Health_Medical",
    # Health_Fitness
    "GOODLIFE": "Health_Fitness", "FITNESS": "Health_Fitness", "GYM": "Health_Fitness",
    "YOGA": "Health_Fitness",
    # Pets
    "PETSMART": "Pets", "PET VALU": "Pets", "PETCO": "Pets", "VETERINARY": "Pets",
    # Gifts_Donations
    "GOFUNDME": "Gifts_Donations", "DONATION": "Gifts_Donations", "RED CROSS": "Gifts_Donations",
    # Fees_AnnualCard (the yearly card fee — kept separate so it's visible)
    "ANNUAL FEE": "Fees_AnnualCard", "ANNUAL MEMBERSHIP": "Fees_AnnualCard",
    "CARD FEE": "Fees_AnnualCard", "PRIMARY CARD ANNUAL": "Fees_AnnualCard",
    # Fees_Banking / Cash
    "SERVICE CHARGE": "Fees_Banking", "NETWORK TRANSACTION FEE": "Fees_Banking",
    "FX CASH": "Fees_Banking", "OVERLIMIT": "Fees_Banking", "NSF": "Fees_Banking",
    "ATM WITHDRAWAL": "Cash_Withdrawal", "ABM WITHDRAWAL": "Cash_Withdrawal",
    "CASH WITHDRAWAL": "Cash_Withdrawal", "CASH ADVANCE": "Cash_Withdrawal",
    "WITHDRAWAL": "Cash_Withdrawal",
}

# Only the mechanical movements are pre-filled. Named e-transfers / deposits
# are left blank so the user picks (Allowance / SplitSent / Income / ...).
DEFAULT_TRANSFERS = {
    "PAYMENT THANK YOU": "Transfer_CardPayment", "PAIEMEN": "Transfer_CardPayment",
    "TO CARD": "Transfer_CardPayment",
    "INTERNET TRANSFER": "Transfer_OwnAccount", "MEMO - TRANSFER": "Transfer_OwnAccount",
    "FEE REBATE": "Rebate_CardFee",
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
def read_bank_csv(path, source):
    """Read a CIBC or BMO transaction export into date/description/debit/credit.

    Auto-detects the bank: BMO files carry a header row (often after a preamble
    line), CIBC files have neither. Both are parsed with the csv module so
    ragged rows and quoted commas can't break the import."""
    rows = _read_rows(path)
    hdr = _bmo_header_index(rows)
    if hdr is not None:
        return _parse_bmo(rows, hdr, source)
    return _parse_cibc(rows, source)


def _read_rows(path):
    """The CSV as a list of cell-lists, accepting a path or a file-like."""
    if hasattr(path, "read"):           # file-like (StringIO from the website)
        return list(csv.reader(path))
    with open(path, newline="") as f:   # filesystem path (local CLI)
        return list(csv.reader(f))


def _clean(s):
    """Trim whitespace and one surrounding pair of quotes (BMO wraps cells)."""
    s = (s or "").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "'\"":
        s = s[1:-1].strip()
    return s


def _parse_cibc(rows, source):
    """CIBC: no header; columns are date, description, debit, credit, [card]."""
    records = []
    for cells in rows:
        if not cells or not _clean(cells[0]):
            continue
        records.append({
            "date": _clean(cells[0]),
            "description": _clean(cells[1]) if len(cells) > 1 else "",
            "debit": _to_float(cells[2]) if len(cells) > 2 else 0.0,   # money out
            "credit": _to_float(cells[3]) if len(cells) > 3 else 0.0,  # money in
            "source": source,
        })
    return pd.DataFrame(records)


def _bmo_header_index(rows):
    """Index of the BMO header row, or None if this isn't a BMO export.

    BMO has two layouts, both with a 'Description' column:
      - credit card : Item #, Card #, Transaction Date, Posting Date,
                       Transaction Amount, Description
      - chequing    : First Bank Card, Transaction Type, Date Posted,
                       Transaction Amount, Description
    """
    for i, cells in enumerate(rows):
        cleaned = [_clean(c).lower() for c in cells]
        if "description" in cleaned and (
                "transaction date" in cleaned or "date posted" in cleaned
                or "transaction type" in cleaned):
            return i
    return None


def _parse_bmo(rows, hdr, source):
    """BMO. The amount sign convention differs by product, so the direction is
    taken from the Transaction Type column (DR/CR) when present (chequing);
    otherwise from the amount sign (credit card: purchase +, payment -)."""
    header = [_clean(c).lower() for c in rows[hdr]]

    def col(*needles):
        for idx, name in enumerate(header):
            if all(n in name for n in needles):
                return idx
        return None

    i_date = col("transaction", "date")     # credit card
    if i_date is None:
        i_date = col("date")                # chequing: "Date Posted"
    i_amt = col("amount")
    if i_amt is None:                       # amount column labelled just "Transaction"
        for idx, name in enumerate(header):
            if "transaction" in name and "type" not in name and "date" not in name:
                i_amt = idx
                break
    i_desc = col("description")
    i_type = col("type")                    # "Transaction Type" (DR/CR) on chequing
    if i_date is None or i_amt is None:
        return pd.DataFrame()
    need = max(i for i in (i_date, i_amt, i_desc, i_type) if i is not None)

    records = []
    for cells in rows[hdr + 1:]:
        if len(cells) <= need:
            continue
        date = _clean(cells[i_date])
        if not date:
            continue
        amt = _to_float(cells[i_amt])
        ttype = _clean(cells[i_type]).upper() if i_type is not None else ""
        if "CR" in ttype:                       # chequing credit = money in
            debit, credit = 0.0, abs(amt)
        elif "DR" in ttype:                     # chequing debit = money out
            debit, credit = abs(amt), 0.0
        else:                                   # credit card: + purchase, - payment
            debit = amt if amt > 0 else 0.0
            credit = -amt if amt < 0 else 0.0
        records.append({
            "date": date,
            "description": _clean(cells[i_desc]) if i_desc is not None else "",
            "debit": debit,
            "credit": credit,
            "source": source,
        })
    return pd.DataFrame(records)


def _to_float(s):
    s = _clean(s).replace("$", "").replace(",", "")
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

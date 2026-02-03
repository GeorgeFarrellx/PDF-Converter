import re
import datetime as _dt
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Optional

import pdfplumber


# ----------------------------
# Helpers
# ----------------------------

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

# Start of a transaction row (NatWest export)
# Example: "31 Mar 2025 D/D UK FUELS LTD , 0111... £443.98 £24,746.07"
_TXN_START_RE = re.compile(
    r"^\s*(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\s+(?:(?P<year>\d{4})\s+)?(?P<type>[A-Z/]{2,4})\b"
)

# Currency amounts (NatWest export uses £, sometimes appears like "£3,088.88")
_MONEY_RE = re.compile(r"£\s*\(?-?[\d,]+\.\d{2}\)?")

# Header/footer noise lines to ignore
_IGNORE_LINE_RE_LIST = [
    re.compile(r"^\s*Page\s+\d+\s+of\s+\d+\s*$", re.IGNORECASE),
    re.compile(r"^\s*Date\s+Type\s+Description\s+Paid\s+in\s+Paid\s+out\s+Balance\s*$", re.IGNORECASE),
    re.compile(r"^\s*Transactions\s*$", re.IGNORECASE),
    re.compile(r"^\s*©\s*National\s+Westminster\s+Bank\b", re.IGNORECASE),
    re.compile(r"^\s*National\s+Westminster\s+Bank\b", re.IGNORECASE),
    re.compile(r"^\s*Authorised\s+by\s+the\s+Prudential\b", re.IGNORECASE),
]

_PERIOD_RE = re.compile(
    r"Showing:\s*(?P<d1>\d{2})\s+(?P<m1>[A-Za-z]{3})\s+(?P<y1>\d{4})\s+to\s+(?P<d2>\d{2})\s+(?P<m2>[A-Za-z]{3})\s+(?P<y2>\d{4})",
    re.IGNORECASE
)

_ACCOUNT_NAME_RE = re.compile(r"^\s*Account\s+name:\s*(?P<name>.+?)\s*$", re.IGNORECASE)


def _is_ignorable_line(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return True
    for rx in _IGNORE_LINE_RE_LIST:
        if rx.search(s):
            return True
    return False


def _parse_money(s: str) -> Optional[float]:
    if s is None:
        return None
    t = s.strip()
    if not t:
        return None
    neg = False
    if "(" in t and ")" in t:
        neg = True
        t = t.replace("(", "").replace(")", "")
    t = t.replace("£", "").replace(",", "").strip()
    if not t:
        return None
    try:
        val = Decimal(t)
    except InvalidOperation:
        return None
    if neg:
        val = -val
    # Always return float for downstream Excel
    return float(val)


def _title_case_keep_slashes(s: str) -> str:
    # Title-case but preserve internal slashes tokens like "D/D" if present (we usually map anyway).
    parts = re.split(r"(\s+)", s.strip())
    out = []
    for p in parts:
        if p.isspace() or p == "":
            out.append(p)
        else:
            out.append(p[:1].upper() + p[1:].lower())
    return "".join(out).strip()


def _map_natwest_type(raw_type: str) -> str:
    t = (raw_type or "").strip().upper()
    mapping = {
        "D/D": "Direct Debit",
        "BAC": "Bacs",
        "DPC": "Faster Payment",
        "CHG": "Charge",
    }
    return mapping.get(t, _title_case_keep_slashes(raw_type or ""))


def _clean_description(desc: str) -> str:
    # Collapse whitespace and tidy punctuation spacing a bit, keep commas as-is.
    if desc is None:
        return ""
    s = re.sub(r"\s+", " ", desc).strip()
    # Remove spaces before commas
    s = re.sub(r"\s+,", ",", s)
    # Remove duplicated commas spacing " , " -> ", "
    s = re.sub(r",\s*", ", ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


def _apply_global_transaction_type_rules(txn: Dict) -> Dict:
    """
    Mutates txn fields according to the GLOBAL TRANSACTION TYPE RULES.
    """
    ttype = (txn.get("Transaction Type") or "").strip()
    desc = (txn.get("Description") or "").strip()

    desc_lower = desc.lower()

    # Returned Direct Debit rule
    if desc_lower.startswith("returned direct debit"):
        txn["Transaction Type"] = "Direct Debit"
        # Ensure prefix is preserved exactly once
        if not desc.startswith("Returned Direct Debit"):
            # normalise casing but keep prefix
            rest = desc[len("returned direct debit"):].lstrip() if desc_lower.startswith("returned direct debit") else desc
            txn["Description"] = ("Returned Direct Debit " + rest).strip()
        else:
            txn["Description"] = desc

        return txn

    # ApplePay / Clearpay / Contactless / endswith GB -> Card Payment
    if ("applepay" in desc_lower) or ("clearpay" in desc_lower) or ("contactless" in desc_lower) or desc.endswith("GB"):
        txn["Transaction Type"] = "Card Payment"
    else:
        # Otherwise keep bank wording but Title Case
        txn["Transaction Type"] = _title_case_keep_slashes(ttype)

    # Remove type prefix from Description (except Returned Direct Debit case already returned)
    # We attempt a gentle removal if description literally starts with the type label.
    norm_type = (txn.get("Transaction Type") or "").strip()
    if norm_type:
        if desc_lower.startswith(norm_type.lower() + " "):
            desc = desc[len(norm_type):].lstrip()
        elif desc_lower == norm_type.lower():
            desc = ""

    txn["Description"] = desc.strip()
    return txn


def _infer_year_for_missing_year(
    day: int,
    mon_num: int,
    last_seen_mon: Optional[int],
    current_year: int,
    period_start_year: Optional[int],
    period_end_year: Optional[int],
) -> int:
    """
    For NatWest export, transactions are typically listed in reverse chronological order.
    If year is missing on a row, infer it using the statement period and month rollovers.
    """
    # Default to end-year if available
    year = current_year
    if period_end_year is not None:
        year = current_year

    # If moving *backwards* in time (down the PDF),
    # when month jumps upwards (e.g. Jan -> Dec), we crossed into previous year.
    if last_seen_mon is not None:
        if mon_num > last_seen_mon:
            year -= 1

    # Clamp to period range if known
    if period_start_year is not None and period_end_year is not None:
        if year < period_start_year:
            year = period_start_year
        if year > period_end_year:
            year = period_end_year

    return year


def _extract_period_years(all_text: str) -> (Optional[int], Optional[int]):
    if not all_text:
        return None, None
    m = _PERIOD_RE.search(all_text)
    if not m:
        return None, None
    try:
        y1 = int(m.group("y1"))
        y2 = int(m.group("y2"))
        return y1, y2
    except Exception:
        return None, None


# ----------------------------
# Required API
# ----------------------------

def extract_transactions(pdf_path) -> List[Dict]:
    """
    NatWest 'online transactions service' export parser.
    Returns transactions in the order they appear in the PDF (typically newest -> oldest).
    """
    transactions: List[Dict] = []

    # Read all text once for period inference and account name scanning if needed
    all_text_chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            all_text_chunks.append(txt)
    all_text = "\n".join(all_text_chunks)

    period_start_year, period_end_year = _extract_period_years(all_text)

    current_block = None  # dict with parsed header + lines
    last_seen_mon = None
    current_year = period_end_year if period_end_year is not None else _dt.date.today().year

    def finalize_block(block):
        if not block:
            return

        raw_type = block["raw_type"]
        date_obj = block["date"]
        desc_lines = block["desc_lines"]

        block_text = " ".join([x.strip() for x in desc_lines if x is not None]).strip()

        # Extract money values across the block
        monies = _MONEY_RE.findall(block_text)
        monies = [m.strip() for m in monies if m and m.strip()]

        amount_val = None
        balance_val = None

        if len(monies) >= 2:
            # Usually: "... £AMOUNT £BALANCE"
            amount_val = _parse_money(monies[-2])
            balance_val = _parse_money(monies[-1])
        elif len(monies) == 1:
            amount_val = _parse_money(monies[-1])
            balance_val = None
        else:
            # If we can't find any currency values, treat as non-transaction block.
            return

        # Remove currency tokens from description area
        desc_wo_money = block_text
        for m in monies:
            desc_wo_money = desc_wo_money.replace(m, " ")
        desc_wo_money = _clean_description(desc_wo_money)

        txn = {
            "Date": date_obj,
            "Transaction Type": _map_natwest_type(raw_type),
            "Description": desc_wo_money,
            # placeholder, may be overwritten by balance-delta derivation:
            "Amount": float(amount_val) if amount_val is not None else 0.0,
            "Balance": float(balance_val) if balance_val is not None else None,
            # keep raw_type for internal heuristics
            "_raw_type": raw_type,
        }
        transactions.append(txn)

    # Parse line-by-line, building blocks
    for raw_line in all_text.splitlines():
        line = raw_line.rstrip("\n")
        if _is_ignorable_line(line):
            continue

        m = _TXN_START_RE.match(line)
        if m:
            # Start new transaction block
            finalize_block(current_block)

            day = int(m.group("day"))
            mon = (m.group("mon") or "").strip().lower()
            mon_num = _MONTHS.get(mon[:3], None)
            raw_type = (m.group("type") or "").strip()
            year_str = m.group("year")
            year = None

            if mon_num is None:
                # If month can't be parsed, skip this as a malformed row
                current_block = None
                continue

            if year_str:
                try:
                    year = int(year_str)
                except Exception:
                    year = None

            if year is None:
                year = _infer_year_for_missing_year(
                    day=day,
                    mon_num=mon_num,
                    last_seen_mon=last_seen_mon,
                    current_year=current_year,
                    period_start_year=period_start_year,
                    period_end_year=period_end_year,
                )
            # Update month/year tracking for missing-year inference (reverse chronological list)
            last_seen_mon = mon_num
            current_year = year

            try:
                date_obj = _dt.date(year, mon_num, day)
            except Exception:
                current_block = None
                continue

            # Remove leading "DD Mon [YYYY] TYPE" from line and keep remainder as description line 1
            prefix_len = m.end()
            remainder = line[prefix_len:].strip()
            desc_lines = []
            if remainder:
                desc_lines.append(remainder)

            current_block = {
                "date": date_obj,
                "raw_type": raw_type,
                "desc_lines": desc_lines,
            }
        else:
            # Continuation line for current transaction block
            if current_block is None:
                continue
            # Skip repeated header fragments that sometimes appear mid-page
            if _is_ignorable_line(line):
                continue
            current_block["desc_lines"].append(line.strip())

    finalize_block(current_block)

    # If there are no transactions, return empty list
    if not transactions:
        return []

    # Derive signed Amounts from balance deltas when possible (preferred for reconciliation).
    # The PDF is typically reverse chronological (newest -> oldest):
    # amount_i = balance_i - balance_{i+1}
    for i in range(len(transactions) - 1):
        b0 = transactions[i].get("Balance")
        b1 = transactions[i + 1].get("Balance")
        if isinstance(b0, (int, float)) and isinstance(b1, (int, float)):
            delta = round(float(b0) - float(b1), 2)
            transactions[i]["Amount"] = delta

    # Heuristic sign for final (oldest) transaction if we couldn't delta-derive it.
    # This only affects the last row because it has no following balance.
    last = transactions[-1]
    if last is not None:
        raw_type = (last.get("_raw_type") or "").strip().upper()
        desc = (last.get("Description") or "")
        amt = last.get("Amount")
        if isinstance(amt, (int, float)):
            amt_abs = abs(float(amt))
            sign = None

            if "From A/C" in desc or "FROM A/C" in desc:
                sign = +1
            elif "To A/C" in desc or "TO A/C" in desc:
                sign = -1
            elif raw_type in {"BAC"}:
                sign = +1
            elif raw_type in {"D/D", "CHG"}:
                sign = -1

            if sign is not None:
                last["Amount"] = round(sign * amt_abs, 2)
            else:
                # Default to negative for safety (many last-row items are payments out),
                # but keep as-is if already negative.
                last["Amount"] = round(float(amt), 2)

    # Apply global transaction type rules + final description cleanup
    cleaned = []
    for txn in transactions:
        # Remove internal key
        txn.pop("_raw_type", None)
        txn["Description"] = _clean_description(txn.get("Description") or "")
        txn = _apply_global_transaction_type_rules(txn)
        cleaned.append(txn)

    return cleaned


def extract_statement_balances(pdf_path) -> Dict[str, Optional[float]]:
    """
    NatWest export does not show explicit 'Start balance / End balance' summary in this format.
    We derive:
      - end_balance = balance after the latest transaction (first transaction in the PDF list)
      - start_balance = balance before the earliest transaction (computed from oldest balance - oldest amount)
    If balances are missing, returns None appropriately.
    """
    txns = extract_transactions(pdf_path)
    if not txns:
        return {"start_balance": None, "end_balance": None}

    # end_balance: balance of first row with a balance
    end_balance = None
    for t in txns:
        b = t.get("Balance")
        if isinstance(b, (int, float)):
            end_balance = float(b)
            break

    # start_balance: for the oldest txn, balance_before = balance_after - amount
    start_balance = None
    oldest = None
    for t in reversed(txns):
        b = t.get("Balance")
        a = t.get("Amount")
        if isinstance(b, (int, float)) and isinstance(a, (int, float)):
            oldest = t
            start_balance = round(float(b) - float(a), 2)
            break

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_account_holder_name(pdf_path) -> str:
    """
    Best-effort extraction of the client/account name.
    NatWest export includes: "Account name: <NAME>"
    """
    with pdfplumber.open(pdf_path) as pdf:
        # Usually on first page
        for page in pdf.pages[:2]:
            txt = page.extract_text() or ""
            for line in txt.splitlines():
                m = _ACCOUNT_NAME_RE.match(line.strip())
                if m:
                    name = (m.group("name") or "").strip()
                    # Avoid generic headings
                    if name and name.lower() not in {"transactions"}:
                        return name

    return ""

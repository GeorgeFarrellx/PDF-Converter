# Version: hsbc-1.4.py
import os
import re
from datetime import date, datetime
import pdfplumber


# --- Regex / constants ---

DATE_RE_FULL = re.compile(r"^(?P<d>\d{2})\s+(?P<mon>[A-Za-z]{3})\s+(?P<yy>\d{2})\b")
DATE_RE_SHORT = re.compile(r"^(?P<d>\d{2})\s+(?P<mon>[A-Za-z]{3})\b")
MONEY_RE = re.compile(r"£?\d{1,3}(?:,\d{3})*\.\d{2}")

# HSBC statements show a table column called "Payment type"
# Common codes seen in your samples: DD, CR, DR, VIS, ATM, BP, )))
ROW_CODE_RE = re.compile(r"^(DR|DD|CR|VIS|ATM|BP|CHQ|SO|BGC|FPI)\b")


MONTHS_3 = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}

# Full month names sometimes appear in the statement period header
MONTHS_FULL = {
    "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
    "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12,
}

# Optional: map HSBC codes to “friendly” types (still overridden by global rules below)
HSBC_TYPE_MAP = {
    "DD": "Direct Debit",
    "CR": "Credit",
    "DR": "Debit",
    "VIS": "Card Payment",
    "ATM": "Cash Withdrawal",
    "BP": "Bank Payment",
    "CHQ": "Cheque",
    "SO": "Standing Order",
    "BGC": "Credit",
    "FPI": "Credit",
    ")))": "Card Payment",
}


# --- Helpers ---

def _to_float(money_text: str) -> float:
    return float(money_text.replace("£", "").replace(",", ""))


def _money_from_cell(cell_text: str):
    if not cell_text:
        return None
    m = MONEY_RE.search(cell_text)
    return _to_float(m.group(0)) if m else None


def _normalise_digit_splits_in_line(s: str) -> str:
    # HSBC PDFs sometimes split digits inside amounts with a single space (e.g. "53,903.1 8").
    # Remove single spaces inside numbers but keep multi-space column gaps intact.
    s = re.sub(r"(?<=\d) (?=[\d,.])", "", s)
    s = re.sub(r"(?<=[,\.]) (?=\d)", "", s)
    return s


def _extract_amount_and_balance_from_line(line: str, paid_out_idx: int, paid_in_idx: int, balance_idx: int):
    """
    Robustly pull the transaction amount (signed) and optional balance from a table line.
    """
    line_n = _normalise_digit_splits_in_line(line)

    slack_start = max(0, paid_out_idx - 10)
    numeric_region = line_n[slack_start:]
    matches = [
        (slack_start + m.start(), slack_start + m.end(), _to_float(m.group(0)))
        for m in MONEY_RE.finditer(numeric_region)
    ]

    if not matches:
        return None, None

    if len(matches) >= 2:
        balance = matches[-1][2]
        amt_match = matches[-2]
    else:
        balance = None
        amt_match = matches[-1]

    amt_val = amt_match[2]
    signed_amount = -amt_val if amt_match[1] <= paid_in_idx else amt_val

    return signed_amount, balance


def _is_balance_label(text: str) -> bool:
    t = re.sub(r"\s+", "", (text or "")).upper()
    return ("BALANCEBROUGHTFORWARD" in t) or ("BALANCECARRIEDFORWARD" in t)


def _is_table_header(line: str) -> bool:
    s = (line or "").lower()
    return ("payment type and details" in s) and ("paid out" in s) and ("paid in" in s) and ("balance" in s)


def _is_row_start(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    if t.startswith(")))"):
        return True
    return bool(ROW_CODE_RE.match(t))


def _parse_code_and_first_desc(text: str):
    t = (text or "").strip()
    if t.startswith(")))"):
        return ")))", t[3:].strip()
    parts = t.split(None, 1)
    code = parts[0].strip()
    first_desc = parts[1].strip() if len(parts) > 1 else ""
    return code, first_desc


def _apply_global_type_rules(base_type: str, description: str) -> str:
    desc = (description or "").strip()
    low = desc.lower()

    # Returned Direct Debit rule
    if low.startswith("returned direct debit"):
        return "Direct Debit"

    # Card Payment override rules
    if "applepay" in low:
        return "Card Payment"
    if "clearpay" in low:
        return "Card Payment"
    if "contactless" in low:
        return "Card Payment"
    if desc.endswith(" GB"):
        return "Card Payment"

    return base_type


def _infer_statement_years(page1_text: str):
    """
    Best-effort inference if a statement ever omits the year in the Date column.
    HSBC headers often include: '4 June to 3 July 2024'.
    Returns (start_year, end_year) or (None, None).
    """
    if not page1_text:
        return None, None

    t = " ".join(page1_text.split())
    # Example: "4 June to 3 July 2024"
    m = re.search(
        r"\b(?P<d1>\d{1,2})\s+(?P<m1>[A-Za-z]+)\s+to\s+(?P<d2>\d{1,2})\s+(?P<m2>[A-Za-z]+)\s+(?P<y>\d{4})\b",
        t
    )
    if not m:
        return None, None

    y = int(m.group("y"))
    m1 = MONTHS_FULL.get(m.group("m1").title())
    m2 = MONTHS_FULL.get(m.group("m2").title())
    if not m1 or not m2:
        return None, None

    start_year = y if m1 <= m2 else (y - 1)
    end_year = y
    return start_year, end_year


def _parse_period_from_text(page1_text: str):
    if not page1_text:
        return None, None
    t = " ".join(page1_text.split())
    m_explicit_years = re.search(
        r"\b(?P<d1>\d{1,2})\s+(?P<m1>[A-Za-z]+)\s+(?P<y1>\d{4})\s+to\s+(?P<d2>\d{1,2})\s+(?P<m2>[A-Za-z]+)\s+(?P<y2>\d{4})\b",
        t,
        re.IGNORECASE,
    )
    if m_explicit_years:
        y1 = int(m_explicit_years.group("y1"))
        y2 = int(m_explicit_years.group("y2"))
        m1 = MONTHS_FULL.get(m_explicit_years.group("m1").title())
        m2 = MONTHS_FULL.get(m_explicit_years.group("m2").title())
        if not m1 or not m2:
            return None, None
        try:
            start = date(y1, m1, int(m_explicit_years.group("d1")))
            end = date(y2, m2, int(m_explicit_years.group("d2")))
            return start, end
        except Exception:
            return None, None

    m = re.search(
        r"\b(?P<d1>\d{1,2})\s+(?P<m1>[A-Za-z]+)\s+to\s+(?P<d2>\d{1,2})\s+(?P<m2>[A-Za-z]+)\s+(?P<y>\d{4})\b",
        t,
    )
    if not m:
        return None, None
    y = int(m.group("y"))
    m1 = MONTHS_FULL.get(m.group("m1").title())
    m2 = MONTHS_FULL.get(m.group("m2").title())
    if not m1 or not m2:
        return None, None
    start_year = y if m1 <= m2 else (y - 1)
    try:
        start = date(start_year, m1, int(m.group("d1")))
        end = date(y, m2, int(m.group("d2")))
        return start, end
    except Exception:
        return None, None


def _parse_period_from_filename(pdf_path: str):
    name = os.path.basename(pdf_path or "")
    m = re.search(
        r"(?P<d1>\d{1,2})[./-](?P<m1>\d{1,2})[./-](?P<y1>\d{2,4})\s*[-–—]\s*"
        r"(?P<d2>\d{1,2})[./-](?P<m2>\d{1,2})[./-](?P<y2>\d{2,4})",
        name,
    )
    if not m:
        return None, None
    try:
        d1 = int(m.group("d1"))
        m1 = int(m.group("m1"))
        y1 = int(m.group("y1"))
        d2 = int(m.group("d2"))
        m2 = int(m.group("m2"))
        y2 = int(m.group("y2"))
        if y1 < 100:
            y1 += 2000
        if y2 < 100:
            y2 += 2000
        return date(y1, m1, d1), date(y2, m2, d2)
    except Exception:
        return None, None


def extract_statement_period(pdf_path: str):
    """Public wrapper to extract the statement coverage period (start_date, end_date)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page1 = pdf.pages[0] if pdf.pages else None
            text = page1.extract_text() if page1 else ""
        start, end = _parse_period_from_text(text or "")
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _parse_date_from_left(left_text: str, current_year: int | None, last_month: int | None):
    """
    Returns (date_obj, new_current_year, new_last_month, remaining_left_after_date)
    """
    left = (left_text or "").strip()
    if not left:
        return None, current_year, last_month, ""

    m = DATE_RE_FULL.match(left)
    if m:
        d = int(m.group("d"))
        mon = m.group("mon").title()
        yy = int(m.group("yy"))
        month = MONTHS_3.get(mon)
        if not month:
            return None, current_year, last_month, left

        dt = date(2000 + yy, month, d)
        remainder = left[m.end():].strip()
        return dt, current_year, month, remainder

    m2 = DATE_RE_SHORT.match(left)
    if m2 and current_year is not None:
        d = int(m2.group("d"))
        mon = m2.group("mon").title()
        month = MONTHS_3.get(mon)
        if not month:
            return None, current_year, last_month, left

        # year rollover: if months go backwards (e.g. Dec -> Jan), bump year
        if last_month is not None and month < last_month:
            current_year += 1

        dt = date(current_year, month, d)
        remainder = left[m2.end():].strip()
        return dt, current_year, month, remainder

    return None, current_year, last_month, left


# --- Required public functions ---

def extract_transactions(pdf_path: str):
    """
    Output rows are dicts with CAPITALISED headings:
      Date, Transaction Type, Description, Amount, Balance
    Date is a datetime.date.
    Amount/Balance are floats (Balance may be blank '').
    """
    rows = []

    current_date = None
    current_year = None
    last_month = None
    current_txn = None  # {date, code, desc_lines, amount, balance}

    def commit():
        nonlocal current_txn
        if not current_txn:
            return
        if current_txn.get("amount") is None:
            current_txn = None
            return

        # Join description lines
        desc_lines = [x for x in current_txn["desc_lines"] if x]
        description = " ".join(desc_lines).strip()
        description = re.sub(r"\s+", " ", description)

        code = current_txn.get("code", "").strip()
        base_type = HSBC_TYPE_MAP.get(code, (code.title() if code else "Transaction"))
        tx_type = _apply_global_type_rules(base_type, description)
        description = re.sub(r"\s+CD\s*\d{4}\b\s*$", "", description, flags=re.IGNORECASE).strip()

        rows.append({
            "Date": current_txn["date"],
            "Transaction Type": tx_type,
            "Description": description,
            "Amount": round(float(current_txn["amount"]), 2),
            "Balance": current_txn["balance"] if current_txn.get("balance") is not None else "",
        })
        current_txn = None

    with pdfplumber.open(pdf_path) as pdf:
        # Year inference fallback (only used if we ever see dates without a year)
        page1_text = (pdf.pages[0].extract_text() or "")
        start_year, _end_year = _infer_statement_years(page1_text)
        current_year = start_year

        for page in pdf.pages:
            text = page.extract_text(layout=True) or ""
            lines = text.splitlines()

            # Find table header and column indices per page
            header_line = None
            for ln in lines:
                if _is_table_header(ln):
                    header_line = ln
                    break
            if not header_line:
                continue

            paid_out_idx = header_line.lower().find("paid out")
            paid_in_idx = header_line.lower().find("paid in")
            balance_idx = header_line.lower().find("balance")
            if paid_out_idx < 0 or paid_in_idx < 0 or balance_idx < 0:
                continue

            # Start processing lines after the header
            started = False
            for ln in lines:
                if not started:
                    if ln == header_line:
                        started = True
                    continue

                line = (ln or "").rstrip()
                if not line.strip():
                    continue

                # stop once we hit FSCS info section (post-table content)
                if line.lower().startswith("information about the financial services compensation scheme"):
                    commit()
                    break

                left = line[:paid_out_idx].strip()
                paid_out_cell = line[paid_out_idx:paid_in_idx].strip()
                paid_in_cell = line[paid_in_idx:balance_idx].strip()
                balance_cell = line[balance_idx:].strip()

                # Skip brought forward / carried forward summary rows
                if _is_balance_label(left) or _is_balance_label(line):
                    continue

                # Parse date (if present) from left column
                parsed_date, current_year, last_month, left_after_date = _parse_date_from_left(left, current_year, last_month)
                if parsed_date:
                    current_date = parsed_date
                if current_date is None:
                    continue

                # New transaction row start?
                if left_after_date and _is_row_start(left_after_date):
                    commit()
                    code, first_desc = _parse_code_and_first_desc(left_after_date)
                    current_txn = {
                        "date": current_date,
                        "code": code,
                        "desc_lines": [first_desc] if first_desc else [],
                        "amount": None,
                        "balance": None,
                    }
                else:
                    # Continuation line (details column)
                    if current_txn and left_after_date:
                        current_txn["desc_lines"].append(left_after_date)

                # Amount appears in Paid out / Paid in column => finalise txn
                if current_txn:
                    signed_amt, bal_val = _extract_amount_and_balance_from_line(
                        line, paid_out_idx, paid_in_idx, balance_idx
                    )

                    if signed_amt is not None:
                        current_txn["amount"] = signed_amt
                        if bal_val is not None:
                            current_txn["balance"] = bal_val
                        commit()

        commit()

    return rows


def extract_statement_balances(pdf_path: str):
    """
    Extract statement start/end balances (best-effort).
    HSBC Account Summary (page 1) includes Opening Balance and Closing Balance.

    Note: HSBC PDFs sometimes split digits inside amounts (e.g. "53,903.1 8").
    We normalise digit/decimal spacing before regex extraction.
    """
    start_balance = None
    end_balance = None

    with pdfplumber.open(pdf_path) as pdf:
        page1 = pdf.pages[0].extract_text() or ""
        t = " ".join(page1.split())

        # Fix digit-splitting inside amounts (e.g. "53,903.1 8" -> "53,903.18")
        t = re.sub(r"(?<=\d)\s+(?=[\d,.])", "", t)
        t = re.sub(r"(?<=[,\.])\s+(?=\d)", "", t)

        # allow for "OpeningBalance" or "Opening Balance"
        m1 = re.search(r"Opening\s*Balance\s*(£?\d{1,3}(?:,\d{3})*\.\d{2})", t, re.IGNORECASE)
        m2 = re.search(r"Closing\s*Balance\s*(£?\d{1,3}(?:,\d{3})*\.\d{2})", t, re.IGNORECASE)

        if m1:
            start_balance = _to_float(m1.group(1))
        if m2:
            end_balance = _to_float(m2.group(1))

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_account_holder_name(pdf_path: str) -> str:
    """
    HSBC Business statements show:
      'Account Name  Sortcode  Account Number ...'
      next line: '<NAME>  40-03-33  42192047 ...'
    We return the Account Name.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page1 = pdf.pages[0].extract_text() or ""
        lines = [ln.strip() for ln in page1.splitlines() if ln.strip()]

    # Find the header line, then read the next line
    header_idx = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        if "account name" in low and "sortcode" in low and "account number" in low:
            header_idx = i
            break

    if header_idx is not None:
        for j in range(header_idx + 1, min(header_idx + 5, len(lines))):
            candidate = lines[j]
            m = re.search(r"^(?P<name>.+?)\s+\d{2}-\d{2}-\d{2}\s+\d{6,}", candidate)
            if m:
                return m.group("name").strip()

    # Fallback: first line containing sortcode pattern, take left portion
    for ln in lines:
        m = re.search(r"^(?P<name>.+?)\s+\d{2}-\d{2}-\d{2}\s+\d{6,}", ln)
        if m:
            return m.group("name").strip()

    return ""

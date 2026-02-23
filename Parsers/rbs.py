# Version: rbs.py
"""RBS / Royal Bank of Scotland (Business Current Account) PDF parser.

Expected output columns (used by Main):
- Date (datetime.date)
- Transaction Type (str)
- Description (str)
- Amount (float)  # credits positive, debits negative
- Balance (float | None)

Also provides:
- extract_statement_balances(pdf_path) -> {start_balance, end_balance}
- extract_account_holder_name(pdf_path) -> str

Notes:
- Text-based PDFs (no OCR)
- Handles multi-line descriptions (continuation lines)
- Handles pages with 'BROUGHT FORWARD' rows (ignored)
- Best-effort year inference from 'Period Covered' and year rollovers
"""

from __future__ import annotations

import os
import re
from datetime import date

import pdfplumber


# ----------------------------
# Helpers
# ----------------------------

MONTHS = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}

DATE_RE = re.compile(r"^(?P<dd>\d{1,2})\s+(?P<mon>[A-Z]{3})(?:\s+(?P<yyyy>\d{4}))?\b")
MONEY_RE = re.compile(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b")

# Transaction type prefixes seen in RBS statements
TYPE_PREFIXES = [
    "Automated Credit",
    "OnLine Transaction",
    "Card Transaction",
    "Direct Debit",
    "Standing Order",
    "Charges",
    # Missing in some statements; must be treated as new transaction rows
    "Transfer",
    "Cash Withdrawal",
]

# Which types are generally credits vs debits
CREDIT_TYPES = {
    "Automated Credit",
}
DEBIT_TYPES = {
    "OnLine Transaction",
    "Card Transaction",
    "Direct Debit",
    "Standing Order",
    "Charges",
    "Cash Withdrawal",
    # Transfer can be credit or debit; keep sign inference as primary
    "Transfer",
}

def _is_type_prefix_row(row_text: str) -> bool:
    """True if this line begins a new transaction without repeating the date."""
    s = (row_text or "").lstrip()
    for p in TYPE_PREFIXES:
        if s.lower().startswith(p.lower()):
            return True
    return s.upper().startswith("BROUGHT FORWARD")


def _is_date_row(parts: list[str]) -> tuple[bool, int, int, int | None, int]:
    """Return (is_date, dd, mm, yyyy_or_none, tokens_consumed).

    Accepts:
      - "29 JUN 2024 ..."
      - "01 JUL ..."
      - "1 JUL ..."
    """
    if len(parts) < 2:
        return False, 0, 0, None, 0

    t0 = (parts[0] or "").strip()
    t1 = (parts[1] or "").strip().upper()

    if (1 <= len(t0) <= 2) and t0.isdigit() and t1 in MONTHS:
        dd = int(t0)
        mm = MONTHS[t1]
        yyyy = None
        consumed = 2

        if len(parts) >= 3:
            t2 = (parts[2] or "").strip()
            if len(t2) == 4 and t2.isdigit():
                yyyy = int(t2)
                consumed = 3

        return True, dd, mm, yyyy, consumed

    return False, 0, 0, None, 0

def _money_to_float(s: str) -> float:
    return float(s.replace(",", ""))


def _parse_period_year(pdf_path: str) -> int | None:
    """Try to extract the statement year from the 'Period Covered' line on page 1."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None

    # Example: "Period Covered 01 JUN 2024 to 28 JUN 2024"
    m = re.search(r"Period Covered\s+\d{2}\s+[A-Z]{3}\s+(\d{4})\s+to\s+\d{2}\s+[A-Z]{3}\s+(\d{4})", text)
    if m:
        # Usually same year; if not, the start year is fine as baseline
        try:
            return int(m.group(1))
        except Exception:
            return None

    return None


def _parse_period_dates(pdf_path: str) -> tuple[date | None, date | None]:
    """Try to extract full statement period dates from the 'Period Covered' line on page 1."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None, None

    m = re.search(
        r"Period Covered\s+(?P<d1>\d{2})\s+(?P<m1>[A-Z]{3})\s+(?P<y1>\d{4})\s+to\s+(?P<d2>\d{2})\s+(?P<m2>[A-Z]{3})\s+(?P<y2>\d{4})",
        text,
    )
    if not m:
        return None, None
    try:
        m1 = MONTHS.get(m.group("m1").upper())
        m2 = MONTHS.get(m.group("m2").upper())
        if not m1 or not m2:
            return None, None
        start = date(int(m.group("y1")), m1, int(m.group("d1")))
        end = date(int(m.group("y2")), m2, int(m.group("d2")))
        return start, end
    except Exception:
        return None, None


def _parse_period_from_filename(pdf_path: str) -> tuple[date | None, date | None]:
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


def extract_statement_period(pdf_path: str) -> tuple[date | None, date | None]:
    """Public wrapper to extract the statement coverage period (start_date, end_date)."""
    try:
        start, end = _parse_period_dates(pdf_path)
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _infer_year(prev_dt: date | None, dd: int, mm: int, yyyy: int | None, base_year: int | None) -> date:
    """Build a date using an explicit year if provided, else infer from base_year and rollovers."""
    if yyyy is not None:
        return date(yyyy, mm, dd)

    year = base_year if base_year is not None else (prev_dt.year if prev_dt else date.today().year)

    candidate = date(year, mm, dd)

    # Year rollover handling: if dates jump "backwards" significantly, assume we crossed into next year.
    if prev_dt is not None:
        # If candidate is more than ~180 days before prev_dt, it's likely next year.
        if (candidate - prev_dt).days < -180:
            candidate = date(year + 1, mm, dd)

    return candidate


def _split_type_and_description(raw: str) -> tuple[str, str]:
    """Return (transaction_type, description_without_type_prefix)."""
    s = (raw or "").strip()

    # Returned Direct Debit rule (keep as description prefix, but transaction type becomes Direct Debit)
    if s.lower().startswith("returned direct debit"):
        return "Direct Debit", s  # keep full description

    for prefix in TYPE_PREFIXES:
        if s.startswith(prefix):
            desc = s[len(prefix):].strip()
            ttype = prefix.title()

            # Normalise types to match project rules
            if prefix == "Card Transaction":
                ttype = "Card Payment"
            elif prefix == "Charges":
                ttype = "Bank Charges"

            # Apply your global overrides for Card Payment classification
            dlow = desc.lower()
            if "applepay" in dlow:
                return "Card Payment", desc
            if "clearpay" in dlow:
                return "Card Payment", desc
            # Contactless is not a common RBS prefix, but keep rule parity
            if "contactless" in dlow:
                return "Card Payment", desc
            if desc.endswith("GB"):
                return "Card Payment", desc

            return ttype, desc

    # No known prefix; keep as-is
    desc = s
    ttype = ""  # will be title-cased later if we can infer

    dlow = desc.lower()
    if "applepay" in dlow or "clearpay" in dlow or "contactless" in dlow or desc.endswith("GB"):
        ttype = "Card Payment"

    return ttype, desc


def extract_account_holder_name(pdf_path: str) -> str:
    """Best-effort: extract the account holder name from page 1.

    For this RBS template, `extract_words()` is unreliable (it can split into single
    characters), so we use `extract_text()` and line heuristics.

    Desired output example:
      "MRS K HATTON & MISS O HATTON
DOLLY MIXTURES DAY NURSERY."
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""

    lines = [ln.strip() for ln in (text.splitlines() if text else []) if ln.strip()]

    # Preferred: address block has the personal name line followed by business name.
    person_idx = None
    for i, ln in enumerate(lines):
        if re.match(r"^(MR|MRS|MS|MISS)\b", ln, flags=re.IGNORECASE):
            person_idx = i
            break

    if person_idx is not None:
        person = lines[person_idx].strip()
        biz = ""
        if person_idx + 1 < len(lines):
            biz = lines[person_idx + 1].strip()

        # Clean business line
        if biz.upper().startswith("T/A "):
            biz = biz[4:].strip()
        # Strip trailing statement metadata that sometimes appears on the same line
        biz = re.split(r"\bStatement Date\b", biz, maxsplit=1)[0].strip()
        biz = " ".join(biz.split())
        biz = biz.replace("NU RSERY", "NURSERY")
        if biz and not biz.endswith("."):
            biz += "."

        if person and biz:
            return person + "\n" + biz
        if person:
            return person


    # Fallback: take the Account Name header block
    for i, ln in enumerate(lines):
        if ln.lower().startswith("account name"):
            j = i + 1
            picked = []
            while j < len(lines) and len(picked) < 2:
                if lines[j].strip():
                    picked.append(lines[j].strip())
                j += 1
            if picked:
                return " ".join(picked)

    return ""


def extract_statement_balances(pdf_path: str) -> dict:
    """Return start/end balances from the Summary on page 1.

    Uses:
      Previous Balance £xx
      New Balance £xx

    Returns: {"start_balance": float|None, "end_balance": float|None}
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return {"start_balance": None, "end_balance": None}
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return {"start_balance": None, "end_balance": None}

    start = None
    end = None

    m1 = re.search(r"Previous Balance\s+£\s*(\d{1,3}(?:,\d{3})*\.\d{2})", text)
    if m1:
        try:
            start = _money_to_float(m1.group(1))
        except Exception:
            start = None

    m2 = re.search(r"New Balance\s+£\s*(\d{1,3}(?:,\d{3})*\.\d{2})", text)
    if m2:
        try:
            end = _money_to_float(m2.group(1))
        except Exception:
            end = None

    return {"start_balance": start, "end_balance": end}


def extract_transactions(pdf_path: str) -> list[dict]:
    """Extract transactions from all pages (text parsing).

    Important: On this RBS PDF template, `page.extract_words()` can split text into
    single characters, which breaks coordinate-based parsing. However,
    `page.extract_text()` is clean and contains the full table rows.

    We therefore parse line-by-line, and determine debit/credit signs using the
    *balance delta* method:
      - Each transaction row provides a running Balance.
      - Let prev_balance be the prior row's Balance.
      - Let amount_mag be the row's transaction amount (absolute value).
      - If (balance - prev_balance) ~= +amount_mag => credit
      - If (balance - prev_balance) ~= -amount_mag => debit

    This correctly handles "OnLine Transaction" rows that can be either credits
    (Refund / Salary) or debits.
    """

    base_year = _parse_period_year(pdf_path)

    transactions: list[dict] = []

    current_dt: date | None = None
    block_lines: list[str] = []

    prev_dt: date | None = None
    prev_balance: float | None = None

    def _flush_block() -> None:
        nonlocal current_dt, block_lines, prev_dt, prev_balance

        if current_dt is None or not block_lines:
            block_lines = []
            return

        raw = " ".join(ln.strip() for ln in block_lines if ln.strip())
        raw = " ".join(raw.split()).strip()

        # Ignore brought forward blocks
        if raw.upper().startswith("BROUGHT FORWARD"):
            block_lines = []
            return

        # Collect all monetary values in the block.
        monies = [m.group(0) for m in MONEY_RE.finditer(raw)]
        if len(monies) < 2:
            block_lines = []
            return

        try:
            amt_mag = _money_to_float(monies[-2])
            bal = _money_to_float(monies[-1])
        except Exception:
            block_lines = []
            return

        # Remove trailing monetary columns from the description
        desc_text = raw
        desc_text = re.sub(
            r"\s*" + re.escape(monies[-2]) + r"\s*" + re.escape(monies[-1]) + r"\s*$",
            "",
            desc_text,
        ).strip()

        ttype, desc = _split_type_and_description(desc_text)
        desc = (desc or "").strip()
        desc = re.sub(r"\s+CD\s*\d{4}\b\s*$", "", desc, flags=re.IGNORECASE).strip()

        signed_amt = None
        if prev_balance is not None:
            delta = round(bal - prev_balance, 2)
            if abs(delta - round(amt_mag, 2)) <= 0.02:
                signed_amt = amt_mag
            elif abs(delta + round(amt_mag, 2)) <= 0.02:
                signed_amt = -amt_mag

        # Fallback if we couldn't infer sign (should be rare)
        if signed_amt is None:
            if (ttype or "").strip().lower() in {"automated credit"}:
                signed_amt = amt_mag
            else:
                signed_amt = -amt_mag

        transactions.append(
            {
                "Date": current_dt,
                "Transaction Type": (ttype or "Unknown"),
                "Description": desc,
                "Amount": round(float(signed_amt), 2),
                "Balance": round(float(bal), 2),
            }
        )

        prev_dt = current_dt
        prev_balance = round(float(bal), 2)
        block_lines = []

    def _line_starts_new_tx(line: str) -> tuple[bool, date | None, str]:
        parts = line.split()
        isd, dd, mm, yyyy, consumed = _is_date_row(parts)
        if not isd:
            return False, None, line

        dt = _infer_year(prev_dt, dd, mm, yyyy, base_year)
        remainder = " ".join(parts[consumed:]).strip()
        return True, dt, remainder

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text:
                continue

            lines = [ln.rstrip() for ln in text.splitlines()]

            # Find the table header; if not found, still parse once we see a date row.
            in_table = False

            for ln in lines:
                low = (ln or "").strip().lower()
                if not low:
                    continue

                if ("date" in low and "description" in low and "balance" in low and ("paid" in low or "withdrawn" in low)):
                    in_table = True
                    continue

                if low.startswith("retstmt"):
                    break

                if not in_table:
                    starts, _dt, _rem = _line_starts_new_tx(ln.strip())
                    if starts:
                        in_table = True
                    else:
                        continue

                line = ln.strip()

                # Skip the header row if it repeats
                if ("date" in low and "description" in low):
                    continue

                # Handle BROUGHT FORWARD lines to seed prev_balance
                if line.upper().startswith("BROUGHT FORWARD"):
                    monies = [m.group(0) for m in MONEY_RE.finditer(line)]
                    if monies:
                        try:
                            prev_balance = _money_to_float(monies[-1])
                        except Exception:
                            pass
                    continue

                # New transaction by date line
                starts, dt, rem = _line_starts_new_tx(line)
                if starts:
                    _flush_block()
                    current_dt = dt
                    block_lines = [rem] if rem else []
                    continue

                # New transaction on same date by type prefix
                if current_dt is not None and _is_type_prefix_row(line):
                    _flush_block()
                    block_lines = [line]
                    continue

                # Continuation line
                if current_dt is not None:
                    block_lines.append(line)

            _flush_block()

    return transactions

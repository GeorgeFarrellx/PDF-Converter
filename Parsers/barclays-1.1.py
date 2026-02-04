# Version: barclays-1.1.py
"""Barclays (Business Current Account) PDF parser.

Text-based PDFs (no OCR).

Required output columns (used by Main):
- Date (datetime.date)
- Transaction Type (str)
- Description (str)
- Amount (float)  # credits positive, debits negative
- Balance (float | None)

Also provides:
- extract_statement_balances(pdf_path) -> {start_balance, end_balance}
- extract_account_holder_name(pdf_path) -> str

Notes:
- Handles multi-line descriptions (continuation lines)
- Handles statement year changes using the "At a glance" period
- Ignores non-transaction content (footers, protection text, totals blocks)
- Uses a robust money+balance extraction strategy and trims stray summary text
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

DATE_RE = re.compile(r"^(?P<dd>\d{1,2})\s+(?P<mon>[A-Za-z]{3})\b")

# Money like 1,234.56
MONEY_RE = re.compile(r"(?<!\d)-?\d{1,3}(?:,\d{3})*\.\d{2}\b")

# Statement "At a glance" period line examples:
#   23 Mar - 24 Apr 2024
#   25 Dec 2024 - 24 Jan 2025
PERIOD_SAME_END_YEAR_RE = re.compile(
    r"(?P<sd>\d{1,2})\s+(?P<sm>[A-Za-z]{3})\s*-\s*(?P<ed>\d{1,2})\s+(?P<em>[A-Za-z]{3})\s+(?P<ey>\d{4})"
)
PERIOD_BOTH_YEARS_RE = re.compile(
    r"(?P<sd>\d{1,2})\s+(?P<sm>[A-Za-z]{3})\s+(?P<sy>\d{4})\s*-\s*(?P<ed>\d{1,2})\s+(?P<em>[A-Za-z]{3})\s+(?P<ey>\d{4})"
)

# Some Barclays statements split the end-year onto the next header line, e.g.
#   "20 Dec 2024 - 17 Jan" then ".... 2025"
PERIOD_STARTYEAR_ENDNOYEAR_RE = re.compile(
    r"(?P<sd>\d{1,2})\s+(?P<sm>[A-Za-z]{3})\s+(?P<sy>\d{4})\s*-\s*(?P<ed>\d{1,2})\s+(?P<em>[A-Za-z]{3})\b"
)

START_BAL_RE = re.compile(r"Start\s+balance\s+£\s*(?P<amt>[\d,]+\.\d{2})", re.IGNORECASE)
END_BAL_RE = re.compile(r"End\s+balance\s+£\s*(?P<amt>[\d,]+\.\d{2})", re.IGNORECASE)

# Phrases that often appear after the running balance in extracted text
TRUNCATE_AFTER_PHRASES = [
    "u commission charges",
    "commission charges",
    "u interest paid",
    "interest paid",
    "end balance",
]

# Boilerplate / footer lines to skip as transaction continuation
SKIP_LINE_PREFIXES = [
    "barclays bank",
    "registered in",
    "authorised by",
    "page",
    "continued",
    "your deposit is eligible",
    "by the financial services",
    "compensation scheme",
    "on ",  # often "On 25 Mar" protection message
]


def _money_to_float(s: str) -> float:
    return float(s.replace(",", ""))


def _month_num(mon: str) -> int | None:
    if not mon:
        return None
    return MONTHS.get(mon.strip().upper())


def _parse_period_from_page1(pdf_path: str) -> tuple[date | None, date | None]:
    """Extract statement start/end date from the PDF.

    Barclays PDFs sometimes start with an account summary cover page. The actual
    account statement (with the "At a glance" period) often begins on page 2.

    We therefore scan the first few pages for the period line.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None, None

            max_pages = min(4, len(pdf.pages))
            for p in range(max_pages):
                text = pdf.pages[p].extract_text() or ""
                if not text:
                    continue

                lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

                for idx, ln in enumerate(lines):
                    m = PERIOD_BOTH_YEARS_RE.search(ln)
                    if m:
                        sd = int(m.group("sd"))
                        sm = _month_num(m.group("sm"))
                        sy = int(m.group("sy"))
                        ed = int(m.group("ed"))
                        em = _month_num(m.group("em"))
                        ey = int(m.group("ey"))
                        if sm and em:
                            try:
                                return date(sy, sm, sd), date(ey, em, ed)
                            except Exception:
                                pass

                    m = PERIOD_SAME_END_YEAR_RE.search(ln)
                    if m:
                        sd = int(m.group("sd"))
                        sm = _month_num(m.group("sm"))
                        ed = int(m.group("ed"))
                        em = _month_num(m.group("em"))
                        ey = int(m.group("ey"))
                        if sm and em:
                            sy = ey - 1 if sm > em else ey
                            try:
                                return date(sy, sm, sd), date(ey, em, ed)
                            except Exception:
                                pass

                    # Split-year format:
                    #   "20 Dec 2024 - 17 Jan" then next line contains "2025"
                    m = PERIOD_STARTYEAR_ENDNOYEAR_RE.search(ln)
                    if m:
                        sd = int(m.group("sd"))
                        sm = _month_num(m.group("sm"))
                        sy = int(m.group("sy"))
                        ed = int(m.group("ed"))
                        em = _month_num(m.group("em"))
                        if sm and em:
                            tail = " ".join(lines[idx : idx + 3])
                            years = [int(mo.group(0)) for mo in re.finditer(r"\b(19|20)\d{2}\b", tail)]
                            if years:
                                # Prefer a year that is not the start-year if present; otherwise use the last one.
                                ey = next((y for y in years if y != sy), years[-1])
                                try:
                                    return date(sy, sm, sd), date(ey, em, ed)
                                except Exception:
                                    pass

                # Fallback: search a joined blob (handles mid-line wraps).
                blob = " ".join(text.split())

                m = PERIOD_BOTH_YEARS_RE.search(blob)
                if m:
                    sd = int(m.group("sd"))
                    sm = _month_num(m.group("sm"))
                    sy = int(m.group("sy"))
                    ed = int(m.group("ed"))
                    em = _month_num(m.group("em"))
                    ey = int(m.group("ey"))
                    if sm and em:
                        try:
                            return date(sy, sm, sd), date(ey, em, ed)
                        except Exception:
                            pass

                m = PERIOD_SAME_END_YEAR_RE.search(blob)
                if m:
                    sd = int(m.group("sd"))
                    sm = _month_num(m.group("sm"))
                    ed = int(m.group("ed"))
                    em = _month_num(m.group("em"))
                    ey = int(m.group("ey"))
                    if sm and em:
                        sy = ey - 1 if sm > em else ey
                        try:
                            return date(sy, sm, sd), date(ey, em, ed)
                        except Exception:
                            pass

                m = PERIOD_STARTYEAR_ENDNOYEAR_RE.search(blob)
                if m:
                    sd = int(m.group("sd"))
                    sm = _month_num(m.group("sm"))
                    sy = int(m.group("sy"))
                    ed = int(m.group("ed"))
                    em = _month_num(m.group("em"))
                    if sm and em:
                        years = [int(mo.group(0)) for mo in re.finditer(r"\b(19|20)\d{2}\b", blob[m.end() : m.end() + 120])]
                        if years:
                            ey = next((y for y in years if y != sy), years[-1])
                            try:
                                return date(sy, sm, sd), date(ey, em, ed)
                            except Exception:
                                pass

    except Exception:
        return None, None

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
        start, end = _parse_period_from_page1(pdf_path)
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _infer_year(dd: int, mm: int, period_start: date | None, period_end: date | None) -> int:
    """Infer year for a dd/Mon transaction using statement period."""
    if period_start and period_end:
        if period_start.year == period_end.year:
            return period_end.year
        # spans year-end
        return period_start.year if mm >= period_start.month else period_end.year

    # Best-effort fallback: use current year
    return date.today().year


def _looks_like_header(line: str) -> bool:
    low = (line or "").strip().lower()
    if not low:
        return False
    if low.startswith("date ") and "description" in low and "balance" in low:
        return True
    if low.startswith("at a glance"):
        return True
    return False


def _looks_like_total_or_summary(line: str) -> bool:
    low = (line or "").strip().lower()
    if not low:
        return False

    if low.startswith("money out") or low.startswith("money in"):
        return True
    if low.startswith("start balance") or low.startswith("end balance"):
        return True
    if low.startswith("issued on"):
        return True

    # Table footer lines at the end of each statement
    if low.startswith("balance brought forward"):
        return True
    if low.startswith("total payments/receipts") or low.startswith("total payments"):
        return True

    return False


def _should_skip_continuation(line: str) -> bool:
    low = (line or "").strip().lower()
    if not low:
        return True

    # Skip very common footer/boilerplate
    for p in SKIP_LINE_PREFIXES:
        if low.startswith(p):
            # Keep "on <date>" lines out of descriptions
            return True

    # Lines that are just page numbers
    if low.isdigit():
        return True

    return False


def _truncate_after_summary_phrases(s: str) -> str:
    """Cut text at the first occurrence of known trailing summary phrases."""
    low = s.lower()
    cut = None
    for ph in TRUNCATE_AFTER_PHRASES:
        idx = low.find(ph)
        if idx != -1:
            cut = idx if cut is None else min(cut, idx)
    if cut is not None:
        return s[:cut].strip()
    return s.strip()


def _split_type_and_description(raw: str) -> tuple[str, str, bool]:
    """Return (transaction_type, description, is_credit).

    Applies project rules:
    - Returned Direct Debit => type Direct Debit; keep full description
    - ApplePay/Clearpay/Contactless/endswith GB => Card Payment
    - Charges / Commission charges => Bank Charges
    - Otherwise keep Barclays wording (Title Case) and remove type prefix from description
    """
    s = (raw or "").strip()
    if not s:
        return "Unknown", "", False

    low = s.lower()

    if low.startswith("returned direct debit"):
        return "Direct Debit", s, True

    if low.startswith("commission charges") or low == "charges" or low.startswith("charges "):
        return "Bank Charges", s, False

    # Common Barclays prefixes
    prefixes = [
        "Direct Debit",
        "Standing Order",
        "Card Payment",
        "Direct Credit",
        "Bank Transfer",
        "Transfer",
        "Cash Withdrawal",
        "Cash Deposit",
        "Bill Payment",
    ]

    txn_type = ""
    desc = s
    is_credit = False

    for p in prefixes:
        if s.lower().startswith(p.lower()):
            txn_type = p
            rest = s[len(p) :].strip()

            # Barclays often has "to" / "from"
            if rest.lower().startswith("to "):
                desc = rest[3:].strip()
                is_credit = False
            elif rest.lower().startswith("from "):
                desc = rest[5:].strip()
                is_credit = True
            else:
                desc = rest.strip() if rest else s

            break

    if not txn_type:
        # Fallback: try split on "to" / "from" (case-insensitive & whitespace-safe)
        m = re.search(r"\bto\b", low)
        if m:
            txn_type = s[: m.start()].strip()
            desc = s[m.end() :].strip()
            is_credit = False
        else:
            m = re.search(r"\bfrom\b", low)
            if m:
                txn_type = s[: m.start()].strip()
                desc = s[m.end() :].strip()
                is_credit = True
            else:
                txn_type = "Transaction"
                desc = s
                is_credit = False

    # Global card-payment overrides
    dlow = desc.lower()
    if "applepay" in dlow or "clearpay" in dlow or "contactless" in dlow or desc.endswith("GB"):
        txn_type = "Card Payment"
        is_credit = False

    # Title case for type (preserve simple words)
    txn_type = " ".join(w[:1].upper() + w[1:].lower() for w in txn_type.split())

    return txn_type, desc, is_credit


def _parse_amount_and_balance(block: str) -> tuple[float | None, float | None, str]:
    """Return (amount, balance, cleaned_description_source).

    Strategy:
    - Trim known trailing summary fragments (commission charges, end balance, interest paid)
    - Find all money tokens in the cleaned block
    - Use last token as running balance, second-last as transaction amount magnitude
    - Remove those two tokens from the text (even if they are not at the end) to produce
      a clean description source.
    """
    cleaned = _truncate_after_summary_phrases(block)

    monies = [m.group(0) for m in MONEY_RE.finditer(cleaned)]
    if len(monies) < 2:
        return None, None, cleaned

    amt_token = monies[-2]
    bal_token = monies[-1]

    try:
        amt = _money_to_float(amt_token)
        bal = _money_to_float(bal_token)
    except Exception:
        return None, None, cleaned

    def _remove_last_occurrence(text: str, token: str) -> str:
        idx = text.rfind(token)
        if idx == -1:
            return text
        return (text[:idx] + text[idx + len(token) :]).strip()

    cleaned_desc = _remove_last_occurrence(cleaned, bal_token)
    cleaned_desc = _remove_last_occurrence(cleaned_desc, amt_token)
    cleaned_desc = " ".join(cleaned_desc.split()).strip()

    return float(amt), float(bal), cleaned_desc


# ----------------------------
# Required API for Main.py
# ----------------------------


def extract_account_holder_name(pdf_path: str) -> str:
    """Best-effort: extract the account holder name.

    Barclays business statements can begin with an account summary cover page.
    The addressee line is often "THE DIRECTOR" / "THE DIRECTORS" followed by the
    company name (e.g. "LOW STEPPA LTD").

    We prefer the company name and avoid generic headings and addressee titles.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""
            # Page 1 is sometimes a cover/summary; the company name is typically
            # present on page 1 and/or at the top of page 2.
            max_pages = min(2, len(pdf.pages))
            texts = [(pdf.pages[i].extract_text() or "") for i in range(max_pages)]
    except Exception:
        return ""

    lines: list[str] = []
    for t in texts:
        lines.extend([ln.strip() for ln in (t.splitlines() if t else []) if ln.strip()])

    # Prefer explicit company-style names
    ORG_HINTS = (" LTD", " LIMITED", " LLP", " PLC", " LIMITED.", " LTD.")

    def _is_good_name(cand: str) -> bool:
        if not cand:
            return False
        up = cand.strip().upper()
        if up in {"THE DIRECTOR", "THE DIRECTORS"}:
            return False
        if any(ch.isdigit() for ch in cand):
            return False
        if len(cand) < 3 or len(cand) > 80:
            return False
        return True

    # If we see THE DIRECTOR(S), the next meaningful line is usually the company name.
    for i, ln in enumerate(lines):
        if ln.strip().upper() in {"THE DIRECTOR", "THE DIRECTORS"}:
            for j in range(i + 1, min(i + 8, len(lines))):
                cand = lines[j].strip()
                if not _is_good_name(cand):
                    continue
                # Strong preference for org-like suffixes
                cup = cand.upper()
                if any(h in cup for h in ORG_HINTS):
                    return cand
                # Otherwise return the first plausible line
                return cand

    # Otherwise: pick the first all-caps line that doesn't look like a generic heading.
    skip = {
        "YOUR BUSINESS CURRENT ACCOUNT",
        "YOUR BUSINESS ACCOUNTS – AT A GLANCE",
        "YOUR BUSINESS ACCOUNTS - AT A GLANCE",
        "THE DIRECTOR",
        "THE DIRECTORS",
        "AT A GLANCE",
        "DATE DESCRIPTION MONEY OUT £ MONEY IN £ BALANCE £",
        "DATE DESCRIPTION MONEY OUT £ MONEY IN £ BALANCE",
        "DATE DESCRIPTION MONEY OUT MONEY IN BALANCE",
    }

    # First pass: all-caps with obvious org hint
    for ln in lines:
        cand = ln.strip()
        if not cand:
            continue
        up = cand.upper()
        if up in skip:
            continue
        if not _is_good_name(cand):
            continue
        if cand.upper() == cand and any(h in up for h in ORG_HINTS):
            return cand

    # Second pass: first all-caps plausible line
    for ln in lines:
        cand = ln.strip()
        if not cand:
            continue
        up = cand.upper()
        if up in skip:
            continue
        if not _is_good_name(cand):
            continue
        if cand.upper() == cand:
            return cand

    # Last resort: first non-empty non-numeric line
    for ln in lines:
        if ln and not any(ch.isdigit() for ch in ln):
            if ln.strip().upper() not in {"THE DIRECTOR", "THE DIRECTORS"}:
                return ln.strip()

    return ""



def extract_statement_balances(pdf_path: str) -> dict:
    """Return start/end balances from the statement.

    Barclays Business statements typically show:
      Start balance £xxx.xx
      End balance £xxx.xx

    We search across all pages to be safe.
    """
    start = None
    end = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if not text:
                    continue

                blob = " ".join(text.split())

                if start is None:
                    m = START_BAL_RE.search(blob)
                    if m:
                        try:
                            start = _money_to_float(m.group("amt"))
                        except Exception:
                            start = None

                m2 = END_BAL_RE.search(blob)
                if m2:
                    try:
                        end = _money_to_float(m2.group("amt"))
                    except Exception:
                        end = end

    except Exception:
        return {"start_balance": None, "end_balance": None}

    return {"start_balance": start, "end_balance": end}


def extract_transactions(pdf_path: str) -> list[dict]:
    """Extract transactions from all pages."""

    period_start, period_end = _parse_period_from_page1(pdf_path)

    # Use the statement start balance as the "previous balance" so we can compute
    # each transaction amount from balance deltas (more robust than relying on the
    # extracted money-out/money-in column order).
    bal_info = extract_statement_balances(pdf_path)
    prev_balance: float | None = bal_info.get("start_balance")

    txns: list[dict] = []

    current_dt: date | None = None
    block_lines: list[str] = []

    # Barclays often omits the date on subsequent rows for the same day.
    # Detect the start of a new transaction row based on common row prefixes.
    ROW_START_PREFIXES = (
        "direct debit",
        "standing order",
        "card payment",
        "card purchase",
        "card refund",
        "direct credit",
        "on-line banking",
        "on line banking",
        "online banking",
        "internet banking",
        "commission charges",
        "charges",
        "returned direct debit",
        "cash withdrawal",
        "cash machine withdrawal",
        "cash machine deposit",
        "cash deposit",
        "bank transfer",
        "transfer",
        "bill payment",
        "deposit at barclays",
    )

    def _is_new_row_start(s: str) -> bool:
        low = (s or "").strip().lower()
        return any(low.startswith(p) for p in ROW_START_PREFIXES)

    def _block_has_amount_and_balance(lines: list[str]) -> bool:
        blob = " ".join(lines)
        return len(MONEY_RE.findall(blob)) >= 2

    def flush_block() -> None:
        nonlocal current_dt, block_lines, prev_balance

        if current_dt is None or not block_lines:
            block_lines = []
            return

        raw = " ".join(ln.strip() for ln in block_lines if ln.strip())
        raw = " ".join(raw.split()).strip()

        block_lines = []

        if not raw:
            return

        low = raw.lower()

        # Ignore obvious non-transaction blocks
        if low.startswith("start balance"):
            return
        if low.startswith("balance") and "start" in low:
            return
        if low.startswith("money out") or low.startswith("money in"):
            return
        if "date description" in low and "balance" in low:
            return

        amt_mag, bal, desc_source = _parse_amount_and_balance(raw)
        if bal is None:
            return

        # Determine transaction type/description
        txn_type, desc, is_credit = _split_type_and_description(desc_source)

        # Prefer balance-delta amounts when possible
        if prev_balance is not None:
            signed_amt = round(float(bal) - float(prev_balance), 2)
            prev_balance = float(bal)
        else:
            if amt_mag is None:
                return
            signed_amt = round(float(amt_mag) if is_credit else -float(amt_mag), 2)

        txns.append(
            {
                "Date": current_dt,
                "Transaction Type": txn_type,
                "Description": desc.strip(),
                "Amount": signed_amt,
                "Balance": round(float(bal), 2),
            }
        )

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text:
                continue

            lines = [ln.rstrip() for ln in text.splitlines()]

            in_table = False

            for ln in lines:
                line = (ln or "").strip()
                if not line:
                    continue

                if _looks_like_header(line):
                    in_table = True
                    continue

                if _looks_like_total_or_summary(line):
                    # Skip summary lines; don't treat as continuation
                    continue

                # If not yet in a table, wait until we see a date line
                m = DATE_RE.match(line)
                if not in_table:
                    if m:
                        in_table = True
                    else:
                        continue

                # Footers/boilerplate
                low = line.lower()
                if low.startswith("barclays bank"):
                    continue

                # New transaction starts with a date
                m = DATE_RE.match(line)
                if m:
                    # flush previous
                    flush_block()

                    dd = int(m.group("dd"))
                    mon = m.group("mon")
                    mm = _month_num(mon)
                    if not mm:
                        current_dt = None
                        continue

                    yy = _infer_year(dd, mm, period_start, period_end)
                    try:
                        current_dt = date(yy, mm, dd)
                    except Exception:
                        current_dt = None
                        continue

                    remainder = line[m.end() :].strip()
                    # Ignore non-transaction balance rows (statement footer)
                    if remainder.lower().startswith("start balance"):
                        current_dt = None
                        block_lines = []
                        continue
                    if remainder.lower().startswith("balance carried forward"):
                        current_dt = None
                        block_lines = []
                        continue

                    block_lines = [remainder] if remainder else []
                    continue

                # Continuation / same-date rows (date omitted)
                if current_dt is not None:
                    # If we already have a complete row and the next line looks like the start
                    # of a new transaction row, flush and start a new block on the same date.
                    if _is_new_row_start(line) and block_lines and _block_has_amount_and_balance(block_lines):
                        flush_block()
                        block_lines = [line]
                        continue

                    if _should_skip_continuation(line):
                        continue

                    # Avoid numeric continuation lines (exchange rates / non-sterling fees)
                    # once we already have amount+balance captured for this row.
                    if block_lines and _block_has_amount_and_balance(block_lines) and MONEY_RE.search(line):
                        continue

                    block_lines.append(line)

            flush_block()

            # Prevent "Balance brought forward" at top of next page being treated as a continuation
            current_dt = None
            block_lines = []

    return txns

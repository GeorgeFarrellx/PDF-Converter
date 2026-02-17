"""Santander PDF statement parser (text-based, no OCR)

File: santander-1.8.py
Version: 1.8

Notes:
- Supports four Santander layouts: Business Banking statements, Personal current account statements,
  Online Banking current account exports, and Online Banking credit card exports.
- Applies global transaction type rules as specified in the main project instructions.
"""

__version__ = "1.8"

import os
import re
import datetime as _dt
from typing import List, Dict, Optional, Tuple

import pdfplumber


# -----------------------------
# Helpers: text + money parsing
# -----------------------------

_MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

_MONEY_RE = re.compile(r'(?<!\w)(\(?-?\s*£?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?\)?)(?!\w)')

# Online format date: 31/01/2025
_DATE_DMY_SLASH_FULL_RE = re.compile(r'^\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s*(.*)$')

# Business format date: 3rd Dec / 1st Jan etc
_DATE_ORD_MON_RE = re.compile(r'^\s*(\d{1,2})(st|nd|rd|th)\s*([A-Za-z]{3,9})\s*(.*)$', re.IGNORECASE)


def _clean_text(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def _parse_money(token: str) -> Optional[float]:
    if token is None:
        return None
    t = token.strip()
    if not t:
        return None

    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1].strip()

    t = t.replace("£", "").replace(" ", "")
    if t.startswith("-"):
        neg = True
        t = t[1:]

    t = t.replace(",", "")

    try:
        val = float(t)
    except Exception:
        return None

    return -val if neg else val


def _extract_money_tokens(line: str) -> List[str]:
    return _MONEY_RE.findall(line or "")


def _extract_money_values(line: str) -> List[float]:
    """Return parsed money values from a line.

    Santander Online Banking PDFs can include header/footer text like "Page 1 of 37".
    Those integers must NOT be treated as money. We therefore only accept tokens that
    look like real money with exactly 2 decimal places.
    """
    vals: List[float] = []
    for tok in _extract_money_tokens(line):
        t = (tok or "").strip().replace(" ", "")
        # strip surrounding parentheses for the decimal check
        if t.startswith("(") and t.endswith(")"):
            t_chk = t[1:-1]
        else:
            t_chk = t
        t_chk = t_chk.replace("£", "")
        # Must end with .dd
        if len(t_chk) < 3 or t_chk[-3] != "." or (not t_chk[-2:].isdigit()):
            continue
        v = _parse_money(tok)
        if v is not None:
            vals.append(v)
    return vals


def _remove_money_from_text(s: str) -> str:
    return _clean_text(_MONEY_RE.sub(" ", s or ""))


def _title_case_keep_acronyms(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s
    words = s.split()
    out = []
    for w in words:
        if w.isupper() and len(w) <= 4:
            out.append(w)
        else:
            out.append(w[:1].upper() + w[1:].lower() if w else w)
    return " ".join(out)


def _is_online_noise_line(line: str) -> bool:
    """Filter out Online Banking header/footer lines that can appear mid-table."""
    l = _clean_text(line)
    if not l:
        return True
    low = l.lower()
    if "santander online banking" in low:
        return True
    if "transactions" in low and "santander" in low:
        return True
    if "transaction date:" in low:
        return True
    if "account number:" in low:
        return True
    if "card number:" in low:
        return True
    if low.startswith("page ") and " of " in low:
        return True
    return False


def _strip_online_junk(desc: str) -> str:
    """Remove embedded Online Banking header/footer fragments appended to transaction lines."""
    s = _clean_text(desc)
    if not s:
        return ""

    low = s.lower()
    markers = [
        "santander online banking",
        "transaction date:",
        "account number:",
        "date description money in money out balance",
        "page ",
    ]

    cut = len(s)
    for m in markers:
        i = low.find(m)
        if i != -1 and i < cut:
            cut = i

    s = _clean_text(s[:cut])

    # Trim common trailing punctuation/spaces
    while s and s[-1] in {",", "-"}:
        s = s[:-1].rstrip()

    return s


def _extract_type_prefix_strict(desc: str) -> str:
    """Strict type extraction for ONLINE exports.

    Online rows sometimes start with a payee (e.g. "ELIOR UK PLC ...") rather than a real bank type.
    In those cases we must return "" (and let global rules classify, e.g. Google Pay -> Card Payment).
    """
    s = _clean_text(desc)
    if not s:
        return ""

    known = [
        "BILL PAYMENT VIA FASTER PAYMENT",
        "THIRD PARTY PAYMENT MADE VIA FASTER PAYMENT",
        "FASTER PAYMENTS RECEIPT",
        "DIRECT DEBIT PAYMENT",
        "CARD PAYMENT",
        "BANK GIRO CREDIT",
        "CREDIT FROM",
        "TRANSFER TO",
        "TRANSFER",
        "CHARGES",
        "CASH WITHDRAWAL",
        "FOREIGN CURRENCY CONVERSION FEE",
    ]

    up = s.upper()
    for k in known:
        if up.startswith(k):
            return k

    return ""


# -----------------------------------
# Global transaction type normaliser
# -----------------------------------

def _apply_global_type_rules(tx_type: str, desc: str) -> Tuple[str, str]:
    raw_type = (tx_type or "").strip()
    raw_desc = (desc or "").strip()

    # Returned Direct Debit
    if re.search(r"\breturned\s+direct\s+debit\b", (raw_type + " " + raw_desc).lower()):
        ttype = "Direct Debit"
        if not raw_desc.lower().startswith("returned direct debit"):
            raw_desc = ("Returned Direct Debit " + raw_desc).strip()
        else:
            # normalize prefix casing
            raw_desc = "Returned Direct Debit" + raw_desc[len("returned direct debit"):]
            raw_desc = raw_desc.strip()
        return ttype, raw_desc

    dlow = raw_desc.lower()
    if (
        "applepay" in dlow
        or "clearpay" in dlow
        or "contactless" in dlow
        or "google pay" in dlow
        or raw_desc.rstrip().endswith("GB")
    ):
        return "Card Payment", raw_desc

    # Otherwise keep wording (Title Case)
    ttype = _title_case_keep_acronyms(raw_type) if raw_type else ""

    # Remove type prefix from description (except returned direct debit already handled)
    if raw_type:
        prefix_re = re.compile(rf'^\s*{re.escape(raw_type)}\s*[-:]\s*', re.IGNORECASE)
        if prefix_re.match(raw_desc):
            raw_desc = prefix_re.sub("", raw_desc).strip()
        else:
            if raw_desc.lower().startswith(raw_type.lower() + " "):
                raw_desc = raw_desc[len(raw_type):].strip()

    return ttype, raw_desc


# -----------------------------
# Detection: Online vs Business
# -----------------------------

def _detect_statement_kind(first_page_text: str) -> str:
    t = (first_page_text or "").lower()
    if "santander online banking" in t and "money in" in t and "money out" in t and "balance" in t:
        return "ONLINE"
    if "card number:" in t and "date card number description" in t and "money in" in t and "money out" in t:
        return "ONLINE_CREDITCARD"
    if (
        "current account" in t
        and "balance brought forward" in t
        and "your balance at close of business" in t
        and "santander business banking" not in t
    ):
        return "PERSONAL"
    if "santander business banking" in t or ("credits" in t and "debits" in t and "balance" in t and "statement number" in t):
        return "BUSINESS"
    # fallback heuristics
    if "transaction date:" in t and "page" in t and "date description money in money out balance" in t:
        return "ONLINE"
    return "BUSINESS"


# -----------------------------
# Period inference (both types)
# -----------------------------

def _parse_full_date_any(s: str) -> Optional[_dt.date]:
    s = _clean_text(s)
    s = re.sub(r'(?i)(\d{1,2})(st|nd|rd|th)(?=\s*[A-Za-z]{3,9}\b)', r'\1', s)
    m = re.match(r'^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$', s)
    if m:
        d = int(m.group(1))
        mo = int(m.group(2))
        y = int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return _dt.date(y, mo, d)
        except Exception:
            return None

    m = re.match(r'^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$', s)
    if m:
        d = int(m.group(1))
        monname = m.group(2).lower()
        mo = _MONTHS.get(monname[:3])
        y = int(m.group(3))
        if mo:
            try:
                return _dt.date(y, mo, d)
            except Exception:
                return None
    return None


def _infer_period_online(blob: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
    # "Transaction date: 01/05/2024 to 31/01/2025"
    m = re.search(
        r"transaction date:\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s+to\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        blob, re.IGNORECASE
    )
    if m:
        d1 = _parse_full_date_any(m.group(1))
        d2 = _parse_full_date_any(m.group(2))
        return d1, d2
    return None, None


def _infer_period_business(blob: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
    # "Your account summary for 3 December 2024 to 2 January 2025"
    m = re.search(
        r"your account summary for.*?(\d{1,2}(?:st|nd|rd|th)?\s*[A-Za-z]{3,9}\s+\d{4})\s+to\s+(\d{1,2}(?:st|nd|rd|th)?\s*[A-Za-z]{3,9}\s+\d{4})",
        blob, re.IGNORECASE | re.DOTALL
    )
    if m:
        d1 = _parse_full_date_any(_clean_text(m.group(1)))
        d2 = _parse_full_date_any(_clean_text(m.group(2)))
        return d1, d2
    return None, None


def _parse_period_from_filename(pdf_path: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
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
        return _dt.date(y1, m1, d1), _dt.date(y2, m2, d2)
    except Exception:
        return None, None


def extract_statement_period(pdf_path: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
    """Public wrapper to extract the statement coverage period (start_date, end_date)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        start, end = _infer_period_business(full_text)
        if start or end:
            return start, end
        start, end = _infer_period_online(full_text)
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _choose_year_for_business_tx(day: int, month: int, period_start: Optional[_dt.date], period_end: Optional[_dt.date],
                                 prev_date: Optional[_dt.date]) -> int:
    # If period exists, it may cross Dec->Jan
    if period_start and period_end:
        if period_start.year == period_end.year:
            return period_start.year
        # cross-year: months >= start.month -> start.year else end.year
        if month >= period_start.month:
            return period_start.year
        return period_end.year

    # fallback rollover based on previous
    if prev_date:
        if month < prev_date.month:
            return prev_date.year + 1
        return prev_date.year

    return _dt.date.today().year


# -----------------------------
# Account holder name extraction
# -----------------------------

def extract_account_holder_name(pdf_path) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""

    # IMPORTANT:
    # Santander "Online Banking -> Transactions" exports do NOT contain the account holder name.
    # Returning a fabricated identifier (e.g., last digits) causes incorrect client headers/filenames.
    # So for ONLINE statements we return blank and let the user-supplied client name drive output.
    kind = _detect_statement_kind(text)
    if kind in {"ONLINE", "ONLINE_CREDITCARD"}:
        return ""

    blob = _clean_text(text)

    # Business/Personal: "Account name: ERFT LIMITED" or "Account name ERFT LIMITED"
    m = re.search(r"\baccount name\b\s*:?\s*([A-Z0-9&'.,\-]{2,}(?:\s+[A-Z0-9&'.,\-]{2,}){0,20})", blob, re.IGNORECASE)
    if m:
        cand = _clean_text(m.group(1))
        parts = cand.split()
        while parts and len(parts[-1]) == 1:
            parts.pop()
        cand = " ".join(parts)
        if cand and "santander" not in cand.lower() and "statement" not in cand.lower():
            return cand

    # Fallback (Business): pick the first strong-looking addressee line near the top (often ALL CAPS)
    # e.g. "MR EMERSON RANDELL"
    lines = [ln.strip() for ln in (text.splitlines() if text else []) if ln and ln.strip()]
    postcode_re = r"\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b"
    for idx, ln in enumerate(lines[:25]):
        l = _clean_text(ln)
        low = l.lower()
        if any(k in low for k in ["santander", "business", "banking", "operations", "statement", "account number", "sort code", "page "]):
            continue
        if re.search(postcode_re, l, re.IGNORECASE):
            continue
        if " " not in l and re.search(postcode_re, " ".join(lines[idx:idx + 4]), re.IGNORECASE):
            if not any(marker in l for marker in ["LTD", "LIMITED", "LLP", "PLC", "&"]):
                continue
        if re.match(r"^[A-Z][A-Z '\-]{6,60}$", l) and sum(c.isalpha() for c in l) >= 6:
            if " " not in l and not any(marker in l for marker in ["LTD", "LIMITED", "LLP", "PLC", "&"]):
                continue
            return l

    return ""


# -----------------------------
# Statement balances extraction
# -----------------------------

def extract_statement_balances(pdf_path) -> Dict[str, Optional[float]]:
    start_balance = None
    end_balance = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text_chunks = []
            for page in pdf.pages[:2]:
                text_chunks.append(page.extract_text() or "")
            first_blob = _clean_text("\n".join(text_chunks))
            kind = _detect_statement_kind(pdf.pages[0].extract_text() or "")
    except Exception:
        return {"start_balance": None, "end_balance": None}

    if kind in {"BUSINESS", "PERSONAL"}:
        m = re.search(
            r"balance brought forward.*?£\s*([\d,]+\.\d{2})",
            first_blob, re.IGNORECASE
        )
        if m:
            start_balance = _parse_money(m.group(1))

        m = re.search(
            r"your balance at close of business.*?£\s*([\d,]+\.\d{2})",
            first_blob, re.IGNORECASE
        )
        if m:
            end_balance = _parse_money(m.group(1))

        if end_balance is None:
            m = re.search(r"balance\s*carried\s*forward(?:\s*to\s*next\s*statement)?[:\s]*£?\s*([\d,]+\.\d{2})", first_blob, re.IGNORECASE)
            if m:
                end_balance = _parse_money(m.group(1))

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages[:4]:
                    t = _clean_text(page.extract_text() or "")
                    if start_balance is None:
                        mm = re.search(r"previous statement balance\s*([\d,]+\.\d{2})", t, re.IGNORECASE)
                        if mm:
                            start_balance = _parse_money(mm.group(1))
                    if end_balance is None:
                        mm = re.search(r"current statement balance\s*([\d,]+\.\d{2})", t, re.IGNORECASE)
                        if mm:
                            end_balance = _parse_money(mm.group(1))
                    if end_balance is None:
                        mm = re.search(r"balance\s*carried\s*forward(?:\s*to\s*next\s*statement)?[:\s]*£?\s*([\d,]+\.\d{2})", t, re.IGNORECASE)
                        if mm:
                            end_balance = _parse_money(mm.group(1))
        except Exception:
            pass

        return {"start_balance": start_balance, "end_balance": end_balance}

    txs = extract_transactions(pdf_path)

    if kind == "ONLINE_CREDITCARD":
        init_m = re.search(r"initial\s+balance\s*£?\s*([\d,]+\.\d{2})", first_blob, re.IGNORECASE)
        if init_m:
            init_val = _parse_money(init_m.group(1))
            if init_val is not None:
                start_balance = -abs(float(init_val))

        if txs and txs[-1].get("Balance") is not None:
            try:
                end_balance = float(txs[-1]["Balance"])
            except Exception:
                end_balance = None

        if start_balance is not None and end_balance is None:
            total = 0.0
            for t in txs:
                try:
                    total += float(t.get("Amount", 0.0))
                except Exception:
                    pass
            end_balance = round(float(start_balance) + total, 2)

        return {"start_balance": start_balance, "end_balance": end_balance}

    first_tx = None
    last_tx = None
    for t in txs:
        if t.get("Balance") is not None and isinstance(t.get("Date"), _dt.date):
            first_tx = t
            break
    for t in reversed(txs):
        if t.get("Balance") is not None and isinstance(t.get("Date"), _dt.date):
            last_tx = t
            break

    if first_tx is not None:
        try:
            start_balance = round(float(first_tx["Balance"]) - float(first_tx["Amount"]), 2)
        except Exception:
            start_balance = float(first_tx["Balance"])

    if last_tx is not None:
        try:
            end_balance = float(last_tx["Balance"])
        except Exception:
            end_balance = None

    return {"start_balance": start_balance, "end_balance": end_balance}


# -----------------------------
# Transactions: BUSINESS format
# -----------------------------

_CREDIT_HINTS = [
    "receipt", "credit", "bank giro credit", "giro credit", "paid in", "from ",
]
_DEBIT_HINTS = [
    "card payment", "direct debit", "debit", "bill payment", "faster payment to", "payment to",
    "transfer to", "cash withdrawal", "charges", "fee", "unpaid", "paid item",
]


def _infer_sign_from_description(desc: str) -> Optional[int]:
    d = (desc or "").lower()

    for k in _CREDIT_HINTS:
        if k in d:
            return +1

    for k in _DEBIT_HINTS:
        if k in d:
            return -1

    return None


def _extract_type_prefix(desc: str) -> str:
    s = _clean_text(desc)
    if not s:
        return ""

    known = [
        "BILL PAYMENT VIA FASTER PAYMENT",
        "THIRD PARTY PAYMENT MADE VIA FASTER PAYMENT",
        "FASTER PAYMENTS RECEIPT",
        "DIRECT DEBIT PAYMENT",
        "CARD PAYMENT",
        "BANK GIRO CREDIT",
        "CREDIT FROM",
        "TRANSFER TO",
        "TRANSFER",
        "CHARGES",
        "CASH WITHDRAWAL",
        "FOREIGN CURRENCY CONVERSION FEE",
    ]
    up = s.upper()
    for k in known:
        if up.startswith(k):
            return k

    words = s.split()
    prefix_words = []
    for w in words[:4]:
        if re.match(r"^[A-Za-z][A-Za-z0-9/&'\-]*$", w) and (w.isupper() or w.lower() in {"to", "via", "payment"}):
            prefix_words.append(w)
        else:
            break
    return " ".join(prefix_words).strip()


def _parse_business_date(line: str, period_start: Optional[_dt.date], period_end: Optional[_dt.date],
                         prev_date: Optional[_dt.date]) -> Tuple[Optional[_dt.date], str]:
    m = _DATE_ORD_MON_RE.match(line)
    if not m:
        return None, line
    day = int(m.group(1))
    mon_name = m.group(3).lower()
    month = _MONTHS.get(mon_name[:3])
    rem = m.group(4) or ""
    if not month:
        return None, line

    year = _choose_year_for_business_tx(day, month, period_start, period_end, prev_date)
    try:
        return _dt.date(year, month, day), _clean_text(rem)
    except Exception:
        return None, line


def _extract_transactions_business(pdf_path: str) -> List[Dict]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_blob = _clean_text("\n".join([(p.extract_text() or "") for p in pdf.pages[:2]]))
            period_start, period_end = _infer_period_business(first_blob)

            lines: List[str] = []
            for p in pdf.pages:
                t = p.extract_text() or ""
                lines.extend(t.splitlines())
    except Exception:
        return []

    txs: List[Dict] = []
    prev_date: Optional[_dt.date] = None
    prev_balance: Optional[float] = None

    current_date: Optional[_dt.date] = None
    current_desc_parts: List[str] = []
    current_balance: Optional[float] = None
    current_amount_raw: Optional[float] = None

    def flush():
        nonlocal current_date, current_desc_parts, current_balance, current_amount_raw, prev_balance, prev_date

        if current_date is None:
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            return

        desc = _clean_text(" ".join(current_desc_parts))
        if not desc:
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            current_date = None
            return

        tx_type_raw = _extract_type_prefix(desc)

        amount = None
        if current_balance is not None and prev_balance is not None:
            delta = round(float(current_balance) - float(prev_balance), 2)
            amount = delta
        elif current_amount_raw is not None:
            sign = _infer_sign_from_description(desc)
            if sign is None:
                sign = -1
            amount = round(sign * abs(float(current_amount_raw)), 2)
        else:
            amount = 0.0

        ttype, new_desc = _apply_global_type_rules(tx_type_raw, desc)

        txs.append({
            "Date": current_date,
            "Transaction Type": ttype or _title_case_keep_acronyms(tx_type_raw),
            "Description": new_desc,
            "Amount": float(amount),
            "Balance": float(current_balance) if current_balance is not None else None,
        })

        if current_balance is not None:
            prev_balance = float(current_balance)
        elif prev_balance is not None:
            prev_balance = round(float(prev_balance) + float(amount), 2)
        prev_date = current_date

        current_date = None
        current_desc_parts = []
        current_balance = None
        current_amount_raw = None

    in_table = False

    for raw in lines:
        line = _clean_text(raw)
        if not line:
            continue

        low = line.lower()

        if "date" in low and "description" in low and "credits" in low and "debits" in low and "balance" in low:
            in_table = True
            continue

        if not in_table:
            continue

        if "previous statement balance" in low or "balance brought forward" in low:
            vals = _extract_money_values(line)
            if vals:
                prev_balance = float(vals[-1])
            else:
                m_prev = re.search(r"([\-\(]?\s*£?\s*\d{1,3}(?:,\d{3})*\.\d{2}\)?)", line)
                if m_prev:
                    parsed = _parse_money(m_prev.group(1))
                    if parsed is not None:
                        prev_balance = float(parsed)
            continue

        if "current statement balance" in low:
            continue

        if _DATE_ORD_MON_RE.match(line):
            flush()

            d, rem = _parse_business_date(line, period_start, period_end, prev_date)
            if d is None:
                continue
            current_date = d
            current_desc_parts = [rem] if rem else []

            vals = _extract_money_values(line)
            if len(vals) >= 2:
                current_amount_raw = abs(vals[-2])
                current_balance = vals[-1]
                if current_desc_parts:
                    current_desc_parts[-1] = _remove_money_from_text(current_desc_parts[-1])
            elif len(vals) == 1:
                current_amount_raw = abs(vals[-1])
                current_balance = None
                if current_desc_parts:
                    current_desc_parts[-1] = _remove_money_from_text(current_desc_parts[-1])
            else:
                current_amount_raw = None
                current_balance = None

            continue

        if current_date is None:
            continue

        vals = _extract_money_values(line)
        if len(vals) >= 2 and current_balance is None:
            current_amount_raw = abs(vals[-2])
            current_balance = vals[-1]
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
            flush()
        elif len(vals) == 1 and current_amount_raw is None:
            current_amount_raw = abs(vals[-1])
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
        else:
            current_desc_parts.append(line)

    flush()
    return txs


# -----------------------------
# Transactions: ONLINE format
# -----------------------------

def _parse_online_date(line: str) -> Tuple[Optional[_dt.date], str]:
    m = _DATE_DMY_SLASH_FULL_RE.match(line)
    if not m:
        return None, line
    d = int(m.group(1))
    mo = int(m.group(2))
    y = int(m.group(3))
    if y < 100:
        y += 2000
    rem = m.group(4) or ""
    try:
        return _dt.date(y, mo, d), _clean_text(rem)
    except Exception:
        return None, line


def _extract_transactions_online(pdf_path: str) -> List[Dict]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_blob = _clean_text("\n".join([(p.extract_text() or "") for p in pdf.pages[:2]]))
            _ = _infer_period_online(first_blob)

            lines: List[str] = []
            for p in pdf.pages:
                t = p.extract_text() or ""
                lines.extend(t.splitlines())
    except Exception:
        return []

    txs: List[Dict] = []

    in_table = False

    current_date: Optional[_dt.date] = None
    current_desc_parts: List[str] = []
    current_amount_raw: Optional[float] = None
    current_balance: Optional[float] = None

    def flush():
        nonlocal current_date, current_desc_parts, current_amount_raw, current_balance

        if current_date is None:
            current_desc_parts = []
            current_amount_raw = None
            current_balance = None
            return
        desc = _clean_text(" ".join(current_desc_parts))
        desc = _strip_online_junk(desc)

        # Defensive: ignore stray/blank rows that would create phantom transactions.
        if not desc:
            current_date = None
            current_desc_parts = []
            current_amount_raw = None
            current_balance = None
            return

        tx_type_raw = _extract_type_prefix_strict(desc)


        # Placeholder amount; will be overwritten by balance-delta post-processing below
        amount = 0.0
        if current_amount_raw is not None:
            sign = _infer_sign_from_description(desc)
            if sign is None:
                sign = -1
            amount = round(sign * abs(float(current_amount_raw)), 2)

        ttype, new_desc = _apply_global_type_rules(tx_type_raw, desc)

        txs.append({
            "Date": current_date,
            "Transaction Type": ttype or _title_case_keep_acronyms(tx_type_raw),
            "Description": new_desc,
            "Amount": float(amount),
            "Balance": float(current_balance) if current_balance is not None else None,
        })

        current_date = None
        current_desc_parts = []
        current_amount_raw = None
        current_balance = None

    for raw in lines:
        line = _clean_text(raw)
        if not line:
            continue
        low = line.lower()

        if "date" in low and "description" in low and "money in" in low and "money out" in low and "balance" in low:
            in_table = True
            continue
        if not in_table:
            continue

        # Ignore mid-table headers/footers that can create phantom rows
        if _is_online_noise_line(line):
            continue

        # New tx row starts with dd/mm/yyyy
        if _DATE_DMY_SLASH_FULL_RE.match(line):
            # Finish previous transaction when a new date row starts
            flush()

            d, rem = _parse_online_date(line)
            if d is None:
                continue
            current_date = d

            # Parse money values on the date line (usually amount + balance)
            vals = _extract_money_values(line)
            if len(vals) >= 2:
                current_amount_raw = abs(vals[-2])
                current_balance = vals[-1]
                rem = _remove_money_from_text(rem)
            elif len(vals) == 1:
                current_amount_raw = abs(vals[-1])
                current_balance = None
                rem = _remove_money_from_text(rem)
            else:
                current_amount_raw = None
                current_balance = None

            current_desc_parts = [rem] if rem else []
            continue

        # Continuation lines for the current transaction (often contain "ON dd-mm-yyyy" etc)
        if current_date is None:
            continue

        # Avoid polluting descriptions with headers/footers
        if _is_online_noise_line(line):
            continue

        vals = _extract_money_values(line)
        if len(vals) >= 2 and current_balance is None:
            # Some PDFs place money on a following line
            current_amount_raw = abs(vals[-2])
            current_balance = vals[-1]
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
        elif len(vals) == 1 and current_amount_raw is None:
            # Amount only on a following line
            current_amount_raw = abs(vals[-1])
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
        else:
            # Plain description continuation
            current_desc_parts.append(line)

    flush()

    # Post-process amounts using running balances (more reliable than text hints).
    # Online Banking exports are usually reverse chronological (latest first).
    dates = [t.get("Date") for t in txs if isinstance(t.get("Date"), _dt.date)]
    reverse = False
    if dates and dates[0] > dates[-1]:
        reverse = True

    if reverse:
        for i in range(0, len(txs) - 1):
            b0 = txs[i].get("Balance")
            b1 = txs[i + 1].get("Balance")
            if b0 is not None and b1 is not None:
                try:
                    txs[i]["Amount"] = round(float(b0) - float(b1), 2)
                except Exception:
                    pass
    else:
        prev_b = None
        for i in range(0, len(txs)):
            b = txs[i].get("Balance")
            if prev_b is not None and b is not None:
                try:
                    txs[i]["Amount"] = round(float(b) - float(prev_b), 2)
                except Exception:
                    pass
            if b is not None:
                try:
                    prev_b = float(b)
                except Exception:
                    prev_b = None

    return txs


def _extract_transactions_personal(pdf_path: str) -> List[Dict]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_blob = _clean_text("\n".join([(p.extract_text() or "") for p in pdf.pages[:2]]))
            period_start, period_end = _infer_period_business(first_blob)

            lines: List[str] = []
            for p in pdf.pages:
                t = p.extract_text() or ""
                lines.extend(t.splitlines())
    except Exception:
        return []

    txs: List[Dict] = []
    prev_date: Optional[_dt.date] = None
    prev_balance: Optional[float] = None

    current_date: Optional[_dt.date] = None
    current_desc_parts: List[str] = []
    current_balance: Optional[float] = None
    current_amount_raw: Optional[float] = None

    def flush():
        nonlocal current_date, current_desc_parts, current_balance, current_amount_raw, prev_balance, prev_date
        if current_date is None:
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            return

        desc = _clean_text(" ".join(current_desc_parts))
        low = desc.lower().replace(" ", "")
        if "balancebroughtforward" in low or "balancecarriedforward" in low or "carriedforwardtonextstatement" in low:
            if current_balance is not None and "balancebroughtforward" in low:
                prev_balance = float(current_balance)
            current_date = None
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            return

        tx_type_raw = _extract_type_prefix(desc)
        amount = None
        if current_balance is not None and prev_balance is not None:
            amount = round(float(current_balance) - float(prev_balance), 2)
        elif current_amount_raw is not None:
            sign = _infer_sign_from_description(desc)
            if sign is None:
                sign = -1
            amount = round(sign * abs(float(current_amount_raw)), 2)
        else:
            amount = 0.0

        ttype, new_desc = _apply_global_type_rules(tx_type_raw, desc)
        txs.append({
            "Date": current_date,
            "Transaction Type": ttype or _title_case_keep_acronyms(tx_type_raw),
            "Description": new_desc,
            "Amount": float(amount),
            "Balance": float(current_balance) if current_balance is not None else None,
        })

        if current_balance is not None:
            prev_balance = float(current_balance)
        elif prev_balance is not None:
            prev_balance = round(float(prev_balance) + float(amount), 2)
        prev_date = current_date

        current_date = None
        current_desc_parts = []
        current_balance = None
        current_amount_raw = None

    saw_transactions_section = False
    in_table = False

    for raw in lines:
        line = _clean_text(raw)
        if not line:
            continue
        low = line.lower()

        if "your transactions" in low:
            saw_transactions_section = True

        if not saw_transactions_section:
            continue

        if "date" in low and "description" in low and ("moneyin" in low or "money in" in low) and ("moneyout" in low or "money out" in low) and "balance" in low:
            in_table = True
            continue

        if not in_table:
            continue

        if _DATE_ORD_MON_RE.match(line):
            flush()
            d, rem = _parse_business_date(line, period_start, period_end, prev_date)
            if d is None:
                continue
            current_date = d
            current_desc_parts = [rem] if rem else []

            vals = _extract_money_values(line)
            if len(vals) >= 2:
                current_amount_raw = abs(vals[-2])
                current_balance = vals[-1]
                if current_desc_parts:
                    current_desc_parts[-1] = _remove_money_from_text(current_desc_parts[-1])
            elif len(vals) == 1:
                current_amount_raw = abs(vals[-1])
                current_balance = vals[-1]
                if current_desc_parts:
                    current_desc_parts[-1] = _remove_money_from_text(current_desc_parts[-1])
            else:
                current_amount_raw = None
                current_balance = None
            continue

        if current_date is None:
            continue

        vals = _extract_money_values(line)
        low_ns = low.replace(" ", "")
        if "balancebroughtforward" in low_ns and vals:
            prev_balance = float(vals[-1])
            current_date = None
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            continue
        if "balancecarriedforward" in low_ns or "carriedforwardtonextstatement" in low_ns:
            current_date = None
            current_desc_parts = []
            current_balance = None
            current_amount_raw = None
            continue

        if len(vals) >= 2 and current_balance is None:
            current_amount_raw = abs(vals[-2])
            current_balance = vals[-1]
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
            flush()
        elif len(vals) == 1 and current_amount_raw is None:
            current_amount_raw = abs(vals[-1])
            desc_part = _remove_money_from_text(line)
            if desc_part:
                current_desc_parts.append(desc_part)
        else:
            current_desc_parts.append(line)

    flush()
    return txs


def _extract_transactions_online_creditcard(pdf_path: str) -> List[Dict]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            lines: List[str] = []
            for p in pdf.pages:
                t = p.extract_text() or ""
                lines.extend(t.splitlines())
    except Exception:
        return []

    events: List[Dict] = []
    in_table = False

    for raw in lines:
        line = _clean_text(raw)
        if not line:
            continue
        low = line.lower()

        if "date" in low and "card number" in low and "description" in low and "money in" in low and "money out" in low:
            in_table = True
            continue
        if not in_table:
            continue
        if _is_online_noise_line(line):
            continue
        if not _DATE_DMY_SLASH_FULL_RE.match(line):
            continue

        d, rem = _parse_online_date(line)
        if d is None:
            continue

        vals = _extract_money_values(line)
        amount_raw = abs(vals[-1]) if vals else None
        desc = _remove_money_from_text(rem)

        events.append({
            "Date": d,
            "Description": desc,
            "AmountRaw": amount_raw,
            "IsInitial": "initial balance" in desc.lower() and len(vals) == 1,
        })

    if not events:
        return []

    dates = [e["Date"] for e in events]
    if dates and dates[0] > dates[-1]:
        events = list(reversed(events))

    opening = None
    filtered: List[Dict] = []
    for e in events:
        if e["IsInitial"] and opening is None and e["AmountRaw"] is not None:
            opening = -abs(float(e["AmountRaw"]))
            continue
        filtered.append(e)

    txs: List[Dict] = []
    running = opening
    for e in filtered:
        desc = re.sub(r'^\*\*\s*\d{2,6}\s*', '', e["Description"] or '').strip()
        amount_raw = float(e["AmountRaw"]) if e["AmountRaw"] is not None else 0.0
        dlow = desc.lower()
        if any(k in dlow for k in ["payment received", "refund", "cashback", "credit"]):
            amount = abs(amount_raw)
            tx_type_raw = "Credit"
        else:
            amount = -abs(amount_raw)
            tx_type_raw = "Card Payment"

        balance = None
        if running is not None:
            running = round(float(running) + float(amount), 2)
            balance = running

        ttype, new_desc = _apply_global_type_rules(tx_type_raw, desc)
        txs.append({
            "Date": e["Date"],
            "Transaction Type": ttype or _title_case_keep_acronyms(tx_type_raw),
            "Description": new_desc,
            "Amount": float(amount),
            "Balance": float(balance) if balance is not None else None,
        })

    return txs


# -----------------------------
# Public API: extract_transactions
# -----------------------------

def extract_transactions(pdf_path) -> List[Dict]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page_text = pdf.pages[0].extract_text() or ""
    except Exception:
        first_page_text = ""

    kind = _detect_statement_kind(first_page_text)

    if kind == "ONLINE":
        txs = _extract_transactions_online(pdf_path)

        # The Online Banking export lists newest first. We parse in appearance order first (for stability),
        # then return oldest -> newest for downstream reconciliation and Excel output.
        dates = [t.get("Date") for t in txs if isinstance(t.get("Date"), _dt.date)]
        if dates and dates[0] > dates[-1]:
            txs = list(reversed(txs))
    elif kind == "ONLINE_CREDITCARD":
        txs = _extract_transactions_online_creditcard(pdf_path)
    elif kind == "PERSONAL":
        txs = _extract_transactions_personal(pdf_path)
    else:
        txs = _extract_transactions_business(pdf_path)

    cleaned: List[Dict] = []
    for t in txs:
        if not isinstance(t.get("Date"), _dt.date):
            continue

        t["Transaction Type"] = (t.get("Transaction Type") or "").strip()
        t["Description"] = (t.get("Description") or "").strip()

        try:
            t["Amount"] = float(t.get("Amount", 0.0))
        except Exception:
            t["Amount"] = 0.0

        bal = t.get("Balance")
        if bal is not None:
            try:
                t["Balance"] = float(bal)
            except Exception:
                t["Balance"] = None
        else:
            t["Balance"] = None

        cleaned.append(t)

    return cleaned

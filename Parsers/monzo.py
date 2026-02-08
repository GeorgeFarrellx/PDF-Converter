# Version: monzo.py
# Monzo (Business Account) statement parser - text-based PDFs (no OCR)
# Defines:
#   extract_transactions(pdf_path) -> list[dict]
#   extract_statement_balances(pdf_path) -> dict
#   extract_account_holder_name(pdf_path) -> str

from __future__ import annotations

import os
import re
import datetime as _dt
from typing import Optional, List, Dict, Tuple

import pdfplumber


# ----------------------------
# Helpers
# ----------------------------

_DATE_DDMMYYYY_RE = re.compile(r"^\s*(\d{2}/\d{2}/\d{4})\s*(.*)\s*$")
_DATE_DDMM_RE = re.compile(r"^\s*(\d{2}/\d{2})\s*(.*)\s*$")

# Trailing amount + balance at end of line (both money-like)
_TRAIL_AMT_BAL_RE = re.compile(
    r"(?P<amount>[+-]?\d[\d,]*\.\d{2})\s+(?P<balance>[+-]?\d[\d,]*\.\d{2})\s*$"
)

# Sometimes you may encounter amount only at end (balance missing)
_TRAIL_AMT_ONLY_RE = re.compile(r"(?P<amount>[+-]?\d[\d,]*\.\d{2})\s*$")

# Statement date range on first page
_RANGE_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})")

# "Business Account balance" value near top (e.g. -£254.67)
_BUSINESS_BAL_RE = re.compile(r"([+-]?)\s*£\s*([\d,]+\.\d{2})")


def _money_to_float(s: str) -> Optional[float]:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None

    # handle parentheses negatives e.g. (1,234.56)
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()

    # remove currency symbols/spaces
    s = s.replace("£", "").replace(",", "").strip()

    # keep leading sign if present
    try:
        val = float(s)
    except ValueError:
        return None

    if neg:
        val = -abs(val)
    return val


def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _title_case_preserve_acronyms(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s
    # Basic Title Case; keeps P2P reasonably intact
    return " ".join([w[:1].upper() + w[1:].lower() if not re.search(r"\d", w) else w.upper() for w in s.split()])


def _extract_statement_period(pdf_path: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
    """
    Returns (start_date, end_date) parsed from statement header like '31/03/2025 - 07/10/2025'
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None, None
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None, None

    m = _RANGE_RE.search(text)
    if not m:
        return None, None

    try:
        d1 = _dt.datetime.strptime(m.group(1), "%d/%m/%Y").date()
        d2 = _dt.datetime.strptime(m.group(2), "%d/%m/%Y").date()
        return d1, d2
    except Exception:
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
        start, end = _extract_statement_period(pdf_path)
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _infer_year_for_ddmm(ddmm: str, period_start: Optional[_dt.date], period_end: Optional[_dt.date]) -> Optional[int]:
    """
    For statements that might show dates as DD/MM without a year.
    Uses statement period to choose the correct year, including Dec -> Jan rollovers.
    """
    if not ddmm or "/" not in ddmm:
        return None
    try:
        day, month = ddmm.split("/")
        day_i = int(day)
        month_i = int(month)
    except Exception:
        return None

    if period_start and period_end:
        # try both boundary years, choose the one that lands inside the period.
        candidate_years = sorted({period_start.year, period_end.year})
        for y in candidate_years:
            try:
                d = _dt.date(y, month_i, day_i)
            except Exception:
                continue
            if period_start <= d <= period_end:
                return y

        # If period spans year boundary and date is "Jan/Feb" while end is in Jan,
        # or date is "Nov/Dec" while start is in Dec, pick accordingly.
        if period_start.year != period_end.year:
            # heuristic: months near the start date likely belong to start year
            # months near the end date likely belong to end year
            if month_i >= period_start.month:
                return period_start.year
            return period_end.year

        return period_end.year

    # fallback: current year
    return _dt.date.today().year


def _extract_business_account_balance_top(pdf_path: str) -> Optional[float]:
    """
    Attempts to extract the 'Business Account balance' shown near the top of page 1.
    Example from sample: '-£254.67'
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None

    # Try to locate "Business Account balance" and grab a nearby currency number.
    idx = text.find("Business Account balance")
    if idx == -1:
        return None

    # Look back and forward around that phrase for a currency value (Monzo often prints it above).
    window = text[max(0, idx - 200): idx + 200]
    # Find all currency occurrences; pick the one closest to the phrase (heuristic).
    matches = list(_BUSINESS_BAL_RE.finditer(window))
    if not matches:
        return None

    # choose the last match before the phrase if possible, else first after
    phrase_pos = window.find("Business Account balance")
    best = None
    best_dist = None
    for m in matches:
        dist = abs(m.start() - phrase_pos)
        if best_dist is None or dist < best_dist:
            best = m
            best_dist = dist

    if not best:
        return None

    sign = best.group(1) or ""
    num = best.group(2)
    val = _money_to_float(f"{sign}{num}")
    return val


def _iter_all_lines(pdf_path: str) -> List[str]:
    lines: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if not txt:
                continue
            for raw in txt.splitlines():
                s = raw.rstrip("\n")
                if s is not None:
                    lines.append(s)
    return lines


def _looks_like_table_header(line: str) -> bool:
    l = (line or "").strip()
    if not l:
        return True
    # Table headers in your sample: "Date Description (GBP) Amount (GBP) Balance"
    if l.lower().startswith("date ") and "amount" in l.lower() and "balance" in l.lower():
        return True
    if "important information about compensation" in l.lower():
        return True
    return False


def _normalize_type_and_description(raw_type: str, raw_desc: str) -> Tuple[str, str]:
    """
    Applies the global transaction type rules and description cleaning.
    """
    desc = _clean_spaces(raw_desc)
    tx_type = (raw_type or "").strip()

    # Returned Direct Debit rule
    if "returned direct debit" in (desc.lower() + " " + tx_type.lower()):
        tx_type = "Direct Debit"
        if not desc.lower().startswith("returned direct debit"):
            desc = "Returned Direct Debit " + desc
        # Keep that prefix; do not remove it
        return tx_type, desc

    # Card Payment overrides
    desc_lower = desc.lower()
    if "applepay" in desc_lower or "clearpay" in desc_lower or "contactless" in desc_lower or re.search(r"\bGB$", desc.strip()):
        tx_type = "Card Payment"
    else:
        tx_type = _title_case_preserve_acronyms(tx_type) if tx_type else "Other"

    # Remove any leading type prefix from description (rare for Monzo, but global rule)
    if tx_type and tx_type.lower() != "direct debit":
        if desc_lower.startswith(tx_type.lower() + " "):
            desc = desc[len(tx_type) + 1:].lstrip()

    return tx_type, desc


def _extract_type_from_description(desc: str) -> Tuple[str, str]:
    """
    Monzo commonly encodes type in parentheses e.g.:
      "Joseph Lombardi (P2P Payment)"
      "V12 RETAIL FINANCE (Direct Debit) Reference: 021..."
    We treat the parenthetical as the bank's transaction type and remove it from description.
    """
    s = _clean_spaces(desc)

    # Find the first "(...)" that looks like a type token
    # Prefer one that contains keywords commonly used by Monzo.
    type_candidates = []
    for m in re.finditer(r"\(([^)]+)\)", s):
        inner = m.group(1).strip()
        if not inner:
            continue
        type_candidates.append((m.start(), m.end(), inner))

    chosen = None
    for start, end, inner in type_candidates:
        inner_l = inner.lower()
        if any(k in inner_l for k in ["direct debit", "faster payments", "bank transfer", "p2p", "international", "card"]):
            chosen = (start, end, inner)
            break

    if chosen is None and type_candidates:
        # fallback: take the first parenthetical
        chosen = type_candidates[0]

    if not chosen:
        return "", s

    start, end, inner = chosen
    # Remove that parenthetical
    cleaned = (s[:start] + s[end:]).strip()
    cleaned = _clean_spaces(cleaned)
    return inner, cleaned


# ----------------------------
# Core parsing
# ----------------------------

def extract_transactions(pdf_path: str) -> List[Dict]:
    """
    Extracts transactions in the order they appear in the PDF (Monzo statements are typically reverse-chronological).
    Each transaction dict keys:
      Date (datetime.date)
      Transaction Type (str)
      Description (str)
      Amount (float)
      Balance (float or None)
    """
    period_start, period_end = _extract_statement_period(pdf_path)
    lines = _iter_all_lines(pdf_path)

    # We'll stop once we hit boilerplate legal/compensation sections to avoid false positives.
    stop_markers = [
        "Monzo Bank Limited",
        "Important information about compensation",
        "FSCS",
        "Financial Services Compensation Scheme",
    ]

    txns_raw = []

    current_date: Optional[_dt.date] = None
    current_desc_parts: List[str] = []
    current_amount_str: Optional[str] = None
    current_balance_str: Optional[str] = None
    current_amount_had_sign: Optional[bool] = None

    def flush_current():
        nonlocal current_date, current_desc_parts, current_amount_str, current_balance_str, current_amount_had_sign

        if current_date is None:
            return
        if current_amount_str is None:
            return  # incomplete

        amount = _money_to_float(current_amount_str)
        balance = _money_to_float(current_balance_str) if current_balance_str is not None else None

        if amount is None:
            return

        raw_desc = _clean_spaces(" ".join([p for p in current_desc_parts if p is not None and str(p).strip()]))

        raw_type, cleaned_desc = _extract_type_from_description(raw_desc)
        tx_type, final_desc = _normalize_type_and_description(raw_type, cleaned_desc)

        txns_raw.append(
            {
                "Date": current_date,
                "Transaction Type": tx_type,
                "Description": final_desc,
                "Amount": float(amount),
                "Balance": float(balance) if balance is not None else None,
                "_amount_had_sign": bool(current_amount_had_sign) if current_amount_had_sign is not None else False,
            }
        )

        # reset
        current_date = None
        current_desc_parts = []
        current_amount_str = None
        current_balance_str = None
        current_amount_had_sign = None

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if any(marker.lower() in line.lower() for marker in stop_markers):
            # don't flush here; table is finished
            break

        if _looks_like_table_header(line):
            continue

        # If line begins with DD/MM/YYYY
        m_full = _DATE_DDMMYYYY_RE.match(line)
        if m_full:
            # new transaction starts; flush any previous
            flush_current()

            date_str = m_full.group(1)
            rest = (m_full.group(2) or "").strip()

            try:
                current_date = _dt.datetime.strptime(date_str, "%d/%m/%Y").date()
            except Exception:
                current_date = None
                continue

            # If rest contains trailing amount+balance, extract it
            if rest:
                m_tail = _TRAIL_AMT_BAL_RE.search(rest)
                if m_tail:
                    current_amount_str = m_tail.group("amount")
                    current_balance_str = m_tail.group("balance")
                    current_amount_had_sign = current_amount_str.strip().startswith(("+", "-"))

                    desc_only = rest[: m_tail.start()].strip()
                    if desc_only:
                        current_desc_parts.append(desc_only)
                    # we can flush immediately
                    flush_current()
                else:
                    # Date line but not completed yet; could be wrapped description or date-only style
                    current_desc_parts.append(rest)
            else:
                # date-only line (seen in international txns)
                pass

            continue

        # If line begins with DD/MM (no year) - support year inference
        m_short = _DATE_DDMM_RE.match(line)
        if m_short:
            flush_current()
            ddmm = m_short.group(1)
            rest = (m_short.group(2) or "").strip()

            year = _infer_year_for_ddmm(ddmm, period_start, period_end)
            if year is None:
                continue

            try:
                current_date = _dt.datetime.strptime(f"{ddmm}/{year}", "%d/%m/%Y").date()
            except Exception:
                current_date = None
                continue

            if rest:
                m_tail = _TRAIL_AMT_BAL_RE.search(rest)
                if m_tail:
                    current_amount_str = m_tail.group("amount")
                    current_balance_str = m_tail.group("balance")
                    current_amount_had_sign = current_amount_str.strip().startswith(("+", "-"))
                    desc_only = rest[: m_tail.start()].strip()
                    if desc_only:
                        current_desc_parts.append(desc_only)
                    flush_current()
                else:
                    current_desc_parts.append(rest)
            continue

        # Continuation lines: if we're currently building a txn, try to detect trailing amount/balance
        if current_date is not None:
            m_tail2 = _TRAIL_AMT_BAL_RE.search(line)
            if m_tail2 and current_amount_str is None:
                current_amount_str = m_tail2.group("amount")
                current_balance_str = m_tail2.group("balance")
                current_amount_had_sign = current_amount_str.strip().startswith(("+", "-"))

                desc_only = line[: m_tail2.start()].strip()
                if desc_only:
                    current_desc_parts.append(desc_only)

                flush_current()
            else:
                # normal continuation text (e.g., Reference:, Amount: USD..., etc.)
                current_desc_parts.append(line)

    # In case last record wasn't flushed
    flush_current()

    # Optional: balance-delta correction when amount lacks sign and balances exist
    # Monzo statements in this format are typically reverse-chronological. If so:
    #   amount_i ≈ balance_i - balance_{i+1}
    for i in range(len(txns_raw) - 1):
        a = txns_raw[i]
        b = txns_raw[i + 1]
        if a.get("Balance") is None or b.get("Balance") is None:
            continue
        if a.get("_amount_had_sign", True):
            continue  # already has explicit sign

        derived = float(a["Balance"]) - float(b["Balance"])
        if abs(derived - float(a["Amount"])) > 0.01:
            a["Amount"] = derived

    # Remove internal key
    for t in txns_raw:
        if "_amount_had_sign" in t:
            del t["_amount_had_sign"]

    return txns_raw



def extract_statement_balances(pdf_path: str) -> Dict[str, Optional[float]]:
    """
    Returns:
      {"start_balance": float|None, "end_balance": float|None}

    Monzo Business statements (in this template) typically show a "Business Account balance" on page 1
    which matches the ending running balance. They may not explicitly show a start balance, so we infer:
      start_balance = (earliest_txn_balance) - (earliest_txn_amount)
    """
    txns = extract_transactions(pdf_path)

    end_balance = _extract_business_account_balance_top(pdf_path)

    # fallback: last transaction balance in appearance order (usually most recent)
    if end_balance is None and txns:
        if txns[0].get("Balance") is not None:
            end_balance = float(txns[0]["Balance"])

    start_balance = None
    if txns:
        # earliest transaction is the last one in the statement (Monzo order is usually reverse-chronological)
        last = txns[-1]
        if last.get("Balance") is not None and last.get("Amount") is not None:
            start_balance = float(last["Balance"]) - float(last["Amount"])

    return {"start_balance": start_balance, "end_balance": end_balance}



def extract_account_holder_name(pdf_path: str) -> str:
    """
    Best-effort extraction of the account holder / business name.

    Monzo PDFs commonly have a two-column layout on page 1 (left: name/address, right: disclaimer/balances).
    pdfplumber's text extraction can interleave both columns onto the same line, e.g.:
      "Joseph Lombardi This statement doesn't include transfers ..."
      "Brummie Joe Media account and your Pots ..."

    This function:
    - anchors on the statement period line
    - cleans interleaved right-column fragments from each line
    - then extracts the name block before the address (first digit/postcode)

    For your sample, the intended return is the business name line:
      Brummie Joe Media
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""

    raw_lines = [l.rstrip() for l in text.splitlines() if l and l.strip()]
    if not raw_lines:
        return ""

    # Markers that usually belong to the *right* column on page 1.
    # If they appear on a line, we truncate at their start to recover the left-column content.
    right_col_markers = [
        "This statement",
        "account and your Pots",
        "To include them",
        "switch to",
        "non-combined statements",
        "combined statements",
        "Business Account balance",
        "Total outgoings",
        "Total deposits",
        "(Including internal",
        "(Excluding internal",
    ]

    money_on_line_re = re.compile(r"\s[+-]£")

    def clean_line(line: str) -> str:
        s = (line or "").strip()
        if not s:
            return ""

        # Truncate at the earliest right-column marker
        lower = s.lower()
        cut = None
        for m in right_col_markers:
            pos = lower.find(m.lower())
            if pos != -1:
                if cut is None or pos < cut:
                    cut = pos
        if cut is not None:
            s = s[:cut].strip()

        # Truncate at embedded balance values (e.g. "Tamworth -£254.67")
        m2 = money_on_line_re.search(s)
        if m2:
            s = s[:m2.start()].strip()

        return _clean_spaces(s)

    lines = [clean_line(l) for l in raw_lines]
    lines = [l for l in lines if l]
    if not lines:
        return ""

    # Anchor on statement period
    range_idx = None
    for i, l in enumerate(lines):
        if _RANGE_RE.search(l):
            range_idx = i
            break

    # UK postcode pattern (broad but practical)
    uk_postcode_re = re.compile(r"\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b", re.IGNORECASE)

    def is_noise(line: str) -> bool:
        ll = (line or "").strip().lower()
        if not ll:
            return True
        if "business account statement" in ll:
            return True
        if ll.startswith("date ") and "amount" in ll and "balance" in ll:
            return True
        if ll in {"united kingdom"}:
            return True
        if ll.startswith("sort code") or ll.startswith("account number") or ll.startswith("bic") or ll.startswith("iban"):
            return True
        return False

    def looks_like_person_name(s: str) -> bool:
        # simple heuristic: 2-4 words, all alpha, each starts with capital
        parts = [p for p in s.split() if p]
        if not (2 <= len(parts) <= 4):
            return False
        for p in parts:
            if not re.match(r"^[A-Za-z'-]+$", p):
                return False
            if not p[:1].isupper():
                return False
        return True

    if range_idx is not None:
        # Take a reasonable slice after the period line and stop at postcode or at the start of the transaction table.
        end_idx = min(len(lines), range_idx + 60)

        for j in range(range_idx + 1, end_idx):
            if uk_postcode_re.search(lines[j]):
                end_idx = j
                break
            if lines[j].lower().startswith("date ") and "amount" in lines[j].lower():
                end_idx = j
                break

        segment = [s for s in lines[range_idx + 1:end_idx] if s and not is_noise(s)]

        # Address typically starts at the first line containing a digit (e.g. "37 Hedging Lane")
        addr_start = None
        for k, s in enumerate(segment):
            if re.search(r"\d", s):
                addr_start = k
                break

        name_block = segment[:addr_start] if addr_start is not None else segment
        name_block = [s for s in name_block if s and not is_noise(s) and not re.search(r"\d", s)]

        if name_block:
            # Prefer a business-like second line if present (Monzo often prints Person then Business)
            if len(name_block) >= 2:
                first, second = name_block[0], name_block[1]
                second_l = second.lower()
                if any(tok in second_l for tok in [" ltd", " limited", " llp", " plc", " inc", " media", " studio", " company", " services"]):
                    return second
                if looks_like_person_name(first) and second.lower() != first.lower():
                    return second
            return name_block[0]

    # Fallback: pick first plausible non-noise, non-address-like line
    for l in lines:
        if is_noise(l):
            continue
        if re.search(r"\d", l):
            continue
        if any(k in l.lower() for k in ["lane", "road", "street", "avenue", "close", "court", "drive", "place"]):
            continue
        return l

    return ""


# ----------------------------
# Notes for maintainers
# ----------------------------
# This parser targets the Monzo Business statement template where the transaction table lines look like:
#   DD/MM/YYYY <description...> <amount> <balance>
# and where descriptions may wrap onto multiple lines, including cases where the date appears alone on a line
# and the amount/balance appear on a later line (international and pot interest examples).
#
# Global transaction type rules are applied after extraction.
# Statement end balance is read from "Business Account balance" on page 1 when possible; start balance is inferred
# from the earliest transaction (last row in appearance order): start = balance - amount.

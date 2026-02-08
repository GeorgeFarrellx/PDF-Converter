# Version: tsb.py
# TSB (Spend & Save / similar) text-based PDF parser (NO OCR)
# Requires: pdfplumber

from __future__ import annotations

import os
import re
import datetime as _dt
from typing import List, Dict, Optional, Tuple

import pdfplumber


# -----------------------------
# Helpers
# -----------------------------

_FLOAT_RE = re.compile(r'(?<!\d)(-?\d{1,3}(?:,\d{3})*\.\d{2})(?!\d)')
_DATE_START_RE = re.compile(r'^(?P<d>\d{2})\s+(?P<m>[A-Za-z]{3})\s*(?P<y>\d{2})?\b')

# Common payment types that appear in the "Payment type" column (TSB)
_KNOWN_TYPES = [
    "FASTER PAYMENT",
    "DIRECT DEBIT",
    "DIRECT CREDIT",
    "STANDING ORDER",
    "TRANSFER TO",
    "TRANSFER FROM",
    "CHEQUE",
    "CASH WITHDRAWAL",
    "CASH DEPOSIT",
    "INTEREST",
    "CHARGE",
    "CARD PAYMENT",
    "BANK GIRO CREDIT",
    "BILL PAYMENT",
]
_KNOWN_TYPES_SORTED = sorted(_KNOWN_TYPES, key=len, reverse=True)


def _extract_text(page) -> str:
    # These tolerances preserve spaces/columns much better on the provided PDFs
    return page.extract_text(x_tolerance=2, y_tolerance=2) or ""


def _parse_money(val: str) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    # TSB PDFs often extract the currency symbol as "[" instead of "£"
    s = s.replace("£", "").replace("[", "").replace("]", "").replace(",", "").strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    try:
        v = float(s)
    except Exception:
        return None
    return -v if neg else v


def _parse_any_date_str(s: str) -> Optional[_dt.date]:
    s = s.strip()
    for fmt in ("%d %b %Y", "%d %B %Y", "%d%b%Y", "%d%B%Y"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def _parse_statement_period(first_page_text: str) -> Tuple[Optional[_dt.date], Optional[_dt.date]]:
    # Example: "Effective from: 07 May 2024 to 05 June 2024"
    m = re.search(
        r"Effective\s+from:\s*(\d{2}\s+[A-Za-z]{3,9}\s+\d{4})\s+to\s+(\d{2}\s+[A-Za-z]{3,9}\s+\d{4})",
        first_page_text,
    )
    if m:
        return _parse_any_date_str(m.group(1)), _parse_any_date_str(m.group(2))

    # Fallback if PDF text collapses whitespace
    compact = re.sub(r"\s+", "", first_page_text)
    m2 = re.search(r"Effectivefrom:(\d{2}[A-Za-z]{3,9}\d{4})to(\d{2}[A-Za-z]{3,9}\d{4})", compact)
    if m2:
        return _parse_any_date_str(m2.group(1)), _parse_any_date_str(m2.group(2))

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
            text = pdf.pages[0].extract_text() if pdf.pages else ""
        start, end = _parse_statement_period(text or "")
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
    except Exception:
        return None, None


def _parse_statement_balances(first_page_text: str) -> Tuple[Optional[float], Optional[float]]:
    # Example (as extracted by pdfplumber):
    #   "Balance on 06 August 2024 [4,118.25"  ("[" often appears instead of "£")
    matches = re.findall(
        r"Balance\s+on\s+(\d{2}\s+[A-Za-z]{3,9}\s+\d{4})\s+[\[£]?\s*([\d,]+\.\d{2})",
        first_page_text,
    )
    vals: List[float] = []
    for _, v in matches:
        mv = _parse_money(v)
        if mv is not None:
            vals.append(mv)
    if vals:
        return vals[0], vals[-1]

    # Fallback if PDF text collapses whitespace
    compact = re.sub(r"\s+", "", first_page_text)
    matches2 = re.findall(
        r"Balanceon(\d{2}[A-Za-z]{3,9}\d{4})[\[£]?([\d,]+\.\d{2})",
        compact,
    )
    vals2: List[float] = []
    for _, v in matches2:
        mv = _parse_money(v)
        if mv is not None:
            vals2.append(mv)
    if vals2:
        return vals2[0], vals2[-1]

    return None, None


def _infer_date(day: int, mon_abbr: str, yy: Optional[int], period_start: Optional[_dt.date], period_end: Optional[_dt.date]) -> Optional[_dt.date]:
    try:
        mon = _dt.datetime.strptime(mon_abbr, "%b").month
    except Exception:
        return None

    if yy is not None:
        year = 2000 + yy if yy < 80 else 1900 + yy
        return _dt.date(year, mon, day)

    # If year missing, infer from statement period (handles Dec→Jan rollover)
    if period_start and period_end:
        if period_start.year == period_end.year:
            year = period_start.year
        else:
            # spans year boundary:
            # months >= start month -> start year, otherwise -> end year
            year = period_start.year if mon >= period_start.month else period_end.year
        return _dt.date(year, mon, day)

    if period_end:
        return _dt.date(period_end.year, mon, day)
    if period_start:
        return _dt.date(period_start.year, mon, day)

    return None


def _match_known_type(text: str) -> Tuple[Optional[str], Optional[str]]:
    up = text.upper()
    for t in _KNOWN_TYPES_SORTED:
        if up.startswith(t):
            return t, text[len(t):].strip()
    return None, None


def _normalize_type_and_desc(tx_type: str, desc: str) -> Tuple[str, str]:
    tx_type = (tx_type or "").strip()
    desc = (desc or "").strip()

    # Returned Direct Debit rule
    if desc.lower().startswith("returned direct debit") or tx_type.lower().startswith("returned direct debit"):
        tx_type = "Direct Debit"
        if not desc.lower().startswith("returned direct debit"):
            desc = ("Returned Direct Debit " + desc).strip()
        return tx_type, desc

    desc_lower = desc.lower()

    # Card Payment overrides
    if ("applepay" in desc_lower) or ("clearpay" in desc_lower) or ("contactless" in desc_lower) or desc.endswith("GB"):
        tx_type = "Card Payment"
    else:
        tx_type = tx_type.title() if tx_type else tx_type

    # Remove type prefix from Description (unless Returned Direct Debit which is handled above)
    if tx_type and desc_lower.startswith(tx_type.lower()):
        desc = desc[len(tx_type):].strip(" -:")

    return tx_type, desc


def _remove_last_n_floats(text: str, n: int) -> str:
    """
    Remove the last n float-like tokens from a string using spans,
    even if other tokens (e.g., card digits) trail after them.
    """
    spans = [m.span() for m in _FLOAT_RE.finditer(text)]
    if len(spans) < n:
        return text
    cut = spans[-n:]
    s = text
    # remove right-to-left to keep indices stable
    for start, end in reversed(cut):
        s = s[:start] + " " + s[end:]
    return re.sub(r"\s+", " ", s).strip()


def _split_type_details(rest: str) -> Tuple[str, str]:
    rest = re.sub(r"\s+", " ", (rest or "").strip())
    if not rest:
        return "", ""

    # Match known payment type prefixes (best case)
    t, after = _match_known_type(rest)
    if t:
        return t, after

    # Card-style rows often have "... CD 4334" but sometimes "CD" appears before amounts,
    # with the final "4334" trailing after the amounts.
    # If we see a trailing 4-digit token, attach it to "CD" if present.
    m_tail_digits = re.search(r"\b(\d{4})\s*$", rest)
    if m_tail_digits and re.search(r"\bCD\b", rest, flags=re.IGNORECASE):
        digits = m_tail_digits.group(1)
        rest_wo_digits = re.sub(r"\b" + re.escape(digits) + r"\s*$", "", rest).strip()

        # Split at CD
        m_cd = re.search(r"\bCD\b", rest_wo_digits, flags=re.IGNORECASE)
        if m_cd:
            left = rest_wo_digits[:m_cd.start()].strip()
            return left, f"CD {digits}"

    # If we do have a normal "... CD 4334" pattern
    m = re.search(r"\sCD\s+\d{4}\b", rest, flags=re.IGNORECASE)
    if m:
        idx = m.start()
        left = rest[:idx].strip()
        right = rest[idx:].strip()
        return left, right

    # Fallback: treat everything as the type (merchant often appears here on TSB)
    return rest, ""


def _extract_rows(pdf_path: str) -> List[str]:
    """
    Extract raw transaction 'rows' by scanning pages for the 'Your Transactions' section.
    Rows are reconstructed by:
      - starting a new row when a line begins with a date
      - appending subsequent lines as continuations
      - handling "date omitted" patterns by starting a new row with the last seen date
        when a line looks like a full transaction (ends with float tokens)
    """
    rows: List[str] = []
    last_date_str: Optional[str] = None
    in_section = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = _extract_text(page)
            lines = [ln.rstrip() for ln in text.splitlines()]

            for ln in lines:
                if "Your Transactions" in ln:
                    in_section = True
                    continue
                if not in_section:
                    continue

                # Skip header row
                if re.search(r"^Date\s+Payment type\s+Details\s+Money Out", ln):
                    continue

                # Stop when we hit non-transaction sections
                if "Monthly Maximum Charge" in ln:
                    in_section = False
                    continue

                # Skip obvious footers
                if "TSB Bank plc Registered Office" in ln:
                    continue
                if "Continued on next page" in ln:
                    continue
                if not ln.strip():
                    continue

                s = ln.strip()

                dm = _DATE_START_RE.match(s)
                if dm:
                    # date present -> start new row
                    d = dm.group("d")
                    m = dm.group("m")
                    y = dm.group("y")
                    last_date_str = f"{d} {m}" + (f" {y}" if y else "")
                    rows.append(s)
                    continue

                # Possible omitted-date new transaction row:
                # If it contains floats and starts with a payment-type-ish token, treat as new row with last_date_str.
                if last_date_str and _FLOAT_RE.search(s):
                    first_word = (s.split()[0] if s.split() else "").upper()
                    if first_word in {"FASTER", "DIRECT", "STANDING", "TRANSFER", "CHEQUE", "CASH", "INTEREST", "CHARGE"}:
                        rows.append(f"{last_date_str} {s}")
                        continue

                # Otherwise append to previous row as continuation
                if rows:
                    rows[-1] = (rows[-1] + " " + s).strip()

    return rows


def _parse_row(
    raw_row: str,
    period_start: Optional[_dt.date],
    period_end: Optional[_dt.date],
    prev_balance: Optional[float],
) -> Tuple[Optional[Dict], Optional[float]]:
    raw_row = re.sub(r"\s+", " ", (raw_row or "").strip())
    if not raw_row:
        return None, prev_balance

    m = re.match(r"^(?P<d>\d{2})\s+(?P<m>[A-Za-z]{3})\s*(?P<y>\d{2})?\s+(?P<rest>.*)$", raw_row)
    if not m:
        return None, prev_balance

    day = int(m.group("d"))
    mon_abbr = m.group("m")
    yy = int(m.group("y")) if m.group("y") else None
    rest = m.group("rest").strip()

    # Identify floats in the row
    floats = _FLOAT_RE.findall(rest)
    if not floats:
        return None, prev_balance

    up_rest = rest.upper()

    # Opening/Closing summary lines inside table
    if "STATEMENT OPENING BALANCE" in up_rest:
        bal = _parse_money(floats[-1])
        return None, (bal if bal is not None else prev_balance)

    if "STATEMENT CLOSING BALANCE" in up_rest:
        return None, prev_balance

    # Balance is the last float
    balance = _parse_money(floats[-1])

    # Determine amount using balance delta if possible (TSB often shows only one of Money Out / Money In)
    amount: Optional[float] = None

    if len(floats) >= 3:
        # If 3 floats present (rare in normal rows), interpret as Out, In, Balance
        out_val = _parse_money(floats[-3])
        in_val = _parse_money(floats[-2])
        if in_val and abs(in_val) > 0:
            amount = in_val
        elif out_val and abs(out_val) > 0:
            amount = -out_val
        else:
            # fallback to delta if possible
            if balance is not None and prev_balance is not None:
                amount = round(balance - prev_balance, 2)

        rest_wo_nums = _remove_last_n_floats(rest, 3)

    else:
        # Standard case: 2 floats -> (unknown: out or in) + balance
        cand = _parse_money(floats[-2])
        if cand is None:
            return None, prev_balance

        if balance is not None and prev_balance is not None:
            delta = round(balance - prev_balance, 2)

            # If delta matches +/- cand, use that
            if abs(delta - cand) <= 0.01:
                amount = cand
            elif abs(delta + cand) <= 0.01:
                amount = -cand
            else:
                # Trust delta if it doesn't match cleanly (layout oddities)
                amount = delta
        else:
            # No previous balance: heuristic
            if ("DIRECT CREDIT" in up_rest) or ("CREDIT" in up_rest):
                amount = cand
            else:
                amount = -cand

        rest_wo_nums = _remove_last_n_floats(rest, 2)

    # Extract type + description
    tx_type_raw, desc_raw = _split_type_details(rest_wo_nums)

    tx_date = _infer_date(day, mon_abbr, yy, period_start, period_end)

    # Update prev_balance
    new_prev = balance if balance is not None else prev_balance

    # Apply global normalisation rules
    tx_type, desc = _normalize_type_and_desc(tx_type_raw, desc_raw)

    tx = {
        "Date": tx_date,
        "Transaction Type": tx_type,
        "Description": desc,
        "Amount": float(amount) if amount is not None else None,
        "Balance": float(balance) if balance is not None else None,
    }
    return tx, new_prev


# -----------------------------
# Required public functions
# -----------------------------

def extract_transactions(pdf_path: str) -> List[Dict]:
    """
    Return list of dicts with keys:
      Date (datetime.date)
      Transaction Type (string)
      Description (string)
      Amount (float; credits +, debits -)
      Balance (float or None)
    """
    with pdfplumber.open(pdf_path) as pdf:
        first_text = _extract_text(pdf.pages[0])

    period_start, period_end = _parse_statement_period(first_text)
    start_bal, _ = _parse_statement_balances(first_text)

    rows = _extract_rows(pdf_path)

    txs: List[Dict] = []
    prev_balance: Optional[float] = start_bal

    for r in rows:
        tx, prev_balance = _parse_row(r, period_start, period_end, prev_balance)
        if tx:
            txs.append(tx)

    return txs


def extract_statement_balances(pdf_path: str) -> Dict:
    """
    Return:
      {"start_balance": float|None, "end_balance": float|None}
    """
    with pdfplumber.open(pdf_path) as pdf:
        first_text = _extract_text(pdf.pages[0])

    start_bal, end_bal = _parse_statement_balances(first_text)
    return {"start_balance": start_bal, "end_balance": end_bal}


def extract_account_holder_name(pdf_path: str) -> str:
    """
    Return best account holder name from the statement (avoid generic headings).
    For the provided TSB statements, the name appears as the first line on page 1.
    """
    with pdfplumber.open(pdf_path) as pdf:
        first_text = _extract_text(pdf.pages[0])

    lines = [ln.strip() for ln in (first_text or "").splitlines() if ln.strip()]

    # Prefer first line that looks like a personal/business name (no digits, not bank headings)
    for ln in lines[:15]:
        if any(k in ln.lower() for k in ["tsb", "sort code", "account number", "spend & save", "statement number", "effective from"]):
            continue
        if re.search(r"\d", ln):  # address lines often contain digits
            continue
        if len(ln) >= 3:
            return ln

    # fallback
    return lines[0] if lines else "Unknown"

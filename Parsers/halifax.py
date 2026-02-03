# halifax.py
"""Halifax (UK) bank statement parser (text-based PDFs only; NO OCR).

This parser is designed for the Halifax statement layout seen in the provided sample PDFs.

Exports exactly these functions:
- extract_transactions(pdf_path) -> list[dict]
- extract_statement_balances(pdf_path) -> dict {"start_balance": float|None, "end_balance": float|None}
- extract_account_holder_name(pdf_path) -> str

Transaction dict keys (exact):
- Date (datetime.date)
- Transaction Type (string)
- Description (string)
- Amount (float; credits positive, debits negative)
- Balance (float or None)

Notes:
- Uses coordinate-based parsing (pdfplumber.extract_words) with column boundaries inferred from PDF line geometry.
- Handles multi-page tables.
- Handles multi-line descriptions by appending “description-only” continuation lines.
- Handles “date omitted on subsequent lines” by reusing the last seen date when a row has amounts/type but no date.
- Normalises Transaction Type & Description using the provided GLOBAL TRANSACTION TYPE RULES.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import pdfplumber


# -----------------------------
# Helpers
# -----------------------------

_MONTH_NAME_TO_NUM = {
    "JANUARY": 1,
    "FEBRUARY": 2,
    "MARCH": 3,
    "APRIL": 4,
    "MAY": 5,
    "JUNE": 6,
    "JULY": 7,
    "AUGUST": 8,
    "SEPTEMBER": 9,
    "OCTOBER": 10,
    "NOVEMBER": 11,
    "DECEMBER": 12,
}

_MONTH_ABBR_TO_NUM = {
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


def _parse_money(value: Optional[str]) -> Optional[float]:
    """Parse UK money formats like '£1,234.56', '-£4.80', '(£12.34)', '1,234.56'."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None

    neg = False

    # Parentheses indicate negative
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()

    s = s.replace("£", "").replace(",", "").replace(" ", "")

    # Normal minus / unicode minus
    if s.startswith("-"):
        neg = True
        s = s[1:]
    s = s.replace("−", "")

    # If still not a clean float, try to find the first money-ish token
    try:
        v = float(s)
        return -v if neg else v
    except Exception:
        m = re.search(r"-?£?\(?\d[\d,]*\.\d{2}\)?", str(value))
        if not m:
            return None
        return _parse_money(m.group(0))


def _parse_period_from_text(text: str) -> Tuple[Optional[date], Optional[date]]:
    """Parse 'CURRENT ACCOUNT 01 April 2024 to 30 April 2024'."""
    m = re.search(
        r"CURRENT\s+ACCOUNT\s+(\d{2}\s+[A-Za-z]+\s+\d{4})\s+to\s+(\d{2}\s+[A-Za-z]+\s+\d{4})",
        text or "",
    )
    if not m:
        return None, None

    def _parse_full_date(ds: str) -> Optional[date]:
        parts = ds.split()
        if len(parts) != 3:
            return None
        d = int(parts[0])
        mon = _MONTH_NAME_TO_NUM.get(parts[1].upper())
        y = int(parts[2])
        if not mon:
            return None
        return date(y, mon, d)

    return _parse_full_date(m.group(1)), _parse_full_date(m.group(2))


def _extract_period(pdf_path: str) -> Tuple[Optional[date], Optional[date]]:
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:2]:
            txt = page.extract_text() or ""
            start, end = _parse_period_from_text(txt)
            if start and end:
                return start, end
    return None, None


def _extract_summary_balances_from_text(text: str) -> Tuple[Optional[float], Optional[float]]:
    """Find the first two 'Balance on <date> <amount>' occurrences."""
    # Most common: 'Balance on 01 May 2024 -£4.80'
    matches = re.findall(
        r"Balance on\s+(\d{2}\s+[A-Za-z]+\s+\d{4})\s+(-?£[\d,]+\.\d{2}|\(?-?£[\d,]+\.\d{2}\)?)",
        text or "",
    )

    # Fallback if £ is missing in extracted text
    if not matches:
        matches = re.findall(
            r"Balance on\s+(\d{2}\s+[A-Za-z]+\s+\d{4})\s+(-?[\d,]+\.\d{2})",
            text or "",
        )

    amounts: List[float] = []
    for _dstr, astr in matches:
        v = _parse_money(astr)
        if v is not None:
            amounts.append(float(v))

    if len(amounts) >= 2:
        return amounts[0], amounts[1]
    if len(amounts) == 1:
        return amounts[0], None
    return None, None


def _get_column_boundaries(page) -> List[float]:
    """Infer column boundaries from vertical lines; fallback to known defaults."""
    defaults = [115.66, 264.26, 303.37, 381.58, 459.79]
    try:
        vlines = [
            l
            for l in page.lines
            if abs(l.get("x0", 0) - l.get("x1", 0)) < 0.5
            and (l.get("height", 0) or abs(l.get("y1", 0) - l.get("y0", 0))) > 10
        ]
        xs = sorted({round(l["x0"], 2) for l in vlines})
        xs = [x for x in xs if 80 < x < 520]
        if len(xs) >= 5:
            return xs[:5]
    except Exception:
        pass
    return defaults


def _group_words_into_lines(words: List[dict], tol: float = 2.0) -> List[List[dict]]:
    words = sorted(words, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
    lines: List[List[dict]] = []
    current: List[dict] = []
    current_top: Optional[float] = None

    for w in words:
        t = float(w.get("top", 0))
        if current_top is None or abs(t - current_top) <= tol:
            current.append(w)
            current_top = t if current_top is None else (current_top * 0.7 + t * 0.3)
        else:
            lines.append(current)
            current = [w]
            current_top = t

    if current:
        lines.append(current)

    return lines


_LABEL_TOKENS = {
    "Date",
    "Description",
    "Type",
    "Money",
    "In",
    "Out",
    "Balance",
    "(£)",
    "Column",
    "blank.",
    "CDolumn",
    "CTolumn",
    "D0ate",
    "DDescription",
    "DPescription",
    "DIescriptionFL",
    "Moneyb",
    "Ilna",
    "n(k£.)",
    "Obulat",
    "n(£k).",
    ".",
}


def _clean_tokens(tokens: List[str]) -> List[str]:
    cleaned: List[str] = []
    for tok in tokens:
        t = (tok or "").strip()
        if not t:
            continue
        if t in {".", "·", "•"}:
            continue
        if t in _LABEL_TOKENS:
            continue
        cleaned.append(t)
    return cleaned


def _parse_date_from_tokens(
    date_tokens: List[str],
    last_date: Optional[date],
    period: Tuple[Optional[date], Optional[date]],
) -> Optional[date]:
    toks = _clean_tokens(date_tokens)

    day: Optional[int] = None
    mon: Optional[int] = None
    yr: Optional[int] = None

    for i in range(len(toks)):
        if re.fullmatch(r"\d{1,2}", toks[i]):
            if i + 1 < len(toks) and toks[i + 1][:3].upper() in _MONTH_ABBR_TO_NUM:
                day = int(toks[i])
                mon = _MONTH_ABBR_TO_NUM[toks[i + 1][:3].upper()]
                if i + 2 < len(toks) and re.fullmatch(r"\d{2,4}", toks[i + 2]):
                    y = int(toks[i + 2])
                    yr = (2000 + y) if y < 100 else y
                break

    if day is None or mon is None:
        return None

    # If year present, use it
    if yr is not None:
        try:
            return date(yr, mon, day)
        except Exception:
            return None

    # Otherwise infer year from statement period (or last_date)
    start, end = period
    candidate_years: List[int] = []
    if start:
        candidate_years.append(start.year)
    if end and end.year not in candidate_years:
        candidate_years.append(end.year)
    if not candidate_years and last_date:
        candidate_years.append(last_date.year)
    if not candidate_years:
        candidate_years.append(datetime.utcnow().year)

    candidates: List[date] = []
    for y in candidate_years:
        try:
            candidates.append(date(y, mon, day))
        except Exception:
            continue

    if not candidates:
        return None

    if start and end:
        # Prefer a candidate within the statement coverage period
        for d in candidates:
            if start <= d <= end:
                return d
        # Otherwise pick closest to the range boundaries
        candidates.sort(key=lambda d: min(abs((d - start).days), abs((d - end).days)))
        return candidates[0]

    return candidates[0]


def _find_number_in_tokens(tokens: List[str]) -> Optional[float]:
    for t in tokens:
        m = re.search(r"-?£?\(?\d[\d,]*\.\d{2}\)?", t or "")
        if m:
            return _parse_money(m.group(0))
    return None


def _parse_type_code(type_tokens: List[str]) -> Optional[str]:
    toks = _clean_tokens(type_tokens)
    for t in toks:
        if re.fullmatch(r"[A-Z]{2,4}", t):
            return t
    return None


def _join_description(desc_tokens: List[str]) -> str:
    toks = _clean_tokens(desc_tokens)
    return " ".join(toks).strip()


def _parse_legend_mapping_from_text(text: str) -> Dict[str, str]:
    """Parse the 'Transaction types' legend into a dict {"CHG": "Charge", ...}."""
    if not text:
        return {}
    if "Transaction types" in text:
        text = text.split("Transaction types", 1)[1]

    # Remove filler dots and compress whitespace
    text = re.sub(r"\.+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    tokens = text.split(" ")
    mapping: Dict[str, str] = {}

    current_code: Optional[str] = None
    current_phrase: List[str] = []

    def flush() -> None:
        nonlocal current_code, current_phrase
        if current_code and current_phrase:
            phrase = " ".join(current_phrase).strip()
            if phrase and not phrase.lower().startswith("blank"):
                mapping[current_code] = phrase
        current_code = None
        current_phrase = []

    for tok in tokens:
        if re.fullmatch(r"[A-Z]{2,4}", tok):
            if current_code is not None:
                flush()
            current_code = tok
            current_phrase = []
        else:
            if current_code is not None:
                if tok.lower() == "blank":
                    continue
                # Stop when footer starts
                if tok.lower() == "if":
                    flush()
                    break
                current_phrase.append(tok)

    flush()
    return mapping


_IGNORE_SUBSTRINGS = [
    "continued on next page",
    "if you think something is incorrect",
    "halifax is a division",
    "bank of scotland plc",
    "prudential regulation authority",
    "financial conduct authority",
    "registered office",
    "registration number",
    "document requested by",
]


def _should_ignore_line(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True

    low = t.lower()

    if low.startswith("logo"):
        return True

    if re.search(r"page\s+\d+\s+of\s+\d+", low):
        return True

    if "column" in low or "cdolumn" in low or "ctolumn" in low:
        return True

    if low.startswith("(continued"):
        return True

    for s in _IGNORE_SUBSTRINGS:
        if s in low:
            return True

    return False


def _normalise_type_and_description(tx_type: str, desc: str) -> Tuple[str, str]:
    tx_type = (tx_type or "").strip()
    desc = (desc or "").strip()

    # Returned Direct Debit rule
    if re.search(r"\breturned\s+direct\s+debit\b", desc, flags=re.I) or re.search(
        r"\breturned\s+direct\s+debit\b", tx_type, flags=re.I
    ):
        tx_type = "Direct Debit"
        if not desc.lower().startswith("returned direct debit"):
            desc = "Returned Direct Debit " + desc
        else:
            # Ensure capitalised prefix
            if not desc.startswith("Returned Direct Debit"):
                desc = "Returned Direct Debit" + desc[len("returned direct debit") :]
        return tx_type, desc.strip()

    # Card payment overrides
    if (
        re.search(r"apple\s*pay", desc, flags=re.I)
        or re.search(r"clearpay", desc, flags=re.I)
        or re.search(r"contactless", desc, flags=re.I)
        or re.search(r"\bGB$", desc, flags=re.I)
    ):
        tx_type = "Card Payment"
    else:
        tx_type = tx_type.title()

    # Remove type prefix from description (except Returned Direct Debit handled above)
    if tx_type:
        if re.match(rf"^{re.escape(tx_type)}(\s+|[-:])", desc, flags=re.I):
            desc = re.sub(rf"^{re.escape(tx_type)}(\s+|[-:])+", "", desc, flags=re.I).strip()

    return tx_type, desc


# -----------------------------
# Required public functions
# -----------------------------


def extract_transactions(pdf_path: str) -> List[dict]:
    """Extract transactions from a Halifax statement PDF."""

    period = _extract_period(pdf_path)

    with pdfplumber.open(pdf_path) as pdf:
        # Legend mapping (code -> phrase) is usually on the last page
        legend_map: Dict[str, str] = {}
        for p in reversed(pdf.pages):
            t = p.extract_text() or ""
            if "Transaction types" in t:
                legend_map = _parse_legend_mapping_from_text(t)
                break

        transactions: List[dict] = []
        last_seen_date: Optional[date] = None
        prev_balance: Optional[float] = None
        current_tx: Optional[dict] = None

        for page in pdf.pages:
            page_text = page.extract_text() or ""
            # Skip the legend-only page
            if "Transaction types" in page_text and "Your Transactions" not in page_text:
                continue

            col_xs = _get_column_boundaries(page)
            x_date_end, x_desc_end, x_type_end, x_in_end, x_out_end = col_xs

            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            lines = _group_words_into_lines(words, tol=2.0)

            for line in lines:
                all_text = " ".join(w.get("text", "") for w in line).strip()
                if _should_ignore_line(all_text):
                    continue
                if "Your Transactions" in all_text:
                    continue
                if "Transaction types" in all_text:
                    break

                # Split line words into columns by x-position
                date_tokens = [w["text"] for w in line if w.get("x0", 0) < x_date_end]
                desc_tokens = [w["text"] for w in line if x_date_end <= w.get("x0", 0) < x_desc_end]
                type_tokens = [w["text"] for w in line if x_desc_end <= w.get("x0", 0) < x_type_end]
                in_tokens = [w["text"] for w in line if x_type_end <= w.get("x0", 0) < x_in_end]
                out_tokens = [w["text"] for w in line if x_in_end <= w.get("x0", 0) < x_out_end]
                bal_tokens = [w["text"] for w in line if w.get("x0", 0) >= x_out_end]

                tx_date = _parse_date_from_tokens(date_tokens, last_seen_date, period)
                type_code = _parse_type_code(type_tokens)
                money_in = _find_number_in_tokens(in_tokens)
                money_out = _find_number_in_tokens(out_tokens)
                balance = _find_number_in_tokens(bal_tokens)
                desc = _join_description(desc_tokens)

                has_amt = (money_in is not None) or (money_out is not None)
                has_any = has_amt or (balance is not None) or (type_code is not None)
                has_desc = bool(desc)

                # Candidate if a date is present, or date omitted but looks like a row (type/amount/balance + desc)
                is_candidate = (tx_date is not None) or (last_seen_date is not None and has_any and has_desc)

                if not is_candidate:
                    continue

                if tx_date is None:
                    tx_date = last_seen_date

                # Continuation line: description-only (no type/amount/balance)
                if type_code is None and (not has_amt) and (balance is None):
                    if current_tx and desc:
                        current_tx["Description"] = (current_tx.get("Description", "") + " " + desc).strip()
                    continue

                # New transaction
                last_seen_date = tx_date

                amount: Optional[float] = None

                if money_in is not None and money_out is None:
                    amount = abs(float(money_in))
                elif money_out is not None and money_in is None:
                    amount = -abs(float(money_out))
                elif money_in is not None and money_out is not None:
                    # Prefer balance delta if available
                    if balance is not None and prev_balance is not None:
                        amount = round(float(balance) - float(prev_balance), 2)
                    else:
                        amount = abs(float(money_in)) if abs(float(money_in)) >= abs(float(money_out)) else -abs(float(money_out))
                else:
                    # No explicit Money In/Out values; derive from running balance if possible
                    if balance is not None and prev_balance is not None:
                        amount = round(float(balance) - float(prev_balance), 2)

                if amount is None:
                    amount = 0.0

                # If balances exist, let the delta win (more reliable when extraction is messy)
                if balance is not None and prev_balance is not None:
                    delta = round(float(balance) - float(prev_balance), 2)
                    if abs(delta - float(amount)) > 0.01:
                        amount = delta

                tx_type = (legend_map.get(type_code or "", type_code or "") or "").strip()

                tx = {
                    "Date": tx_date,
                    "Transaction Type": tx_type,
                    "Description": desc,
                    "Amount": float(amount),
                    "Balance": float(balance) if balance is not None else None,
                }
                transactions.append(tx)
                current_tx = tx

                if balance is not None:
                    prev_balance = float(balance)

        # Apply GLOBAL TRANSACTION TYPE RULES
        for tx in transactions:
            tx["Transaction Type"], tx["Description"] = _normalise_type_and_description(
                tx.get("Transaction Type", ""), tx.get("Description", "")
            )

        return transactions


def extract_statement_balances(pdf_path: str) -> dict:
    """Return {'start_balance': float|None, 'end_balance': float|None}.

    Halifax summary lines are typically:
      'Balance on 01 <Month> <Year> ...'
      'Balance on <end date> ...'

    If the summary start balance doesn't reconcile with the transaction table,
    we prefer an implied opening balance derived from:
      opening = first_running_balance - first_amount

    Likewise, if the summary end balance doesn't match the last running balance,
    we prefer the last running balance.
    """

    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""

    start_summary, end_summary = _extract_summary_balances_from_text(first_page_text)

    txs = extract_transactions(pdf_path)

    first_with_balance = next((t for t in txs if t.get("Balance") is not None), None)
    last_with_balance = next((t for t in reversed(txs) if t.get("Balance") is not None), None)

    implied_start: Optional[float] = None
    if first_with_balance is not None:
        implied_start = round(float(first_with_balance["Balance"]) - float(first_with_balance["Amount"]), 2)

    implied_end: Optional[float] = None
    if last_with_balance is not None:
        implied_end = round(float(last_with_balance["Balance"]), 2)

    start_final = start_summary
    end_final = end_summary

    if start_final is None and implied_start is not None:
        start_final = implied_start
    elif start_final is not None and implied_start is not None:
        if abs(round(float(start_final) - float(implied_start), 2)) > 0.01:
            start_final = implied_start

    if end_final is None and implied_end is not None:
        end_final = implied_end
    elif end_final is not None and implied_end is not None:
        if abs(round(float(end_final) - float(implied_end), 2)) > 0.01:
            end_final = implied_end

    return {"start_balance": start_final, "end_balance": end_final}


def extract_account_holder_name(pdf_path: str) -> str:
    """Return the best client name found on the statement (avoid generic headings)."""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    for i, line in enumerate(lines):
        if line.lower().startswith("document requested by"):
            for j in range(i + 1, min(i + 6, len(lines))):
                cand = lines[j].strip()
                if not cand:
                    continue
                if cand.lower() in {"your account", "sort code", "account number"}:
                    continue
                return cand

    return ""

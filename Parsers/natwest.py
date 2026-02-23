# Version: natwest.py
import os
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
_TXN_START_EXPORT_RE = re.compile(
    r"^\s*(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\s+(?:(?P<year>\d{4})\s+)?(?P<type>[A-Z/]{2,4})\b"
)

# Start of a transaction row (NatWest statement table)
# Example: "12 APR Card Transaction ... 24.32 333.36"
_TXN_START_TABLE_RE = re.compile(
    r"^\s*(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\b\s*(?P<rest>.*)$"
)
_DATE_PREFIX_RE = re.compile(r"^\s*(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\b")
_TXN_KEYWORD_RE = re.compile(
    r"^\s*(Card\s+Transaction|Direct\s+Debit|OnLine\s+Transaction|Automated\s+Credit|Automated\s+Debit|Standing\s+Order|Cash\s+Withdrawal|Charges)\b",
    re.IGNORECASE,
)
_BROUGHT_FWD_RE = re.compile(r"\bBROUGHT\s+FORWARD\b", re.IGNORECASE)
_SUMMARY_GUARD_RE = re.compile(r"\b(Period\s+Covered|Statement\s+Date)\b", re.IGNORECASE)
_TABLE_EMBED_TXN_RE = re.compile(r"(?<!\d)(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\b", re.IGNORECASE)
_CHARGES_LINE_RE = re.compile(
    r"^\s*(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\s+Charges\b(?P<rest>.*)$",
    re.IGNORECASE | re.MULTILINE,
)
_HEADER_PREAMBLE_CUE_RE = re.compile(
    r"(Account\s+Name\s+Account\s+No\s+Sort\s+Code\s+Page\s+No|Current\s+Account\s+Summary|Welcome\s+to\s+your\s+NatWest\s+Statement|\bBIC\b|\bIBAN\b|Registered\s+Office:)",
    re.IGNORECASE,
)
_TRANSACTION_KEYWORD_FINDER_RE = re.compile(
    r"(Returned\s+Direct\s+Debit|Card\s+Transaction|Direct\s+Debit|OnLine\s+Transaction|Automated\s+Credit|Automated\s+Debit|Standing\s+Order|Cash\s+Withdrawal|Charges)",
    re.IGNORECASE,
)
_CARD_TRANSACTION_PREFIX_RE = re.compile(
    r"^Card\s+Transaction\s+\d{3,4}\s+\d{2}[A-Z]{3}\d{2}\s+(?:(?:CD|C|D)\s+)?",
    re.IGNORECASE,
)
_GENERIC_KEYWORD_PREFIX_RE = re.compile(
    r"^(Direct\s+Debit|OnLine\s+Transaction|Automated\s+Credit|Automated\s+Debit|Standing\s+Order|Cash\s+Withdrawal|Charges)\s+",
    re.IGNORECASE,
)

# Currency amounts (NatWest export/table can include or omit £)
_MONEY_RE = re.compile(r"(?:£\s*)?\(?-?[\d,]+\.\d{2}\)?(?!\s*%)")

# Header/footer noise lines to ignore
_IGNORE_LINE_RE_LIST = [
    re.compile(r"^\s*Page\s+\d+\s+of\s+\d+\s*$", re.IGNORECASE),
    re.compile(r"^\s*Date\s+Type\s+Description\s+Paid\s+in\s+Paid\s+out\s+Balance\s*$", re.IGNORECASE),
    re.compile(r"^\s*Date\s+Description\s+Paid\s+In\(£\)\s+Withdrawn\(£\)\s+Balance\(£\)\s*$", re.IGNORECASE),
    re.compile(r"^\s*Transactions\s*$", re.IGNORECASE),
    re.compile(r"^\s*BROUGHT\s+FORWARD\b.*$", re.IGNORECASE),
    re.compile(r"^\s*Debit\s+interest\s+details\b.*$", re.IGNORECASE),
    re.compile(r"^\s*Overdraft\s+Limit\b.*$", re.IGNORECASE),
    re.compile(r"^\s*Overdraft\s+Rate\b.*$", re.IGNORECASE),
    re.compile(r"^\s*(?:UNARRANGED|ARRANGED)\b.*$", re.IGNORECASE),
    re.compile(r"^\s*©\s*National\s+Westminster\s+Bank\b", re.IGNORECASE),
    re.compile(r"^\s*National\s+Westminster\s+Bank\b", re.IGNORECASE),
    re.compile(r"^\s*Authorised\s+by\s+the\s+Prudential\b", re.IGNORECASE),
    re.compile(r"^\s*.*\bPeriod\s+Covered\s+\d{2}\s+[A-Za-z]{3}\s+\d{4}\s+to\s+\d{2}\s+[A-Za-z]{3}\s+\d{4}\b.*$", re.IGNORECASE),
]

_PERIOD_RE = re.compile(
    r"Showing:\s*(?P<d1>\d{2})\s+(?P<m1>[A-Za-z]{3})\s+(?P<y1>\d{4})\s+to\s+(?P<d2>\d{2})\s+(?P<m2>[A-Za-z]{3})\s+(?P<y2>\d{4})",
    re.IGNORECASE
)
_TABLE_PERIOD_RE = re.compile(
    r"\bPeriod\s+Covered\s+(?P<d1>\d{2})\s+(?P<m1>[A-Za-z]{3})\s+(?P<y1>\d{4})\s+to\s+(?P<d2>\d{2})\s+(?P<m2>[A-Za-z]{3})\s+(?P<y2>\d{4})",
    re.IGNORECASE,
)

_ACCOUNT_NAME_RE = re.compile(r"^\s*Account\s+name:\s*(?P<name>.+?)\s*$", re.IGNORECASE)
_OPENING_BALANCE_RE = re.compile(r"\bBROUGHT\s+FORWARD\b.*?(?:£\s*)?(-?[\d,]+\.\d{2})", re.IGNORECASE)
_CLOSING_BALANCE_RE = re.compile(r"\b(CARRIED\s+FORWARD|BALANCE\s+AT)\b.*?(?:£\s*)?(-?[\d,]+\.\d{2})", re.IGNORECASE)
_PREV_BAL_RE = re.compile(r"\bPrevious\s+Balance\s+(?:£\s*)?(-?[\d,]+\.\d{2})", re.IGNORECASE)
_NEW_BAL_RE = re.compile(r"\bNew\s+Balance\s+(?:£\s*)?(-?[\d,]+\.\d{2})", re.IGNORECASE)
_ACCOUNT_NAME_INLINE_RE = re.compile(r"\b(Account\s+name|Name)\b\s*:\s*([A-Z][A-Z\s'\-]{3,})", re.IGNORECASE)
_ACCOUNT_NAME_NEXT_LINE_LABEL_RE = re.compile(r"^\s*(Account\s+name|Name)\s*:?\s*$", re.IGNORECASE)


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
    if t.startswith("-"):
        neg = True
        t = t[1:].strip()
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
    if _HEADER_PREAMBLE_CUE_RE.search(s):
        m = _TRANSACTION_KEYWORD_FINDER_RE.search(s)
        if m:
            s = s[m.start():]
    if s.lower().startswith("returned direct debit"):
        s = re.sub(r"\s+", " ", s).strip()
        s = re.sub(r"\s+,", ",", s)
        s = re.sub(r",\s*", ", ", s)
        s = re.sub(r"\s{2,}", " ", s).strip()
        return s
    s = _CARD_TRANSACTION_PREFIX_RE.sub("", s, count=1)
    s = _GENERIC_KEYWORD_PREFIX_RE.sub("", s, count=1)
    # Remove spaces before commas
    s = re.sub(r"\s+,", ",", s)
    # Remove duplicated commas spacing " , " -> ", "
    s = re.sub(r",\s*", ", ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    s = re.sub(r"\s+CD\s*\d{4}\b\s*$", "", s, flags=re.IGNORECASE).strip()
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


def _extract_period_dates(all_text: str) -> (Optional[_dt.date], Optional[_dt.date]):
    if not all_text:
        return None, None
    m = _PERIOD_RE.search(all_text)
    if not m:
        return None, None
    try:
        d1 = int(m.group("d1"))
        m1 = _MONTHS.get(m.group("m1").lower())
        y1 = int(m.group("y1"))
        d2 = int(m.group("d2"))
        m2 = _MONTHS.get(m.group("m2").lower())
        y2 = int(m.group("y2"))
        if not m1 or not m2:
            return None, None
        return _dt.date(y1, m1, d1), _dt.date(y2, m2, d2)
    except Exception:
        return None, None


def _split_embedded_table_lines(lines: List[str]) -> List[str]:
    split_lines: List[str] = []

    for raw_line in lines:
        line = (raw_line or "").rstrip("\n")
        if "period covered" in line.lower():
            split_lines.append(line)
            continue
        valid_matches = []

        for m in _TABLE_EMBED_TXN_RE.finditer(line):
            mon = (m.group("mon") or "").lower()
            if mon[:3] not in _MONTHS:
                continue
            start = m.start()
            if start > 0 and line[start - 1] == "/":
                continue
            remainder = line[m.end():]
            if not re.search(r"\s+[A-Za-z]", remainder):
                continue
            valid_matches.append(m)

        if not valid_matches:
            split_lines.append(line)
            continue

        if len(valid_matches) == 1 and valid_matches[0].start() == 0:
            split_lines.append(line)
            continue

        first_start = valid_matches[0].start()
        if first_start > 0:
            prefix = line[:first_start].strip()
            if prefix:
                split_lines.append(prefix)

        for i, m in enumerate(valid_matches):
            seg_start = m.start()
            seg_end = valid_matches[i + 1].start() if (i + 1) < len(valid_matches) else len(line)
            segment = line[seg_start:seg_end].strip()
            if segment:
                split_lines.append(segment)

    return split_lines


def _extract_table_period_dates(all_text: str) -> (Optional[_dt.date], Optional[_dt.date]):
    if not all_text:
        return None, None
    m = _TABLE_PERIOD_RE.search(all_text)
    if not m:
        return None, None
    try:
        d1 = int(m.group("d1"))
        m1 = _MONTHS.get(m.group("m1").lower())
        y1 = int(m.group("y1"))
        d2 = int(m.group("d2"))
        m2 = _MONTHS.get(m.group("m2").lower())
        y2 = int(m.group("y2"))
        if not m1 or not m2:
            return None, None
        return _dt.date(y1, m1, d1), _dt.date(y2, m2, d2)
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
        return _dt.date(y1, m1, d1), _dt.date(y2, m2, d2)
    except Exception:
        return None, None


def extract_statement_period(pdf_path: str):
    """Public wrapper to extract the statement coverage period (start_date, end_date)."""
    try:
        all_text_chunks = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                all_text_chunks.append(page.extract_text() or "")
        all_text = "\n".join(all_text_chunks)
        has_table_header = "date description paid in" in all_text.lower()
        if has_table_header:
            start, end = _extract_table_period_dates(all_text)
            if start or end:
                return start, end
        start, end = _extract_period_dates(all_text)
        if start or end:
            return start, end
        return _parse_period_from_filename(pdf_path)
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

    has_table_header = "date description paid in" in all_text.lower()
    period_start_date = None
    period_end_date = None
    if has_table_header:
        period_start_date, period_end_date = _extract_table_period_dates(all_text)
    if period_start_date and period_end_date:
        period_start_year, period_end_year = period_start_date.year, period_end_date.year
    else:
        period_start_year, period_end_year = _extract_period_years(all_text)

    if has_table_header:
        parse_lines = _split_embedded_table_lines(all_text.splitlines())
        pending_rows: List[Dict] = []
        current_date = None
        current_desc_lines: List[str] = []
        current_raw_type = ""
        last_seen_mon = None
        current_year = period_end_year if period_end_year is not None else _dt.date.today().year

        def _clamp_table_year(year_value: int) -> int:
            if period_start_year is not None and year_value < period_start_year:
                return period_start_year
            if period_end_year is not None and year_value > period_end_year:
                return period_end_year
            return year_value

        def _remove_terminal_money_tokens(raw_text: str, tokens: List[str]) -> str:
            cleaned = raw_text.rstrip()
            for token in reversed(tokens[-2:]):
                cleaned = re.sub(rf"\s*{re.escape(token)}\s*$", "", cleaned)
            return cleaned.strip()

        for raw_line in parse_lines:
            line = re.sub(r"\s+", " ", (raw_line or "").strip())
            if not line:
                continue
            if _SUMMARY_GUARD_RE.search(line):
                continue
            if _is_ignorable_line(line):
                continue
            if _BROUGHT_FWD_RE.search(line):
                continue

            had_explicit_date = False
            date_match = _DATE_PREFIX_RE.match(line)
            if date_match:
                had_explicit_date = True
                day = int(date_match.group("day"))
                mon = (date_match.group("mon") or "").lower()
                mon_num = _MONTHS.get(mon[:3])
                if mon_num is not None:
                    inferred_year = current_year
                    if last_seen_mon is not None:
                        if mon_num < last_seen_mon:
                            inferred_year += 1
                        elif mon_num > last_seen_mon and last_seen_mon <= 2 and mon_num >= 11:
                            inferred_year -= 1
                    inferred_year = _clamp_table_year(inferred_year)
                    try:
                        current_date = _dt.date(inferred_year, mon_num, day)
                        current_year = inferred_year
                        last_seen_mon = mon_num
                    except Exception:
                        current_date = None
                else:
                    current_date = None
                line = line[date_match.end():].strip()

            monies = [m.strip() for m in _MONEY_RE.findall(line) if m and m.strip()]
            candidate_balance = _parse_money(monies[-1]) if monies else None
            keyword_match = _TXN_KEYWORD_RE.match(line)
            if keyword_match:
                current_raw_type = keyword_match.group(1).strip()
            is_row_terminator = (
                len(monies) >= 2
                and current_date is not None
                and candidate_balance is not None
                and (
                    had_explicit_date
                    or keyword_match is not None
                    or current_raw_type
                    or current_desc_lines
                )
            )

            if is_row_terminator:
                line_desc = _remove_terminal_money_tokens(line, monies)
                desc_parts = [part for part in current_desc_lines if part]
                if line_desc:
                    desc_parts.append(line_desc)
                description = " ".join(desc_parts).strip()
                raw_type = current_raw_type
                if keyword_match:
                    raw_type = keyword_match.group(1).strip()
                pending_rows.append(
                    {
                        "Date": current_date,
                        "Transaction Type": raw_type,
                        "Description": description,
                        "Amount": 0.0,
                        "Balance": float(candidate_balance),
                        "_raw_type": raw_type,
                    }
                )
                current_desc_lines = []
                current_raw_type = ""
            elif line:
                current_desc_lines.append(line)

        if not pending_rows:
            return []

        previous_balance = None
        prev_match = _PREV_BAL_RE.search(all_text)
        if prev_match:
            previous_balance = _parse_money(prev_match.group(1))
        prev_running = previous_balance
        for i, row in enumerate(pending_rows):
            if isinstance(prev_running, (int, float)):
                row["Amount"] = round(float(row["Balance"]) - float(prev_running), 2)
            elif i > 0:
                row["Amount"] = round(float(row["Balance"]) - float(pending_rows[i - 1]["Balance"]), 2)
            prev_running = float(row["Balance"])

        cleaned = []
        for txn in pending_rows:
            txn.pop("_raw_type", None)
            txn["Description"] = _clean_description(txn.get("Description") or "")
            txn = _apply_global_transaction_type_rules(txn)
            cleaned.append(txn)
        return cleaned

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

        if has_table_header and re.search(r"\bBROUGHT\s+FORWARD\b", block_text, re.IGNORECASE):
            return

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

    parse_lines = all_text.splitlines()
    if has_table_header:
        parse_lines = _split_embedded_table_lines(parse_lines)

    # Parse line-by-line, building blocks
    for raw_line in parse_lines:
        line = raw_line.rstrip("\n")
        if _is_ignorable_line(line):
            continue

        m = _TXN_START_EXPORT_RE.match(line)
        raw_type = ""
        if not m and has_table_header:
            tm = _TXN_START_TABLE_RE.match(line)
            if tm:
                line_lower = line.lower()
                if "period covered" in line_lower and re.search(r"\b\d{2}\s+[A-Za-z]{3}\s+\d{4}\s+to\s+\d{2}\s+[A-Za-z]{3}\s+\d{4}\b", line, re.IGNORECASE):
                    continue
                if "statement date" in line_lower:
                    continue
                m = tm
        if m:
            # Start new transaction block
            finalize_block(current_block)

            day = int(m.group("day"))
            mon = (m.group("mon") or "").strip().lower()
            mon_num = _MONTHS.get(mon[:3], None)
            raw_type = (m.groupdict().get("type") or "").strip()
            year_str = m.groupdict().get("year")
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
                if has_table_header:
                    year = current_year
                    if last_seen_mon is not None and mon_num < last_seen_mon:
                        year += 1
                    if period_start_year is not None and period_end_year is not None:
                        if year < period_start_year:
                            year = period_start_year
                        if year > period_end_year:
                            year = period_end_year
                else:
                    year = _infer_year_for_missing_year(
                        day=day,
                        mon_num=mon_num,
                        last_seen_mon=last_seen_mon,
                        current_year=current_year,
                        period_start_year=period_start_year,
                        period_end_year=period_end_year,
                    )
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

    if has_table_header:
        charges_last_seen_mon = None
        charges_current_year = period_end_year if period_end_year is not None else _dt.date.today().year
        injected_any = False
        for cm in _CHARGES_LINE_RE.finditer(all_text):
            day = int(cm.group("day"))
            mon = (cm.group("mon") or "").strip().lower()
            mon_num = _MONTHS.get(mon[:3], None)
            if mon_num is None:
                continue

            year = charges_current_year
            if charges_last_seen_mon is not None and mon_num < charges_last_seen_mon:
                year += 1
            if period_start_year is not None and period_end_year is not None:
                if year < period_start_year:
                    year = period_start_year
                if year > period_end_year:
                    year = period_end_year

            charges_last_seen_mon = mon_num
            charges_current_year = year

            try:
                date_obj = _dt.date(year, mon_num, day)
            except Exception:
                continue

            full_line = cm.group(0) or ""
            monies = _MONEY_RE.findall(full_line)
            monies = [m.strip() for m in monies if m and m.strip()]
            if len(monies) < 2:
                continue

            balance_val = _parse_money(monies[-1])
            if balance_val is None:
                continue

            matched_existing = False
            for txn in transactions:
                txn_balance = txn.get("Balance")
                txn_desc = (txn.get("Description") or "")
                if (
                    txn.get("Date") == date_obj
                    and isinstance(txn_balance, (int, float))
                    and round(float(txn_balance), 2) == round(float(balance_val), 2)
                    and "charges" in txn_desc.lower()
                ):
                    matched_existing = True
                    break

            if matched_existing:
                continue

            rest = _clean_description(cm.group("rest") or "")
            description = "Charges"
            if rest:
                description = f"Charges {rest}".strip()

            transactions.append(
                {
                    "Date": date_obj,
                    "Transaction Type": "Charges",
                    "Description": description,
                    "Amount": 0.0,
                    "Balance": float(balance_val),
                    "_raw_type": "CHG",
                }
            )
            injected_any = True

        if injected_any:
            transactions.sort(key=lambda t: (t.get("Date"), t.get("Balance") if isinstance(t.get("Balance"), (int, float)) else float("inf")))

    # If there are no transactions, return empty list
    if not transactions:
        return []

    if has_table_header:
        summary_start_balance = None
        prev_match = _PREV_BAL_RE.search(all_text)
        if prev_match:
            summary_start_balance = _parse_money(prev_match.group(1))

        prev_balance = summary_start_balance
        for i, txn in enumerate(transactions):
            bcur = txn.get("Balance")
            if not isinstance(bcur, (int, float)):
                continue
            if prev_balance is None and i > 0:
                bprev_row = transactions[i - 1].get("Balance")
                if isinstance(bprev_row, (int, float)):
                    prev_balance = float(bprev_row)
            if isinstance(prev_balance, (int, float)):
                txn["Amount"] = round(float(bcur) - float(prev_balance), 2)
            prev_balance = float(bcur)
    else:
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

        if len(transactions) >= 2:
            first_date = transactions[0].get("Date")
            last_date = transactions[-1].get("Date")
            if isinstance(first_date, _dt.date) and isinstance(last_date, _dt.date) and first_date > last_date:
                transactions.reverse()

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
    start_balance = None
    end_balance = None
    try:
        all_text_chunks = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                all_text_chunks.append(page.extract_text() or "")
        all_text = "\n".join(all_text_chunks)
        has_table_header = "date description paid in" in all_text.lower()

        if has_table_header:
            prev_match = _PREV_BAL_RE.search(all_text)
            new_match = _NEW_BAL_RE.search(all_text)
            if prev_match and new_match:
                start_balance = _parse_money(prev_match.group(1))
                end_balance = _parse_money(new_match.group(1))
                return {"start_balance": start_balance, "end_balance": end_balance}

        opening_matches = list(_OPENING_BALANCE_RE.finditer(all_text))
        if opening_matches:
            start_balance = _parse_money(opening_matches[0].group(1))

        closing_matches = list(_CLOSING_BALANCE_RE.finditer(all_text))
        if closing_matches:
            end_balance = _parse_money(closing_matches[-1].group(2))
    except Exception:
        pass

    if start_balance is not None and end_balance is not None:
        return {"start_balance": start_balance, "end_balance": end_balance}

    try:
        txns = extract_transactions(pdf_path)
    except Exception:
        txns = []

    if not txns:
        return {"start_balance": start_balance, "end_balance": end_balance}

    earliest_txn = txns[0]
    latest_txn = txns[-1]
    first_date = earliest_txn.get("Date")
    last_date = latest_txn.get("Date")
    if not (isinstance(first_date, _dt.date) and isinstance(last_date, _dt.date) and first_date <= last_date):
        earliest_txn = txns[-1]
        latest_txn = txns[0]

    if end_balance is None:
        latest_index = txns.index(latest_txn)
        for i in range(latest_index, -1, -1):
            b = txns[i].get("Balance")
            if isinstance(b, (int, float)):
                end_balance = float(b)
                break
        if end_balance is None:
            for i in range(latest_index + 1, len(txns)):
                b = txns[i].get("Balance")
                if isinstance(b, (int, float)):
                    end_balance = float(b)
                    break

    if start_balance is None:
        earliest_balance = earliest_txn.get("Balance")
        earliest_amount = earliest_txn.get("Amount")
        if isinstance(earliest_balance, (int, float)) and isinstance(earliest_amount, (int, float)):
            start_balance = round(float(earliest_balance) - float(earliest_amount), 2)

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_account_holder_name(pdf_path) -> str:
    """
    Best-effort extraction of the client/account name.
    NatWest export includes: "Account name: <NAME>"
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""

            txt = pdf.pages[0].extract_text() or ""
            if not txt.strip() and len(pdf.pages) > 1:
                txt = pdf.pages[1].extract_text() or ""
    except Exception:
        return ""

    def _normalise_name_text(value: str) -> str:
        text = value or ""
        text = re.sub(r"\b([A-Z])\s{2,}([A-Z]{2,})\b", r"\1\2", text)
        text = re.sub(r"\b([A-Z]{2,})\s{2,}([A-Z]{2,})\b", r"\1\2", text)
        text = re.sub(r"\s+", " ", text).strip()
        text = re.sub(r"(?<=\b[A-Z])\s+(?=[A-Z]\b)", "", text)
        while "  " in text:
            text = text.replace("  ", " ")
        return text.strip()

    if "date description paid in" in txt.lower():
        lines_raw = [(line or "").strip() for line in txt.splitlines()]
        lines_raw = [line for line in lines_raw if line]
        header_idx = -1
        for idx, line in enumerate(lines_raw):
            if "account name" in line.lower() and "account no" in line.lower() and "sort code" in line.lower():
                header_idx = idx
                break

        if header_idx >= 0:
            stop_marker_re = re.compile(
                r"\b(Current\s+Account|Summary|Statement\s+Date|Period\s+Covered|Transactions?)\b",
                re.IGNORECASE,
            )
            trading_lines = []
            for line in lines_raw[header_idx + 1:header_idx + 12]:
                if stop_marker_re.search(line):
                    break
                if re.search(r"\d", line):
                    if trading_lines:
                        break
                    continue
                if re.fullmatch(r"[A-Z&/\-\s']+", line):
                    trading_lines.append(line)
                    continue
                if trading_lines:
                    break

            if trading_lines:
                name = _normalise_name_text(" ".join(trading_lines))
                if name:
                    return name

    lines = [re.sub(r"\s+", " ", (line or "")).strip() for line in txt.splitlines()]
    lines = [line for line in lines if line]

    for line in lines:
        m = _ACCOUNT_NAME_RE.match(line)
        if m:
            name = _normalise_name_text(m.group("name") or "")
            if name and name.lower() not in {"transactions"}:
                return name

    normalised_text = "\n".join(lines)
    inline = _ACCOUNT_NAME_INLINE_RE.search(normalised_text)
    if inline:
        name = _normalise_name_text(inline.group(2) or "")
        if name and name.lower() not in {"transactions"}:
            return name

    for idx, line in enumerate(lines):
        if _ACCOUNT_NAME_NEXT_LINE_LABEL_RE.match(line) and idx + 1 < len(lines):
            candidate = _normalise_name_text(lines[idx + 1])
            if re.match(r"^[A-Z][A-Z\s'\-]{3,}$", candidate, re.IGNORECASE):
                return candidate

    for line in lines:
        if re.match(r"^(MR|MRS|MS|MISS|DR)\b", line, re.IGNORECASE) and not re.search(r"\d", line):
            return _normalise_name_text(line)

    blocked_headers = {
        "NATWEST", "STATEMENT", "PAGE", "SORT CODE", "ACCOUNT NUMBER", "TRANSACTIONS", "ACCOUNT"
    }
    for line in lines:
        upper_line = line.upper()
        if any(h in upper_line for h in blocked_headers):
            continue
        if re.search(r"\d", line):
            continue
        if " " not in line:
            continue
        if not re.match(r"^[A-Z\s'\-]+$", upper_line):
            continue
        return _normalise_name_text(line)

    return ""

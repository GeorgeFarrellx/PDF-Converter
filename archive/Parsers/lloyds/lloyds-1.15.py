"""lloyds-1.15.py

Lloyds Bank (UK) Business Account statement parser (text-based PDFs, NO OCR).

Exports exactly:
- extract_transactions(pdf_path) -> list[dict]
- extract_statement_balances(pdf_path) -> dict {start_balance, end_balance}
- extract_account_holder_name(pdf_path) -> str

Each transaction dict keys (exact):
- Date (datetime.date)
- Transaction Type (str)
- Description (str)
- Amount (float; credits +, debits -)
- Balance (float|None)
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Dict, List, Optional

import pdfplumber


# --- Type mapping (Lloyds statement legend) ---
_TYPE_CODE_MAP = {
    "BGC": "Bank Giro Credit",
    "BP": "Bill Payments",
    "CHG": "Charge",
    "CHQ": "Cheque",
    "COR": "Correction",
    "CPT": "Cashpoint",
    "DD": "Direct Debit",
    "DEB": "Debit Card",
    "DEP": "Deposit",
    "FEE": "Fixed Service",
    "FPI": "Faster Payment In",
    "FPO": "Faster Payment Out",
    "MPI": "Mobile Payment In",
    "MPO": "Mobile Payment Out",
    "PAY": "Payment",
    "SO": "Standing Order",
    "TFR": "Transfer",
}

_MONTHS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


# --- Helpers ---

def _to_float(val: Optional[str]) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in {"blank", "none", "-"}:
        return None
    s = s.replace("£", "").replace(",", "").strip()
    # Parentheses indicate negative (unlikely in these statements for amounts, but defensive)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        return None


def _clean_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _title_case_bank_wording(s: str) -> str:
    # Title-case but keep common acronyms uppercase.
    if not s:
        return s
    # basic title case
    t = s.strip().title()
    # restore common acronyms
    for acr in ("UK", "GB", "VAT", "HMRC", "DVLA", "EE"):
        t = re.sub(rf"\b{acr.title()}\b", acr, t)
    return t


def _apply_global_type_rules(tx_type: str, desc: str) -> (str, str):
    """Apply global rules provided by user.

    - Returned Direct Debit -> Type "Direct Debit" and Description keeps prefix.
    - ApplePay/Clearpay/Contactless/endswith GB -> Type "Card Payment"
    - Else: keep bank wording but Title Case
    - Remove type prefix from Description except Returned Direct Debit
    """

    desc = desc or ""
    tx_type = tx_type or ""

    # Returned Direct Debit rule
    if re.search(r"\breturned\s+direct\s+debit\b", (desc + " " + tx_type), flags=re.I):
        tx_type_out = "Direct Debit"
        # Ensure description starts with prefix
        if not desc.lower().startswith("returned direct debit"):
            desc = "Returned Direct Debit " + desc
        desc = _clean_ws(desc)
        return tx_type_out, desc

    # Card Payment rules
    if re.search(r"apple\s*pay", desc, flags=re.I) or re.search(r"clearpay", desc, flags=re.I) or re.search(
        r"contactless", desc, flags=re.I
    ) or desc.rstrip().upper().endswith("GB"):
        tx_type_out = "Card Payment"
    else:
        tx_type_out = _title_case_bank_wording(tx_type)

    # Remove type prefix from description (except returned DD already handled)
    # (Only if the description literally starts with the type wording)
    d_low = desc.strip().lower()
    t_low = tx_type_out.strip().lower()
    if t_low and d_low.startswith(t_low + " "):
        desc = desc.strip()[len(tx_type_out) :].lstrip(" -–—")

    desc = _clean_ws(desc)
    return tx_type_out, desc


def _parse_statement_period(text: str) -> Optional[tuple[_dt.date, _dt.date]]:
    """Parse statement coverage period from header.

    Expected (from samples):
    BUSINESS ACCOUNT 01 December 2025 to 31 December 2025
    """
    if not text:
        return None

    m = re.search(
        r"\bBUSINESS\s+ACCOUNT\b\s+(\d{2})\s+([A-Za-z]+)\s+(\d{4})\s+to\s+(\d{2})\s+([A-Za-z]+)\s+(\d{4})",
        text,
        flags=re.I,
    )
    if not m:
        return None

    d1, mon1, y1, d2, mon2, y2 = m.groups()
    mon1i = _MONTHS.get(mon1.strip().lower()[:3])
    mon2i = _MONTHS.get(mon2.strip().lower()[:3])
    if not mon1i or not mon2i:
        return None

    try:
        start = _dt.date(int(y1), mon1i, int(d1))
        end = _dt.date(int(y2), mon2i, int(d2))
        return start, end
    except Exception:
        return None


def _parse_tx_date(day: str, mon_abbr: str, year_part: Optional[str], period: Optional[tuple[_dt.date, _dt.date]]) -> Optional[_dt.date]:
    mon = _MONTHS.get(mon_abbr.strip().lower()[:3])
    if not mon:
        return None

    # If year is present (e.g., "25"), use it.
    if year_part:
        yy = int(year_part)
        year = 2000 + yy if yy < 80 else 1900 + yy
        try:
            return _dt.date(year, mon, int(day))
        except Exception:
            return None

    # Otherwise infer from statement period
    if period:
        start, end = period
        # Most transactions are within [start, end]; infer by month + rollover.
        # If statement spans a year boundary, months less than start.month likely belong to end.year.
        if start.year != end.year and mon < start.month:
            year = end.year
        else:
            year = start.year
        try:
            return _dt.date(year, mon, int(day))
        except Exception:
            return None

    return None


def _normalize_extracted_text(text: str) -> str:
    """Lloyds PDFs often extract with standalone '.' separators between tokens.

    Example:
      "Balance on 01 April 2025
.
£8,412.46"

    Normalise these so downstream regexes and line parsing are stable.
    """
    if not text:
        return ""
    # Remove lines that are just '.'
    text = re.sub(r"(?m)^\s*\.\s*$", "", text)
    # Remove the common token separator pattern "\n.\n"
    text = text.replace("\n.\n", "\n")
    # Also collapse inline dotted separators " . "
    text = text.replace(" . ", " ")
    return text


def _extract_all_text(pdf_path: str) -> str:
    parts: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            t = p.extract_text(use_text_flow=True) or ""
            t = _normalize_extracted_text(t)
            if t:
                parts.append(t)
    return "\n".join(parts)


# --- Public API ---

def extract_statement_balances(pdf_path: str) -> Dict[str, Optional[float]]:
    """Return statement start/end balances.

    Lloyds quirk (important for continuity):
    The header "Balance on <start date> £X" can already include the first transaction on that day.
    If so, we infer a "true opening" balance by reversing the first transaction amount:
        true_opening = header_start_balance - first_tx_amount
    This keeps both per-PDF reconciliation and cross-statement continuity correct.
    """
    text = _extract_all_text(pdf_path)

    # Examples in samples:
    # Balance on 01 December 2025 £11,129.48
    # Balance on 31 December 2025 £14,372.71
    start_bal = None
    end_bal = None

    m_start = re.search(
        r"Balance\s+on\s+\d{2}\s+[A-Za-z]+\s+\d{4}\s+£?\s*([0-9,]+\.[0-9]{2})",
        text,
        flags=re.I,
    )
    if m_start:
        start_bal = _to_float(m_start.group(1))

    # Prefer the *last* Balance on ... in the statement header area as end balance
    m_all = re.findall(
        r"Balance\s+on\s+\d{2}\s+[A-Za-z]+\s+\d{4}\s+£?\s*([0-9,]+\.[0-9]{2})",
        text,
        flags=re.I,
    )
    if m_all:
        end_bal = _to_float(m_all[-1])

    # Infer "true opening" if header start already includes the first transaction.
    # This applies to both Business and Classic/Premium statement layouts.
    # Pattern we see: the first transaction row has Balance == header start balance, meaning the header start
    # already includes that first movement.
    try:
        if start_bal is not None:
            header_start = float(start_bal)
            txs = extract_transactions(pdf_path)
            if txs:
                # Find the earliest transaction with a running balance
                for tx in txs:
                    if tx.get("Balance") is None:
                        continue
                    bal = float(tx.get("Balance"))
                    amt = float(tx.get("Amount"))
                    if abs(bal - header_start) <= 0.01 and abs(amt) > 0.001:
                        start_bal = round(header_start - amt, 2)
                    break
    except Exception:
        pass

    return {"start_balance": start_bal, "end_balance": end_bal}


def extract_account_holder_name(pdf_path: str) -> str:
    """Extract the best account holder/client name from page 1.

    Handles:
    - Classic/personal statements: "Document requested by:" then name.
    - Business statements: recipient/company name above the address block.

    Key rule: NEVER return an address line (e.g., "4 WOODLAND ROAD" or "HALESOWEN").
    """
    with pdfplumber.open(pdf_path) as pdf:
        first = pdf.pages[0].extract_text(use_text_flow=True) or ""

    first = _normalize_extracted_text(first)
    lines = [ln.strip() for ln in first.splitlines() if ln.strip()]

    def _is_date_line(s: str) -> bool:
        return bool(re.fullmatch(r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}", s.strip()))

    def _looks_like_postcode(s: str) -> bool:
        return bool(re.search(r"\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b", s.upper()))

    def _looks_like_address(s: str) -> bool:
        if _looks_like_postcode(s):
            return True
        if re.search(r"\bROAD\b|\bSTREET\b|\bAVENUE\b|\bLANE\b|\bDRIVE\b|\bCLOSE\b|\bPARK\b|\bWAY\b", s, flags=re.I):
            return True
        if re.search(r"\d", s):
            return True
        return False

    def _is_boilerplate(s: str) -> bool:
        low = s.lower()
        if low.startswith("logo"):
            return True
        if "lloyds bank" in low:
            return True
        if "registered office" in low:
            return True
        if "prudential regulation" in low or "financial conduct" in low:
            return True
        if "authorised" in low or "regulated" in low:
            return True
        if "registration number" in low:
            return True
        if low.startswith("page "):
            return True
        if _is_date_line(s):
            return True
        return False

    # 1) Classic/personal format: "Document requested by:" then name
    for i, ln in enumerate(lines):
        if ln.lower().startswith("document requested by"):
            for j in range(i + 1, min(i + 8, len(lines))):
                cand = lines[j].strip()
                if not cand:
                    continue
                if _is_boilerplate(cand) or _looks_like_address(cand):
                    continue
                return cand

    # 2) Business-style: everything before "Your Account" is header block
    idx_your_account = None
    for i, ln in enumerate(lines):
        if re.search(r"\bYour\s+Account\b", ln, flags=re.I):
            idx_your_account = i
            break

    candidates = lines[: idx_your_account or 0]

    # Best heuristic: find first address line, then take the nearest suitable line ABOVE it.
    first_addr_idx = None
    for i, ln in enumerate(candidates):
        if _looks_like_address(ln):
            first_addr_idx = i
            break

    if first_addr_idx is not None and first_addr_idx > 0:
        for j in range(first_addr_idx - 1, -1, -1):
            cand = candidates[j].strip()
            if not cand:
                continue
            if _is_boilerplate(cand) or _looks_like_address(cand):
                continue
            return cand

    # Secondary heuristic: first uppercase "name-like" line BEFORE any address begins
    search_block = candidates[:first_addr_idx] if first_addr_idx is not None else candidates
    for ln in search_block:
        if _is_boilerplate(ln) or _looks_like_address(ln):
            continue
        if re.fullmatch(r"[A-Z0-9 &'\-.,]+", ln) and len(ln) >= 3:
            # avoid pure postcode fragments
            if _looks_like_postcode(ln):
                continue
            return ln.strip()

    # Fallback: first non-boilerplate, non-address line
    for ln in candidates[:20]:
        if _is_boilerplate(ln) or _looks_like_address(ln):
            continue
        return ln.strip()

    return ""



def extract_transactions(pdf_path: str) -> List[Dict]:
    """Extract transaction rows from Lloyds statement."""

    transactions: List[Dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        # Statement period from full text (helps year inference if a statement ever omits year)
        full_text = "\n".join([_normalize_extracted_text(p.extract_text(use_text_flow=True) or "") for p in pdf.pages])
        period = _parse_statement_period(full_text)

        # Seed previous balance for balance-delta amount derivation (use header start balance, not extract_statement_balances)
        _m_hdr_start = re.search(
            r"Balance\s+on\s+\d{2}\s+[A-Za-z]+\s+\d{4}\s+£?\s*([0-9,]+\.[0-9]{2})",
            full_text,
            flags=re.I,
        )
        header_start_balance: Optional[float] = _to_float(_m_hdr_start.group(1)) if _m_hdr_start else None
        prev_balance: Optional[float] = header_start_balance
        opening_adjusted = False

        # State for multi-line rows
        cur_date: Optional[_dt.date] = None
        pending_value: Optional[str] = None  # one of: "in", "out", "bal"

        cur_desc_parts: List[str] = []

        cur_type_code_or_word: Optional[str] = None
        cur_money_in: Optional[float] = None
        cur_money_out: Optional[float] = None
        cur_balance: Optional[float] = None

        def finalize_if_ready() -> None:
            nonlocal cur_date, cur_desc_parts, cur_type_code_or_word, cur_money_in, cur_money_out, cur_balance, prev_balance, pending_value, opening_adjusted
            if cur_date is None:
                return
            if cur_money_in is None and cur_money_out is None and cur_balance is None:
                return

            desc = _clean_ws(" ".join([p for p in cur_desc_parts if p and p.strip()]))

            # Remove standalone purchase-date tokens like 29NOV25 / 04APR25 (common on Lloyds DEB/CPT rows)
            desc = re.sub(r"\b\d{2}[A-Z]{3}\d{2}\b", "", desc).strip()
            desc = _clean_ws(desc)

            tx_type_raw = (cur_type_code_or_word or "").strip()
            # Expand Lloyds type code to wording where possible
            if tx_type_raw in _TYPE_CODE_MAP:
                tx_type_raw = _TYPE_CODE_MAP[tx_type_raw]

            # Base amount from Money In/Out
            raw_amount = 0.0
            if cur_money_in is not None and cur_money_out is None:
                raw_amount = float(cur_money_in)
            elif cur_money_out is not None and cur_money_in is None:
                raw_amount = -float(cur_money_out)
            elif cur_money_in is not None and cur_money_out is not None:
                raw_amount = float(cur_money_in) - float(cur_money_out)

            amount = raw_amount

            # Override amount using running-balance deltas when available (improves reconciliation)
            if cur_balance is not None and prev_balance is not None:
                try:
                    delta = round(float(cur_balance) - float(prev_balance), 2)

                    if abs(delta) > 0.001:
                        amount = float(delta)
                    else:
                        # Balance did not move.
                        # Usually informational rows -> amount 0.00.
                        amount = 0.0

                        # Critical Lloyds quirk:
                        # If the statement header start balance already includes the first transaction on the start date,
                        # the first transaction row can have Balance == header_start_balance, making delta appear as 0.00.
                        # In that case, we must keep the raw Money In/Out for that first row.
                        if (
                            not opening_adjusted
                            and header_start_balance is not None
                            and cur_balance is not None
                            and abs(float(cur_balance) - float(header_start_balance)) <= 0.01
                            and abs(float(prev_balance) - float(header_start_balance)) <= 0.01
                            and (cur_money_in is not None or cur_money_out is not None)
                            and abs(float(raw_amount)) > 0.001
                        ):
                            amount = raw_amount
                            opening_adjusted = True
                except Exception:
                    pass

            tx_type_final, desc_final = _apply_global_type_rules(tx_type_raw, desc)

            transactions.append(
                {
                    "Date": cur_date,
                    "Transaction Type": tx_type_final,
                    "Description": desc_final,
                    "Amount": float(amount),
                    "Balance": cur_balance if cur_balance is not None else None,
                }
            )

            # Advance running balance seed
            if cur_balance is not None:
                prev_balance = cur_balance

            # Reset row state
            cur_date = None
            cur_desc_parts = []
            cur_type_code_or_word = None
            cur_money_in = None
            cur_money_out = None
            cur_balance = None
            pending_value = None

        def parse_money_in_out(line: str) -> None:
            nonlocal cur_money_in, cur_money_out
            # Handles both:
            #   "Money Out (£) 29.31"  (same line)
            # and
            #   "Money Out (£)" then next line "29.31" (split line)
            m_in = re.search(r"Money\s+In\s*\(£\)\s*(blank[.]|blank|[0-9,]+[.][0-9]{2})", line, flags=re.I)
            if m_in:
                v = m_in.group(1)
                cur_money_in = _to_float(None if v.lower().startswith("blank") else v)

            m_out = re.search(r"Money\s+Out\s*\(£\)\s*(blank[.]|blank|[0-9,]+[.][0-9]{2})", line, flags=re.I)
            if m_out:
                v = m_out.group(1)
                cur_money_out = _to_float(None if v.lower().startswith("blank") else v)

        def parse_balance(line: str) -> None:
            nonlocal cur_balance
            m = re.search(r"Balance\s*\(£\)\s*([0-9,]+[.][0-9]{2})", line, flags=re.I)
            if m:
                cur_balance = _to_float(m.group(1))


        # Iterate pages and parse transaction section
        for page in pdf.pages:
            text = _normalize_extracted_text(page.extract_text(use_text_flow=True) or "")
            if not text.strip():
                continue

            lines = [ln.strip() for ln in text.splitlines()]

            # Keep only after "Your Transactions" if present on this page; otherwise, this page may still be inside tx list.
            if any("Your Transactions" in ln for ln in lines):
                # Start parsing after the heading
                start_idx = 0
                for i, ln in enumerate(lines):
                    if "Your Transactions" in ln:
                        start_idx = i + 1
                        break
                lines = lines[start_idx:]

            for ln in lines:
                s = ln.strip()
                if not s or s == ".":
                    continue

                # If the PDF extracts values on the next line (e.g. "Money Out (£)" then "29.31"),
                # capture them here so they don't get appended into the description.
                if pending_value is not None:
                    if re.fullmatch(r"[0-9,]+[.][0-9]{2}", s):
                        v = _to_float(s)
                        if pending_value == "in":
                            cur_money_in = v
                        elif pending_value == "out":
                            cur_money_out = v
                        elif pending_value == "bal":
                            cur_balance = v
                        pending_value = None
                        continue

                    if re.fullmatch(r"blank[.]?", s, flags=re.I):
                        if pending_value == "in":
                            cur_money_in = None
                        elif pending_value == "out":
                            cur_money_out = None
                        pending_value = None
                        continue


                # Stop at legend section
                if re.search(r"^Transaction\s+types\b", s, flags=re.I):
                    # finish any row in progress
                    finalize_if_ready()
                    break

                # Skip column header noise
                if re.search(r"^Column\b", s, flags=re.I):
                    continue

                # --- Row building ---
                if s.startswith("Date "):
                    # New row starting; finalize any previous row
                    finalize_if_ready()

                    # Date line may contain Description and/or Type+Amounts inline
                    m = re.search(r"Date\s+(\d{2})\s+([A-Za-z]{3})\s+(\d{2})(?:\b|\s)", s)
                    if not m:
                        cur_date = None
                        continue

                    day, mon, yy = m.group(1), m.group(2), m.group(3)
                    cur_date = _parse_tx_date(day, mon, yy, period)

                    # Inline description
                    m_desc = re.search(r"\bDescription\b\s+(.*)$", s)
                    if m_desc:
                        desc_tail = m_desc.group(1)
                        # If Type appears later in the same line, split it out
                        split = re.split(r"\bType\b\s+", desc_tail, maxsplit=1)
                        if split:
                            cur_desc_parts.append(split[0].strip(" ."))
                            if len(split) > 1:
                                # We have inline type + money/balance
                                rest = "Type " + split[1]
                                # Parse type code/word
                                m_type = re.search(r"\bType\b\s+([A-Z]{2,4})\b", rest)
                                if m_type:
                                    cur_type_code_or_word = m_type.group(1).strip()
                                parse_money_in_out(rest)
                                parse_balance(rest)
                    continue

                if cur_date is None:
                    # Not in a transaction row
                    continue

                if s.startswith("Description "):
                    cur_desc_parts.append(s.replace("Description ", "", 1).strip(" ."))
                    continue

                if s == "Type":
                    continue

                if s.startswith("Type ") or re.search(r"\bType\b\s+[A-Z]{2,4}\b", s):
                    m_type = re.search(r"\bType\b\s+([A-Z]{2,4})\b", s)
                    if m_type:
                        cur_type_code_or_word = m_type.group(1).strip()
                    parse_money_in_out(s)
                    parse_balance(s)
                    continue

                if "Money In" in s or "Money Out" in s or s in {"Money In (£)", "Money Out (£)"}:
                    # If the amount isn't on the same line, the next line will be numeric.
                    if re.fullmatch(r"Money\s+In\s*\(£\)", s, flags=re.I):
                        pending_value = "in"
                    elif re.fullmatch(r"Money\s+Out\s*\(£\)", s, flags=re.I):
                        pending_value = "out"

                    parse_money_in_out(s)
                    parse_balance(s)
                    continue

                if s.startswith("Balance"):
                    # If the balance isn't on the same line, the next line will be numeric.
                    if re.fullmatch(r"Balance\s*\(£\)", s, flags=re.I):
                        pending_value = "bal"

                    parse_balance(s)
                    finalize_if_ready()
                    continue


                # If the type code appears on its own line (common when 'Type' is a label line above it), capture it.
                if cur_type_code_or_word is None:
                    m_code = re.match(r"^([A-Z]{2,4})\b", s)
                    if m_code and m_code.group(1) in _TYPE_CODE_MAP:
                        cur_type_code_or_word = m_code.group(1)
                        if "Money In" in s or "Money Out" in s:
                            parse_money_in_out(s)
                            parse_balance(s)
                        continue

                # Any other line within a row: treat as continuation of description
                cur_desc_parts.append(s.strip(" ."))


            # End of page: do not force finalize unless row is ready
            finalize_if_ready()

    return transactions


# --- Notes (kept in-code for maintainability, but does not affect API) ---
#
# Transaction identification (from provided Lloyds samples):
# - Uses pdfplumber page.extract_text(use_text_flow=True) and parses the "Your Transactions" section.
# - Rows start with "Date DD Mon YY". Details may be on one line or split across multiple lines:
#   - "Date ... . Description ..." then a separate "Type ... Money In/Out ... Balance ..." line
#   - or split into Date / Description / Type / Money / Balance lines.
#
# Year inference:
# - Lloyds transaction dates include a 2-digit year in these samples (e.g., "01 Dec 25"); that is used directly.
# - If a future template omits the year, the parser falls back to the statement period parsed from:
#   "BUSINESS ACCOUNT <start date> to <end date>", including Dec→Jan rollover when the period spans years.
#
# Statement balances:
# - Matches: "Balance on <DD Month YYYY> £<amount>" (first match = start; last match = end).

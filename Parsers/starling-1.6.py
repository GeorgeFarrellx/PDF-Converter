# starling-1.6.py
# Starling Bank (UK) statement parser (text-based PDFs, NO OCR)

from __future__ import annotations

import re
import datetime as _dt
from typing import List, Dict, Optional

import pdfplumber


_MONEY_RE = re.compile(r"£?\(?-?\d[\d,]*\.\d{2}\)?")
_DATE_RE = re.compile(r"^(\d{2}/\d{2}/\d{4}|\d{2}/\d{2})\b")


def _parse_money(s: Optional[str]) -> Optional[float]:
    if not s:
        return None
    s = s.strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    s = s.replace("£", "").replace(",", "").strip()
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    try:
        v = float(s)
    except Exception:
        return None
    return -v if neg else v


def _extract_statement_period(full_text: str) -> tuple[Optional[_dt.date], Optional[_dt.date]]:
    m = re.search(
        r"Summary\s+(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})",
        full_text,
        re.IGNORECASE,
    )
    if not m:
        m = re.search(
            r"Date range applicable:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})",
            full_text,
            re.IGNORECASE,
        )
    if not m:
        return None, None

    d1 = _dt.datetime.strptime(m.group(1), "%d/%m/%Y").date()
    d2 = _dt.datetime.strptime(m.group(2), "%d/%m/%Y").date()
    return d1, d2


def _infer_date(date_str: str, period_start: Optional[_dt.date], period_end: Optional[_dt.date]) -> _dt.date:
    # Most Starling statements include the year in-table. This is a fallback for dd/mm.
    if len(date_str) == 10:
        return _dt.datetime.strptime(date_str, "%d/%m/%Y").date()

    dd, mm = map(int, date_str.split("/"))
    year = (period_start.year if period_start else _dt.date.today().year)

    if period_start and period_end and period_start.year != period_end.year:
        # rollover period: months up to end.month assumed in end.year, rest in start.year
        year = period_end.year if mm <= period_end.month else period_start.year

    return _dt.date(year, mm, dd)


def _title_case_type(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s

    out = []
    for tok in s.lower().split():
        if tok in {"&"}:
            out.append("&")
        elif tok in {"atm"}:
            out.append("ATM")
        else:
            # keep punctuation tokens as-is; otherwise capitalize
            out.append(tok.capitalize() if any(ch.isalpha() for ch in tok) else tok)
    return " ".join(out)


def _apply_global_transaction_type_rules(txn_type: str, desc: str) -> tuple[str, str]:
    txn_type_raw = (txn_type or "").strip()
    desc_raw = (desc or "").strip()

    type_u = txn_type_raw.upper()
    desc_u = desc_raw.upper()

    # Returned Direct Debit rule
    if "RETURNED DIRECT DEBIT" in type_u or "RETURNED DIRECT DEBIT" in desc_u:
        new_type = "Direct Debit"
        if not desc_raw.startswith("Returned Direct Debit"):
            desc_raw = f"Returned Direct Debit {desc_raw}".strip()
        return new_type, desc_raw

    # Card Payment rules
    if (
        "APPLE PAY" in desc_u
        or "APPLE PAY" in type_u
        or "CLEARPAY" in desc_u
        or "CLEARPAY" in type_u
        or "CONTACTLESS" in desc_u
        or "CONTACTLESS" in type_u
        or desc_u.endswith(" GB")
        or desc_u.endswith("GB")
    ):
        new_type = "Card Payment"
        # Remove type prefix from description (if present) except returned DD (handled above)
        if txn_type_raw and desc_u.startswith(type_u + " "):
            desc_raw = desc_raw[len(txn_type_raw) :].lstrip()
        return new_type, desc_raw

    # Otherwise: keep bank wording but Title Case
    new_type = _title_case_type(txn_type_raw)

    # Remove type prefix from description (if present)
    if txn_type_raw and desc_u.startswith(type_u + " "):
        desc_raw = desc_raw[len(txn_type_raw) :].lstrip()

    return new_type, desc_raw


def _group_lines(words: list[dict], y_tol: float = 2.5) -> list[list[dict]]:
    words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: list[list[dict]] = []
    cur: list[dict] = []
    cur_y: Optional[float] = None

    for w in words_sorted:
        if cur_y is None or abs(w["top"] - cur_y) <= y_tol:
            cur.append(w)
            cur_y = w["top"] if cur_y is None else (cur_y * 0.7 + w["top"] * 0.3)
        else:
            lines.append(cur)
            cur = [w]
            cur_y = w["top"]

    if cur:
        lines.append(cur)

    return lines


def _dynamic_inout_cutoff_from_money_xs(xs: list[float], default_cutoff: float) -> float:
    """Infer the IN/OUT split using money x-positions on a headerless page.

    Starling aligns IN and OUT amounts in two vertical columns. For pages without a table header,
    we infer the split point from the distribution of money token x-positions (excluding balances).
    """
    if not xs:
        return default_cutoff

    xs2 = sorted(float(x) for x in xs)
    if len(xs2) < 8:
        return default_cutoff

    # Use lower/upper terciles as representative column centers.
    left = xs2[len(xs2) // 3]
    right = xs2[(2 * len(xs2)) // 3]

    # If there isn't a clear separation, keep default.
    if (right - left) < 20:
        return default_cutoff

    return (left + right) / 2.0


def _dynamic_balance_threshold_from_money_xs(xs: list[float], default_threshold: float) -> float:
    """Infer the balance-column threshold on a page.

    The balance column is the far-right money column. On some pages, a simple width-based
    heuristic underestimates this and incorrectly excludes OUT amounts. This function bumps
    the threshold to just left of the rightmost money column when appropriate.
    """
    if not xs:
        return default_threshold

    mx = max(float(x) for x in xs)

    # Only adjust when the rightmost money column is clearly to the right of the default.
    if (mx - default_threshold) > 30:
        return mx - 25

    return default_threshold



def _get_table_header_info(page) -> Optional[tuple[dict, float]]:
    # Finds the header row containing DATE/TYPE/TRANSACTION etc and returns x positions + header y.
    words = page.extract_words(x_tolerance=1, y_tolerance=2, keep_blank_chars=False)
    candidates = [w for w in words if w.get("text") == "DATE"]
    if not candidates:
        return None

    best = None
    for c in candidates:
        y = c["top"]
        near = [w for w in words if abs(w["top"] - y) <= 5]
        texts = {w["text"] for w in near}
        if "TYPE" in texts and "TRANSACTION" in texts:
            best = c
            break

    if best is None:
        best = candidates[0]

    y = best["top"]
    near = [w for w in words if abs(w["top"] - y) <= 5]
    xs = {
        w["text"]: w["x0"]
        for w in near
        if w["text"] in {"DATE", "TYPE", "TRANSACTION", "IN", "OUT", "ACCOUNT", "BALANCE"}
    }

    if "DATE" not in xs or "TYPE" not in xs or "TRANSACTION" not in xs:
        return None

    return xs, y


# Common Starling type prefixes (used for headerless pages / date-omitted rows)
_KNOWN_TYPES = [
    "FASTER PAYMENT",
    "ONLINE PAYMENT",
    "CHIP & PIN",
    "CHIP & PIN REFUND",
    "CONTACTLESS",
    "APPLE PAY",
    "DIRECT DEBIT",
    "STANDING ORDER",
    "BANK TRANSFER",
    "CASH WITHDRAWAL",
    "ATM",
]


def _detect_type_prefix(rest: str) -> tuple[str, str]:
    """Given text after the date, split into (type, description) using known prefixes."""
    r = (rest or "").strip()
    if not r:
        return "", ""

    r_u = r.upper()
    for t in _KNOWN_TYPES:
        if r_u.startswith(t + " "):
            return t, r[len(t) :].strip()
        if r_u == t:
            return t, ""

    # Fallback: first token as type
    parts = r.split(None, 1)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1].strip()


def extract_statement_period(pdf_path: str) -> tuple[Optional[_dt.date], Optional[_dt.date]]:
    """Public wrapper to extract the statement coverage period (start_date, end_date).

    Returns (None, None) if not found or if the PDF cannot be read.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return None, None

    try:
        return _extract_statement_period(full_text)
    except Exception:
        return None, None



def extract_transactions(pdf_path: str) -> List[Dict]:
    transactions: List[Dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        # IMPORTANT: keep this as a literal "\n" string (do not split across lines)
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        period_start, period_end = _extract_statement_period(full_text)

        # Carry these across pages (critical for headerless continuation pages)
        prev_txn: Optional[Dict] = None
        prev_balance: Optional[float] = None
        last_date: Optional[_dt.date] = None

        # Carry column x-positions from the most recent header page (Starling often drops the header on continuation pages)
        last_cols: Optional[dict] = None

        for page in pdf.pages:
            hdr = _get_table_header_info(page)

            # Default/heuristic column bounds for pages without header
            page_width = float(getattr(page, "width", 0) or 0)
            bal_threshold = page_width * 0.78 if page_width else 440.0  # far-right balances
            inout_cutoff = page_width * 0.62 if page_width else 360.0  # IN typically left of OUT

            # If we have a header, prefer using exact column x-positions and store them for continuation pages
            if hdr is not None:
                xs, hdr_y = hdr
                date_x = xs["DATE"]
                type_x = xs.get("TYPE", date_x + 60)
                trans_x = xs.get("TRANSACTION", type_x + 80)
                in_x = xs.get("IN", trans_x + 240)
                out_x = xs.get("OUT", in_x + 40)
                bal_x = xs.get("ACCOUNT", xs.get("BALANCE", out_x + 50))
                inout_cutoff = (in_x + out_x) / 2.0 if (in_x is not None and out_x is not None) else (in_x + 40)

                last_cols = {
                    "in_x": float(in_x) if in_x is not None else None,
                    "out_x": float(out_x) if out_x is not None else None,
                    "bal_x": float(bal_x) if bal_x is not None else None,
                }
            else:
                # Headerless continuation page: reuse the last known column positions if available
                if last_cols and last_cols.get("bal_x"):
                    bal_threshold = float(last_cols["bal_x"]) - 10.0
                if last_cols and last_cols.get("in_x") and last_cols.get("out_x"):
                    inout_cutoff = (float(last_cols["in_x"]) + float(last_cols["out_x"])) / 2.0

            words = page.extract_words(x_tolerance=1, y_tolerance=2, keep_blank_chars=False)
            lines = _group_lines(words)

            # For headerless continuation pages without a carried header (rare):
            # derive a better balance-column threshold and IN/OUT split from the page's own money positions
            if hdr is None and (not last_cols or not last_cols.get("bal_x")):
                all_money_xs = [
                    w["x0"]
                    for w in words
                    if _MONEY_RE.fullmatch(w.get("text", "")) and w.get("x0", 0) > 0
                ]
                bal_threshold = _dynamic_balance_threshold_from_money_xs(all_money_xs, bal_threshold)

                # ignore far-left footer/interest-table money by requiring x0 > 300
                money_xs = [
                    w["x0"]
                    for w in words
                    if _MONEY_RE.fullmatch(w.get("text", ""))
                    and w.get("x0", 0) < bal_threshold
                    and w.get("x0", 0) > 300
                ]
                inout_cutoff = _dynamic_inout_cutoff_from_money_xs(money_xs, inout_cutoff)

            for ln in lines:
                top = min(w["top"] for w in ln)

                # If we have a header, only parse below it
                if hdr is not None:
                    _, hdr_y = hdr
                    if top < hdr_y + 5:
                        continue

                line_text = " ".join(w["text"] for w in ln).strip()
                if not line_text:
                    continue

                # Footer markers: stop this page
                if (
                    line_text.startswith("We charge interest")
                    or line_text.startswith("Starling Bank is registered")
                    or "Our terms and conditions" in line_text
                ):
                    break

                if "END OF DAY" in line_text:
                    continue

                # Opening balance row (not a transaction)
                if "OPENING BALANCE" in line_text.upper():
                    monies = [w for w in ln if _MONEY_RE.fullmatch(w["text"])]
                    if monies:
                        prev_balance = _parse_money(monies[-1]["text"])
                    continue

                # Identify date (or date-omitted transaction rows)
                m = _DATE_RE.match(line_text)
                row_date: Optional[_dt.date] = None
                rest_text: str = ""

                if m:
                    date_str = m.group(1)
                    row_date = _infer_date(date_str, period_start, period_end)
                    last_date = row_date
                    rest_text = line_text[m.end() :].strip()
                else:
                    if last_date is None:
                        continue

                    # date omitted: treat as a new transaction row if it starts with a known type and has money
                    rt_u = line_text.upper()
                    has_money = bool(_MONEY_RE.search(line_text))
                    starts_like_txn = any(rt_u.startswith(t + " ") or rt_u == t for t in _KNOWN_TYPES)
                    if has_money and starts_like_txn:
                        row_date = last_date
                        rest_text = line_text
                    else:
                        # otherwise it is a wrapped description continuation
                        if prev_txn is not None:
                            prev_txn["Description"] = (prev_txn["Description"] + " " + line_text).strip()
                        continue

                # Extract money tokens with x-positions
                money_words = [w for w in ln if _MONEY_RE.fullmatch(w["text"])]
                balance: Optional[float] = None
                amount: Optional[float] = None

                # Determine balance (far-right money)
                if hdr is not None:
                    xs, _ = hdr
                    bal_x_exact = xs.get("ACCOUNT", xs.get("BALANCE"))
                    if bal_x_exact is not None:
                        bal_candidates = [w for w in money_words if w["x0"] >= (float(bal_x_exact) - 10)]
                    else:
                        bal_candidates = [w for w in money_words if w["x0"] >= bal_threshold]
                else:
                    if last_cols and last_cols.get("bal_x") is not None:
                        bal_candidates = [w for w in money_words if w["x0"] >= (float(last_cols["bal_x"]) - 10)]
                    else:
                        bal_candidates = [w for w in money_words if w["x0"] >= bal_threshold]

                if bal_candidates:
                    w_bal = max(bal_candidates, key=lambda z: z["x0"])
                    balance = _parse_money(w_bal["text"])

                # Determine amount (money that's NOT the balance)
                if hdr is not None:
                    xs, _ = hdr
                    bal_x_exact = xs.get("ACCOUNT", xs.get("BALANCE"))
                    cutoff_x = (float(bal_x_exact) - 10) if bal_x_exact is not None else bal_threshold
                else:
                    if last_cols and last_cols.get("bal_x") is not None:
                        cutoff_x = float(last_cols["bal_x"]) - 10.0
                    else:
                        cutoff_x = bal_threshold

                amt_candidates = [w for w in money_words if w["x0"] < cutoff_x]
                if amt_candidates:
                    w_amt = max(amt_candidates, key=lambda z: z["x0"])
                    amount_abs = _parse_money(w_amt["text"])

                    if amount_abs is not None:
                        sign: Optional[int] = None

                        # Best: infer sign from balance delta
                        if (balance is not None) and (prev_balance is not None):
                            delta = balance - prev_balance
                            if abs(abs(delta) - abs(amount_abs)) <= 0.05:
                                sign = 1 if delta > 0 else -1

                        # Fallback: infer based on column position
                        if sign is None:
                            if hdr is None and last_cols and last_cols.get("in_x") and last_cols.get("out_x"):
                                # Use distance to IN/OUT headers (more stable than a single cutoff)
                                dx_in = abs(float(w_amt["x0"]) - float(last_cols["in_x"]))
                                dx_out = abs(float(w_amt["x0"]) - float(last_cols["out_x"]))
                                sign = 1 if dx_in <= dx_out else -1
                            else:
                                sign = 1 if w_amt["x0"] < inout_cutoff else -1

                        amount = sign * abs(amount_abs)

                # Determine type + description
                txn_type = ""
                desc = ""

                if hdr is not None:
                    xs, _ = hdr
                    type_x = float(xs.get("TYPE", 0))
                    trans_x = float(xs.get("TRANSACTION", 0))

                    type_words = [
                        w["text"]
                        for w in ln
                        if w["x0"] >= (type_x - 1)
                        and w["x0"] < (trans_x - 1)
                        and not _MONEY_RE.fullmatch(w["text"])
                    ]
                    txn_type = " ".join(type_words).strip()

                    first_money_x = min((w["x0"] for w in money_words), default=10**9)
                    bal_x_exact = float(xs.get("ACCOUNT", xs.get("BALANCE", first_money_x)))

                    desc_words = [
                        w["text"]
                        for w in ln
                        if w["x0"] >= (trans_x - 1)
                        and w["x0"] < min(first_money_x - 1, bal_x_exact - 20)
                    ]
                    desc = " ".join(desc_words).strip()
                else:
                    # headerless: parse from text after date, stripping money tokens
                    tmp = _MONEY_RE.sub(" ", rest_text)
                    tmp = " ".join(tmp.split())
                    txn_type, desc = _detect_type_prefix(tmp)

                txn_type_norm, desc_norm = _apply_global_transaction_type_rules(txn_type, desc)

                row = {
                    "Date": row_date,
                    "Transaction Type": txn_type_norm,
                    "Description": desc_norm,
                    "Amount": amount,
                    "Balance": balance,
                }

                transactions.append(row)
                prev_txn = row
                if balance is not None:
                    prev_balance = balance

    return transactions



def extract_statement_balances(pdf_path: str) -> Dict[str, Optional[float]]:
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    m_open = re.search(r"Opening Balance\s+£?([0-9,]+\.\d{2})", text, re.IGNORECASE)
    m_close = re.search(r"Closing Balance\s+£?([0-9,]+\.\d{2})", text, re.IGNORECASE)

    start_balance = _parse_money(m_open.group(1)) if m_open else None
    end_balance = _parse_money(m_close.group(1)) if m_close else None

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_account_holder_name(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = (pdf.pages[0].extract_text() or "").splitlines()

    # Heuristic: the account holder name is usually the first prominent line after the website line.
    blacklist = {
        "24HR CUSTOMER SERVICE:",
        "WWW.STARLINGBANK.COM",
    }

    cleaned = []
    for line in first_page_text:
        s = (line or "").strip()
        if not s:
            continue
        su = s.upper()
        if any(b in su for b in blacklist):
            continue
        if "SUMMARY" in su and re.search(r"\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4}", su):
            continue
        if "STATEMENT" in su:
            continue
        if su.startswith("OPENING BALANCE") or su.startswith("CLOSING BALANCE"):
            continue
        # avoid address-like lines that start with digits
        if re.match(r"^\d", s):
            continue
        cleaned.append(s)

    # In the provided samples this resolves to the company name on the first page.
    return cleaned[0] if cleaned else ""


# Notes (what this parser does):
# - Transaction identification: locates the table header row containing DATE/TYPE/TRANSACTION/IN/OUT/ACCOUNT BALANCE,
#   then parses rows that begin with a date. Amount sign is inferred by IN vs OUT column position.
# - Year inference: most Starling rows include the year; if omitted (DD/MM), year is inferred from the statement Summary period,
#   with Dec→Jan rollover handling when the period spans two years.
# - Statement balances matched: "Opening Balance £..." and "Closing Balance £..." from the statement text.


# ------------------------------
# Minimal self-tests (no PDFs required)
# ------------------------------

def _run_self_tests() -> None:
    # Money parsing
    assert _parse_money("£1,234.56") == 1234.56
    assert _parse_money("(£12.34)") == -12.34
    assert _parse_money("-12.34") == -12.34
    assert _parse_money(None) is None

    # Statement period parsing
    s = "Summary 01/12/2024 - 31/12/2024"
    ps, pe = _extract_statement_period(s)
    assert ps == _dt.date(2024, 12, 1)
    assert pe == _dt.date(2024, 12, 31)

    # Date inference (explicit year)
    assert _infer_date("03/06/2024", None, None) == _dt.date(2024, 6, 3)

    # Date inference (rollover)
    ps = _dt.date(2024, 12, 20)
    pe = _dt.date(2025, 1, 10)
    assert _infer_date("02/01", ps, pe) == _dt.date(2025, 1, 2)
    assert _infer_date("22/12", ps, pe) == _dt.date(2024, 12, 22)

    # Global rules
    t, d = _apply_global_transaction_type_rules("Returned Direct Debit", "XYZ")
    assert t == "Direct Debit"
    assert d.startswith("Returned Direct Debit")

    t, d = _apply_global_transaction_type_rules("Contactless", "TESCO GB")
    assert t == "Card Payment"

    # Type prefix detection
    tt, dd = _detect_type_prefix("FASTER PAYMENT ABC LTD")
    assert tt == "FASTER PAYMENT" and dd == "ABC LTD"

    # Dynamic IN/OUT cutoff (headerless page heuristic)
    xs = [200, 205, 210, 215, 380, 385, 390, 395]
    c = _dynamic_inout_cutoff_from_money_xs(xs, 300)
    assert 250 < c < 350

    xs_one_col = [320, 321, 322, 323, 324, 325, 326, 327]
    c2 = _dynamic_inout_cutoff_from_money_xs(xs_one_col, 300)
    assert c2 == 300

    # Balance threshold inference
    bt = _dynamic_balance_threshold_from_money_xs([420, 476, 535], 440)
    assert bt > 480

    bt2 = _dynamic_balance_threshold_from_money_xs([420, 476], 440)
    assert bt2 == 440

    # Public wrapper error handling (no PDF present)
    ps2, pe2 = extract_statement_period("__this_file_should_not_exist__.pdf")
    assert ps2 is None and pe2 is None


if __name__ == "__main__":
    _run_self_tests()
    print("starling-1.6 self-tests passed")

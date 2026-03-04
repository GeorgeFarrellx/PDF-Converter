# Version: tide.py
# Tide statement parser - text-based PDFs (no OCR)

from __future__ import annotations

import re
import datetime as _dt
from typing import Optional

import pdfplumber


_STATEMENT_PERIOD_RE = re.compile(
    r"Statement\s+for:\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s*-\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})",
    re.IGNORECASE,
)
_BALANCE_LINE_RE = re.compile(
    r"Balance\s*\(£\)\s*on\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s+([0-9,]+\.[0-9]{2})",
    re.IGNORECASE,
)
_DATE_ROW_RE = re.compile(r"^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\b\s*(.*)$")
_MONEY_TAIL_RE = re.compile(r"([0-9,]+\.[0-9]{2})\s+([0-9,]+\.[0-9]{2})\s*$")
_FEE_RANGE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}\s+to\s+\d{4}-\d{2}-\d{2}$")


def _parse_date(d: str) -> Optional[_dt.date]:
    try:
        return _dt.datetime.strptime((d or "").strip(), "%d %b %Y").date()
    except Exception:
        return None


def _to_float(v: str) -> Optional[float]:
    try:
        return float((v or "").replace(",", ""))
    except Exception:
        return None


def extract_statement_period(pdf_path: str):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return (None, None)
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return (None, None)

    m = _STATEMENT_PERIOD_RE.search(text)
    if not m:
        return (None, None)
    return (_parse_date(m.group(1)), _parse_date(m.group(2)))


def extract_statement_balances(pdf_path: str) -> dict:
    period_start, period_end = extract_statement_period(pdf_path)
    balance_points: list[tuple[_dt.date, float]] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return {"start_balance": None, "end_balance": None}
            first_text = pdf.pages[0].extract_text() or ""
    except Exception:
        return {"start_balance": None, "end_balance": None}

    for m in _BALANCE_LINE_RE.finditer(first_text):
        d = _parse_date(m.group(1))
        b = _to_float(m.group(2))
        if d is not None and b is not None:
            balance_points.append((d, b))

    if not balance_points:
        return {"start_balance": None, "end_balance": None}

    by_date = {d: b for d, b in balance_points}
    sorted_points = sorted(balance_points, key=lambda x: x[0])

    start_balance = by_date.get(period_start) if period_start else None
    end_balance = by_date.get(period_end) if period_end else None

    if start_balance is None:
        start_balance = sorted_points[0][1]
    if end_balance is None:
        end_balance = sorted_points[-1][1]

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_account_holder_name(pdf_path: str) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""

    for raw in text.splitlines():
        line = (raw or "").strip()
        if line.lower().startswith("business owner:"):
            return line.split(":", 1)[1].strip()
    return ""


def _clean_description(parts: list[str]) -> str:
    joined = " ".join([(p or "").strip() for p in parts if (p or "").strip()])
    joined = joined.replace("Fee (£): 0.00", "").strip()
    return re.sub(r"\s+", " ", joined).strip()


def _is_merchant_prefix(line: str) -> bool:
    line = (line or "").strip()
    if not line:
        return False
    if _MONEY_TAIL_RE.search(line):
        return False
    upper_chars = [c for c in line if c.isalpha()]
    mostly_upper = bool(upper_chars) and (sum(1 for c in upper_chars if c.isupper()) / len(upper_chars) >= 0.7)
    return (" - " in line) or mostly_upper


def _finalize_txn(txn: dict, out: list[dict]) -> None:
    if not txn:
        return
    out.append(
        {
            "Date": txn.get("date"),
            "Transaction Type": (txn.get("type") or "").strip(),
            "Description": _clean_description(txn.get("description_parts", [])),
            "Amount": None,
            "Balance": txn.get("balance"),
            "_raw_amount": txn.get("raw_amount"),
        }
    )


def extract_transactions(pdf_path: str) -> list[dict]:
    balances = extract_statement_balances(pdf_path)
    start_balance = balances.get("start_balance")

    txns: list[dict] = []
    pending_prefix: list[str] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                in_table = False
                current: dict | None = None

                for raw in text.splitlines():
                    line = (raw or "").strip()
                    low = line.lower()

                    if not in_table:
                        if "date transaction type details paid in (£) paid out (£) balance (£)" in low:
                            in_table = True
                        continue

                    if "bank account legal" in low:
                        if current is not None:
                            _finalize_txn(current, txns)
                            current = None
                        break

                    if not line:
                        continue

                    if line.startswith("Tide Card:"):
                        continue

                    date_match = _DATE_ROW_RE.match(line)
                    if date_match:
                        if current is not None:
                            _finalize_txn(current, txns)

                        date_str = f"{date_match.group(1)} {date_match.group(2)} {date_match.group(3)}"
                        remainder = (date_match.group(4) or "").strip()

                        raw_amount = None
                        balance = None
                        tail = _MONEY_TAIL_RE.search(remainder)
                        if tail:
                            raw_amount = _to_float(tail.group(1))
                            balance = _to_float(tail.group(2))
                            remainder = remainder[: tail.start()].strip()

                        if remainder:
                            if "  " in remainder:
                                tx_type, desc = remainder.split("  ", 1)
                            else:
                                pieces = remainder.split(" ", 1)
                                tx_type = pieces[0]
                                desc = pieces[1] if len(pieces) > 1 else ""
                        else:
                            tx_type, desc = "", ""

                        desc_parts = []
                        if pending_prefix:
                            desc_parts.extend(pending_prefix)
                            pending_prefix = []
                        if desc:
                            desc_parts.append(desc)

                        current = {
                            "date": _parse_date(date_str),
                            "type": tx_type.strip(),
                            "description_parts": desc_parts,
                            "raw_amount": raw_amount,
                            "balance": balance,
                        }
                        continue

                    if current is None:
                        pending_prefix.append(line)
                        continue

                    if (current.get("type") or "").strip() == "Fee" and current.get("raw_amount") is not None and current.get("balance") is not None:
                        if _FEE_RANGE_RE.match(line):
                            current.setdefault("description_parts", []).append(line)
                        else:
                            _finalize_txn(current, txns)
                            current = None
                            pending_prefix.append(line)
                        continue

                    if current.get("raw_amount") is not None and current.get("balance") is not None and (current.get("type") or "").strip() != "Card Transaction":
                        if _is_merchant_prefix(line):
                            _finalize_txn(current, txns)
                            current = None
                            pending_prefix.append(line)
                            continue

                    current.setdefault("description_parts", []).append(line)

                if current is not None:
                    _finalize_txn(current, txns)

    except Exception:
        return []

    if not txns:
        return txns

    first_date = txns[0].get("Date")
    last_date = txns[-1].get("Date")
    reverse_order = bool(first_date and last_date and first_date >= last_date)

    if reverse_order:
        for i in range(0, len(txns) - 1):
            b0 = txns[i].get("Balance")
            b1 = txns[i + 1].get("Balance")
            if b0 is not None and b1 is not None:
                txns[i]["Amount"] = round(b0 - b1, 2)
        blast = txns[-1].get("Balance")
        if blast is not None and start_balance is not None:
            txns[-1]["Amount"] = round(blast - start_balance, 2)
    else:
        for i in range(1, len(txns)):
            b0 = txns[i - 1].get("Balance")
            b1 = txns[i].get("Balance")
            if b0 is not None and b1 is not None:
                txns[i]["Amount"] = round(b1 - b0, 2)
        bfirst = txns[0].get("Balance")
        if bfirst is not None and start_balance is not None:
            txns[0]["Amount"] = round(bfirst - start_balance, 2)

    for row in txns:
        if row.get("Amount") is not None:
            row["Amount"] = round(float(row["Amount"]), 2)
        else:
            raw_amount = row.get("_raw_amount")
            if raw_amount is None:
                raw_amount = 0.0
            mag = abs(float(raw_amount))
            tx_type = (row.get("Transaction Type") or "").strip().lower()
            if tx_type in {"fee", "direct debit", "card transaction"}:
                row["Amount"] = -mag
            else:
                row["Amount"] = mag
        row.pop("_raw_amount", None)

    return txns

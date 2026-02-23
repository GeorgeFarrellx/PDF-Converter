# zempler.py
# Zempler Bank statement parser (text-based PDFs, NO OCR)

from __future__ import annotations

import datetime as _dt
import os
import re
from typing import Optional

import pdfplumber


_TABLE_HEADER_RE = re.compile(r"date.*card.*description.*amount.*balance", re.IGNORECASE)
_TRANSACTION_RE = re.compile(
    r"^\s*(?P<date>\d{2}/\d{2}/\d{4})\s+"
    r"(?:(?P<card>\d{4})\s+)?"
    r"(?P<description>.+?)\s+"
    r"(?P<amount>-?\s*£\s*[\d,]+\.\d{2})\s+"
    r"(?P<balance>-?\s*£\s*[\d,]+\.\d{2})\s*$"
)
_PERIOD_RE = re.compile(r"From\s+(\d{2}/\d{2}/\d{4})\s+to\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_OPENING_BAL_RE = re.compile(r"Opening\s+Balance:\s*(-?\s*£\s*[\d,]+\.\d{2})", re.IGNORECASE)
_CLOSING_BAL_RE = re.compile(r"Closing\s+Balance:\s*(-?\s*£\s*[\d,]+\.\d{2})", re.IGNORECASE)
_ACCOUNT_NAME_RE = re.compile(r"Account held under company name:\s*(.+)", re.IGNORECASE)
_FOOTER_STOP_RE = re.compile(r"zempler\s+bank\s+ltd\s+is\s+registered", re.IGNORECASE)


def _iter_lines(pdf_path: str) -> list[str]:
    lines: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text:
                continue
            lines.extend(text.splitlines())
    return lines


def _parse_money(value: Optional[str]) -> Optional[float]:
    if not value:
        return None
    s = value.strip().replace(" ", "")
    neg = s.startswith("-")
    s = s.replace("£", "").replace(",", "")
    if s.startswith("-"):
        s = s[1:]
    try:
        parsed = float(s)
    except Exception:
        return None
    return -parsed if neg else parsed


def extract_transactions(pdf_path: str) -> list[dict]:
    transactions: list[dict] = []
    lines = _iter_lines(pdf_path)

    in_table = False
    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if _FOOTER_STOP_RE.search(line):
            break

        if _TABLE_HEADER_RE.search(line):
            in_table = True
            continue

        if not in_table:
            continue

        m = _TRANSACTION_RE.match(line)
        if not m:
            continue

        try:
            tx_date = _dt.datetime.strptime(m.group("date"), "%d/%m/%Y").date()
        except Exception:
            continue

        amount = _parse_money(m.group("amount"))
        balance = _parse_money(m.group("balance"))
        if amount is None:
            continue

        card_ending = (m.group("card") or "").strip()
        description = (m.group("description") or "").strip()
        description = re.sub(r"\s+CD\s*\d{4}\b\s*$", "", description, flags=re.IGNORECASE).strip()
        transactions.append(
            {
                "Date": tx_date,
                "Transaction Type": "Card Payment" if card_ending else "Other",
                "Description": description,
                "Amount": amount,
                "Balance": balance,
            }
        )

    return transactions


def extract_statement_balances(pdf_path: str) -> dict:
    opening = None
    closing = None

    try:
        lines = _iter_lines(pdf_path)
        text = "\n".join(lines)
    except Exception:
        return {"start_balance": None, "end_balance": None}

    m_open = _OPENING_BAL_RE.search(text)
    if m_open:
        opening = _parse_money(m_open.group(1))

    m_close = _CLOSING_BAL_RE.search(text)
    if m_close:
        closing = _parse_money(m_close.group(1))

    return {"start_balance": opening, "end_balance": closing}


def extract_account_holder_name(pdf_path: str) -> str:
    try:
        lines = _iter_lines(pdf_path)
    except Exception:
        return ""

    for line in lines:
        m = _ACCOUNT_NAME_RE.search(line or "")
        if m:
            return (m.group(1) or "").strip()

    skip_tokens = (
        "opening balance",
        "closing balance",
        "from ",
        "account held under",
        "date",
        "card",
        "description",
        "amount",
        "balance",
        "zempler",
    )
    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue
        low = line.lower()
        if any(tok in low for tok in skip_tokens):
            continue
        if any(ch.isdigit() for ch in line):
            continue
        return line

    return ""


def _parse_period_from_filename(pdf_path: str) -> tuple[Optional[_dt.date], Optional[_dt.date]]:
    try:
        name = os.path.basename(pdf_path or "")
        m = re.search(
            r"(?P<d1>\d{1,2})[./-](?P<m1>\d{1,2})[./-](?P<y1>\d{2,4})\s*[-–—]\s*(?P<d2>\d{1,2})[./-](?P<m2>\d{1,2})[./-](?P<y2>\d{2,4})",
            name,
        )
        if not m:
            return None, None

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


def extract_statement_period(pdf_path: str) -> tuple[Optional[_dt.date], Optional[_dt.date]]:
    try:
        lines = _iter_lines(pdf_path)
        text = "\n".join(lines)
        m = _PERIOD_RE.search(text)
        if m:
            try:
                start = _dt.datetime.strptime(m.group(1), "%d/%m/%Y").date()
                end = _dt.datetime.strptime(m.group(2), "%d/%m/%Y").date()
                return start, end
            except Exception:
                pass

    except Exception:
        return None, None

    return _parse_period_from_filename(pdf_path)

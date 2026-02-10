from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BANKS = ["barclays", "halifax", "hsbc", "lloyds", "monzo", "nationwide", "natwest", "rbs", "santander", "starling", "tsb"]


@dataclass
class Txn:
    d: date
    desc: str
    amount: float
    ttype: str


def _mk_txns(start: date, opening: float, seed: int) -> tuple[list[Txn], float]:
    _ = random.Random(seed)
    txns: list[Txn] = [
        Txn(start + timedelta(days=0), "Card Payment APPLE PAY COFFEE GB", -21.35, "VIS"),
        Txn(start + timedelta(days=1), "Returned Direct Debit UTILITIES", 35.00, "CR"),
        Txn(start + timedelta(days=2), "Faster Payment PAYROLL CREDIT", 145.00, "CR"),
        Txn(start + timedelta(days=3), "Direct Debit ENERGY SUPPLIER", -12.49, "DD"),
        Txn(start + timedelta(days=4), "Card Purchase SUPERMARKET", -36.15, "VIS"),
        Txn(start + timedelta(days=5), "Bank Transfer SAVINGS", 70.50, "FPI"),
    ]
    bal = opening
    for t in txns:
        bal = round(bal + t.amount, 2)
    return txns, bal


def _lines_for(bank: str, txns: list[Txn], opening: float, closing: float, ps: date, pe: date) -> list[str]:
    lines = ["TEST CLIENT", "Sort code 00-00-00  Account number 00000000"]
    if bank in {"natwest", "rbs"}:
        lines += ["Account name: TEST CLIENT", "Account holder: TEST CLIENT"]

    if bank == "barclays":
        lines += ["Your business accounts - At a glance", f"{ps:%d %b %Y} - {pe:%d %b %Y}", f"Start balance £{opening:,.2f}", "Date Description Money out £ Money in £ Balance £"]
    elif bank == "halifax":
        lines += [f"CURRENT ACCOUNT {ps:%d %B %Y} to {pe:%d %B %Y}", f"Balance on {ps:%d %B %Y} £{opening:,.2f}", "Your Transactions", "Date Description Type Money in Money out Balance"]
    elif bank == "hsbc":
        lines += [f"{ps.day} {ps:%B} to {pe.day} {pe:%B} {pe.year}", f"Opening Balance £{opening:,.2f} Closing Balance £{closing:,.2f}", "Payment type and details                Paid out    Paid in    Balance"]
    elif bank == "lloyds":
        lines += [f"Balance on {ps:%d %B %Y} £{opening:,.2f}", "Your Transactions", "Date Type Description Paid out Paid in Balance"]
    elif bank == "monzo":
        lines += [f"Opening balance £{opening:,.2f}", "Date Description Money out Money in Balance"]
    elif bank == "nationwide":
        lines += [f"Balance brought forward £{opening:,.2f}", "Date Description Money out Money in Balance"]
    elif bank == "natwest":
        lines += [f"{ps.day} {ps:%B} to {pe.day} {pe:%B} {pe.year}", "Date Type Description Amount Balance"]
    elif bank == "rbs":
        lines += [f"Start balance £{opening:,.2f}", f"End balance £{closing:,.2f}", "Date Type Description Amount Balance"]
    elif bank == "santander":
        lines += [f"Balance brought forward £{opening:,.2f}", "Date Description Payments Receipts Balance"]
    elif bank == "starling":
        lines += [f"Opening balance {opening:,.2f}", "Date Type Description Money out Money in Balance"]
    else:
        lines += [f"Effective from: {ps:%d %B %Y} to {pe:%d %B %Y}", f"Balance on {ps:%d %B %Y} £{opening:,.2f}", "Date Payment type Description Paid out Paid in Balance"]

    bal = opening
    for t in txns:
        bal = round(bal + t.amount, 2)
        outv = f"£{abs(t.amount):,.2f}" if t.amount < 0 else ""
        inv = f"£{abs(t.amount):,.2f}" if t.amount > 0 else ""
        code = "DD" if t.amount < 0 else "CR"

        if bank == "barclays":
            out_plain = f"{abs(t.amount):,.2f}" if t.amount < 0 else ""
            in_plain = f"{abs(t.amount):,.2f}" if t.amount > 0 else ""
            lines.append(f"{t.d:%d %b} {t.desc:<28} {out_plain:>9} {in_plain:>9} {bal:>10,.2f}")
        elif bank == "hsbc":
            lines.append(f"{t.d:%d %b %y} {t.ttype} {t.desc:<24} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank == "lloyds":
            tx_type = "Card Payment" if t.amount < 0 else "Direct Credit"
            lines.append(f"{t.d:%d %b} {tx_type:<13} {t.desc[:20]:<20} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank == "monzo":
            lines.append(f"{t.d:%d %b} {t.desc[:30]:<30} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank == "nationwide":
            lines.append(f"{t.d:%d %b} {t.desc[:26]:<26} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank == "santander":
            lines.append(f"{t.d:%d %b} {t.desc[:24]:<24} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank == "starling":
            tx_type = "Card" if t.amount < 0 else "Transfer"
            lines.append(f"{t.d:%d %b} {tx_type:<8} {t.desc[:20]:<20} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        elif bank in {"natwest", "rbs"}:
            signed = f"-£{abs(t.amount):,.2f}" if t.amount < 0 else f"£{abs(t.amount):,.2f}"
            lines.append(f"{t.d:%d %b %Y} {code} {t.desc} {signed} £{bal:,.2f}")
        elif bank == "tsb":
            ptype = "DIRECT DEBIT" if t.amount < 0 else "FASTER PAYMENT"
            lines.append(f"{t.d:%d %b %y} {ptype:<15} {t.desc[:18]:<18} {outv:>10} {inv:>10} £{bal:>10,.2f}")
        else:
            lines.append(f"{t.d:%d %b} {t.desc:<28} {code:<3} {inv:>8} {outv:>8} {bal:>8.2f}")

    lines.append("Additional detail line for Apple Pay")
    lines.append("Returned Direct Debit reference line")
    if bank in {"barclays", "rbs"}:
        lines.append(f"End balance £{closing:,.2f}")
    elif bank in {"halifax", "tsb", "lloyds"}:
        lines.append(f"Balance on {pe:%d %B %Y} £{closing:,.2f}")
    elif bank in {"nationwide", "santander"}:
        lines.append(f"Balance carried forward £{closing:,.2f}")
    else:
        lines.append(f"Closing balance £{closing:,.2f}")
    return lines


def _write_halifax_pdf(path: Path, txns: list[Txn], opening: float, closing: float, ps: date, pe: date) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Courier", 10)

    # Header lines
    c.drawString(36, 810, "TEST CLIENT")
    c.drawString(36, 796, "Document requested by")
    c.drawString(36, 782, "TEST CLIENT")
    c.drawString(36, 768, "Sort code 00-00-00  Account number 00000000")
    c.drawString(36, 754, f"CURRENT ACCOUNT {ps:%d %B %Y} to {pe:%d %B %Y}")
    c.drawString(36, 740, f"Balance on {ps:%d %B %Y} £{opening:,.2f}")
    c.drawString(36, 726, "Your Transactions")

    # Fixed-x table columns for Halifax coordinate parser
    x_date = 60
    x_desc = 140
    x_type = 320
    x_in = 380
    x_out = 450
    x_bal = 520

    y = 712
    c.drawString(x_date, y, "Date")
    c.drawString(x_desc, y, "Description")
    c.drawString(x_type, y, "Type")
    c.drawString(x_in, y, "Money in")
    c.drawString(x_out, y, "Money out")
    c.drawString(x_bal, y, "Balance")

    y -= 14
    bal = opening
    for idx, t in enumerate(txns):
        bal = round(bal + t.amount, 2)
        tx_type = "CR" if t.amount > 0 else "DD"
        money_in = f"{abs(t.amount):,.2f}" if t.amount > 0 else ""
        money_out = f"{abs(t.amount):,.2f}" if t.amount < 0 else ""
        # Keep non-numeric description; add one continuation row for multiline coverage.
        desc = "Returned Direct Debit utility" if "Returned Direct Debit" in t.desc else "Apple Pay coffee" if "APPLE PAY" in t.desc else t.desc.replace("£", "")

        c.drawString(x_date, y, f"{t.d:%d %b}")
        c.drawString(x_desc, y, desc[:26])
        c.drawString(x_type, y, tx_type)
        if money_in:
            c.drawString(x_in, y, money_in)
        if money_out:
            c.drawString(x_out, y, money_out)
        c.drawString(x_bal, y, f"{bal:,.2f}")

        if idx == 0:
            y -= 12
            c.drawString(x_desc, y, "Multiline statement detail")

        y -= 14

    c.drawString(36, y - 6, f"Balance on {pe:%d %B %Y} £{closing:,.2f}")
    c.save()


def _write_pdf(path: Path, lines: Iterable[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Courier", 10)
    y = 810
    for line in lines:
        c.drawString(36, y, line)
        y -= 14
        if y < 70:
            c.showPage()
            c.setFont("Courier", 10)
            y = 810
    c.save()


def generate_all(out_dir: str = "tests/fixtures_synthetic", seed: int = 1) -> None:
    root = Path(out_dir)
    for i, bank in enumerate(BANKS):
        start = date(2024, 4, 1 + i)
        opening = 1000.0 + i * 25
        txns_a, close_a = _mk_txns(start, opening, seed + i)
        if bank == "halifax":
            _write_halifax_pdf(root / bank / "statement_a.pdf", txns_a, opening, close_a, start, start + timedelta(days=9))
        else:
            _write_pdf(root / bank / "statement_a.pdf", _lines_for(bank, txns_a, opening, close_a, start, start + timedelta(days=9)))

        start_b = start + timedelta(days=3)
        txns_b, close_b = _mk_txns(start_b, close_a, seed + 100 + i)
        if bank == "halifax":
            _write_halifax_pdf(root / bank / "statement_b.pdf", txns_b, close_a, close_b, start_b, start_b + timedelta(days=9))
        else:
            _write_pdf(root / bank / "statement_b.pdf", _lines_for(bank, txns_b, close_a, close_b, start_b, start_b + timedelta(days=9)))


if __name__ == "__main__":
    generate_all()

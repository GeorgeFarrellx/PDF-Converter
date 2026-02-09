from __future__ import annotations

import random
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

TARGET_BANKS = ["barclays", "halifax", "hsbc", "lloyds", "monzo", "nationwide", "natwest", "rbs", "santander", "starling", "tsb"]


@dataclass
class Tx:
    date: str
    code: str
    desc: str
    amount: float


def discover_parsers() -> list[str]:
    stems = sorted(p.stem for p in Path("Parsers").glob("*.py"))
    return [s for s in stems if s in TARGET_BANKS]


def _c(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Courier", 10)
    return c


def _money(v: float) -> str:
    return f"£{abs(v):,.2f}"


def _draw_lines(c, lines, x=40, y=800, dy=14):
    for ln in lines:
        c.drawString(x, y, ln)
        y -= dy


def _txs(seed: int) -> list[Tx]:
    random.seed(seed)
    return [
        Tx("01 Jan 24", "DD", "Returned Direct Debit TEST MERCHANT", -20.0),
        Tx("02 Jan 24", "VIS", "Apple Pay TEST MERCHANT GB", -13.5),
        Tx("03 Jan 24", "CR", "TEST SALARY", 200.0),
        Tx("04 Jan 24", "BP", "TEST MERCHANT MAIN", -45.0),
        Tx("05 Jan 24", "SO", "TEST SUBSCRIPTION", -10.0),
        Tx("06 Jan 24", "BGC", "TEST REFUND", 8.0),
    ]


def _render_text_bank(path: Path, bank_title: str, period: str, start_label: str, end_label: str, header: str):
    txs = _txs(1)
    c = _c(path)
    bal = 1000.0
    lines = [bank_title, "TEST CLIENT", "Sort code: 00-00-00", "Account number: 00000000", period, f"{start_label} {_money(bal)}", header]
    for i, t in enumerate(txs):
        bal = round(bal + t.amount, 2)
        paid_out = _money(-t.amount) if t.amount < 0 else ""
        paid_in = _money(t.amount) if t.amount > 0 else ""
        lines.append(f"{t.date:<10} {t.code:<4} {t.desc:<35} {paid_out:<10} {paid_in:<10} {_money(bal)}")
        if i == 1:
            lines.append("CONTINUATION TEST MERCHANT DETAILS")
    lines.append(f"{end_label} {_money(bal)}")
    _draw_lines(c, lines)
    c.save()


def _render_column_bank(path: Path, bank_title: str, period_line: str, header_words: list[str], date_fmt: Callable[[str], str]):
    txs = _txs(1)
    c = _c(path)
    y = 810
    c.drawString(40, y, bank_title); y -= 14
    c.drawString(40, y, "TEST CLIENT"); y -= 14
    c.drawString(40, y, "Sort code 00-00-00 Account number 00000000"); y -= 14
    c.drawString(40, y, period_line); y -= 20
    c.drawString(40, y, " ".join(header_words)); y -= 16
    bal = 1000.0
    for i, t in enumerate(txs):
        bal = round(bal + t.amount, 2)
        c.drawString(40, y, date_fmt(t.date))
        c.drawString(110, y, t.code)
        c.drawString(150, y, t.desc)
        if t.amount > 0:
            c.drawString(380, y, f"{t.amount:,.2f}")
        else:
            c.drawString(320, y, f"{abs(t.amount):,.2f}")
        c.drawString(470, y, f"{bal:,.2f}")
        y -= 15
        if i == 2:
            c.drawString(150, y, "CONTINUATION TEST MERCHANT"); y -= 15
    c.drawString(40, y - 10, "Opening Balance £1,000.00")
    c.drawString(40, y - 24, f"Closing Balance £{bal:,.2f}")
    c.save()


def render_hsbc(path: Path): _render_text_bank(path, "HSBC", "4 June to 3 July 2024", "Opening Balance", "Closing Balance", "Date Payment type and details Paid out Paid in Balance")
def render_tsb(path: Path): _render_text_bank(path, "TSB", "06 August 2024 to 06 September 2024", "Balance on 06 August 2024", "Balance on 06 September 2024", "Date Payment type Details Money Out Money In Balance")
def render_natwest(path: Path): _render_text_bank(path, "National Westminster Bank", "Showing: 01 Jan 2024 to 31 Jan 2024", "Previous Balance", "New Balance", "Date Type Description Paid in Paid out Balance")
def render_rbs(path: Path): _render_text_bank(path, "Royal Bank of Scotland", "Period Covered 01 JAN 2024 to 31 JAN 2024", "Previous Balance", "New Balance", "Date Type Description Paid in Paid out Balance")
def render_monzo(path: Path): _render_text_bank(path, "Monzo Business", "01/01/2024 - 31/01/2024", "Business Account balance", "Business Account balance", "Date Description (GBP) Amount (GBP) Balance")
def render_barclays(path: Path): _render_text_bank(path, "Barclays", "At a glance 01 Jan 2024 to 31 Jan 2024", "Balance brought forward", "Balance carried forward", "DATE DESCRIPTION MONEY OUT £ MONEY IN £ BALANCE £")
def render_lloyds(path: Path): _render_text_bank(path, "Lloyds Bank BUSINESS ACCOUNT", "BUSINESS ACCOUNT 01 January 2024 to 31 January 2024", "Balance on 01 January 2024", "Balance on 31 January 2024", "Date Type Description Money Out (£) Money In (£) Balance (£)")
def render_halifax(path: Path): _render_column_bank(path, "Halifax", "01 May 2024 to 31 May 2024", ["Date", "Type", "Description", "Paid", "out", "Paid", "in", "Balance"], lambda d: d)
def render_nationwide(path: Path): _render_column_bank(path, "Nationwide", "Statement period: 01 Jan 2024 to 31 Jan 2024", ["Date", "Description", "£Out", "£In", "£Balance"], lambda d: d)
def render_santander(path: Path): _render_column_bank(path, "Santander", "Date range 01/01/2024 - 31/01/2024", ["Date", "Description", "Paid", "out", "Paid", "in", "Balance"], lambda d: d.replace(" Jan 24", "/01/2024"))
def render_starling(path: Path): _render_column_bank(path, "Starling", "Date range applicable: 01/01/2024 - 31/01/2024", ["DATE", "TYPE", "TRANSACTION", "IN", "OUT", "ACCOUNT", "BALANCE"], lambda d: d.replace(" Jan 24", "/01/2024"))


TEMPLATES: dict[str, Callable[[Path], None]] = {
    "barclays": render_barclays,
    "halifax": render_halifax,
    "hsbc": render_hsbc,
    "lloyds": render_lloyds,
    "monzo": render_monzo,
    "nationwide": render_nationwide,
    "natwest": render_natwest,
    "rbs": render_rbs,
    "santander": render_santander,
    "starling": render_starling,
    "tsb": render_tsb,
}


def generate_all(out_dir: str = "tests/fixtures_synthetic", seed: int = 1):
    random.seed(seed)
    out_root = Path(out_dir)
    for bank in discover_parsers():
        if bank not in TEMPLATES:
            continue
        bank_dir = out_root / bank
        TEMPLATES[bank](bank_dir / "statement_a.pdf")
        TEMPLATES[bank](bank_dir / "statement_b.pdf")


if __name__ == "__main__":
    generate_all()

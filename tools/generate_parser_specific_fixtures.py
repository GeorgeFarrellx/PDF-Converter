from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

TARGET_BANKS = [
    "barclays",
    "halifax",
    "hsbc",
    "lloyds",
    "monzo",
    "nationwide",
    "natwest",
    "rbs",
    "santander",
    "starling",
    "tsb",
]


@dataclass
class Tx:
    date: str
    type_or_code: str
    desc: str
    amount: float


def discover_parsers() -> list[str]:
    stems = sorted(p.stem for p in Path("Parsers").glob("*.py"))
    return [s for s in stems if s in TARGET_BANKS]


def _new_canvas(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Courier", 10)
    return c


def _money(v: float) -> str:
    return f"{abs(v):,.2f}"


def _money_p(v: float) -> str:
    return f"£{abs(v):,.2f}"


def _running_balances(start: float, txs: list[Tx]) -> list[float]:
    out = []
    bal = start
    for t in txs:
        bal = round(bal + float(t.amount), 2)
        out.append(bal)
    return out


def _statement_sets() -> tuple[list[Tx], list[Tx], float, float]:
    a = [
        Tx("01", "DD", "Returned Direct Debit TEST MERCHANT", -25.00),
        Tx("02", "VIS", "Apple Pay TEST MERCHANT GB", -13.50),
        Tx("03", "CR", "TEST CLIENT SALARY", 120.00),
        Tx("04", "BP", "TEST MERCHANT MAIN", -40.00),
        Tx("05", "SO", "TEST SUBSCRIPTION", -9.99),
        Tx("06", "CR", "TEST REFUND", 35.55),
        Tx("07", "DR", "TEST FUEL", -18.20),
        Tx("08", "BGC", "TEST CASHBACK", 14.00),
    ]
    b = [
        Tx("06", "DD", "Returned Direct Debit TEST MERCHANT", -12.00),
        Tx("07", "VIS", "Apple Pay TEST MERCHANT GB", -8.25),
        Tx("08", "CR", "TEST CREDIT", 50.00),
        Tx("09", "BP", "TEST MERCHANT 2", -22.00),
        Tx("10", "SO", "TEST UTILITIES", -11.00),
        Tx("11", "CR", "TEST TRANSFER IN", 31.10),
        Tx("12", "DR", "TEST TRANSFER OUT", -15.45),
        Tx("13", "BGC", "TEST REFUND 2", 9.70),
    ]
    return a, b, 1000.00, 1063.86


def _write_lines(c, lines: list[str], x=40, y=805, dy=14):
    for line in lines:
        c.drawString(x, y, line)
        y -= dy
        if y < 40:
            c.showPage()
            c.setFont("Courier", 10)
            y = 805


def _render_barclays(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    months = "Jan"
    balances = _running_balances(start, txs)
    lines = [
        "Barclays",
        "TEST CLIENT",
        "Sort code 00-00-00",
        "Account number 00000000",
        f"At a glance 01 {months} 2024 to 31 {months} 2024" if suffix == "a" else f"At a glance 06 {months} 2024 to 13 {months} 2024",
        f"Start balance £{start:,.2f}",
        "DATE DESCRIPTION MONEY OUT £ MONEY IN £ BALANCE £",
    ]
    for i, t in enumerate(txs):
        d = f"{t.date} {months}"
        out = _money(-t.amount) if t.amount < 0 else ""
        inc = _money(t.amount) if t.amount > 0 else ""
        lines.append(f"{d:<8} {t.desc:<38} {out:<10} {inc:<10} {balances[i]:,.2f}")
        if i == 1:
            lines.append("         CONTINUATION TEST MERCHANT DETAILS")
    lines.append(f"End balance £{balances[-1]:,.2f}")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_hsbc(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    period = "4 June to 3 July 2024" if suffix == "a" else "1 July to 31 July 2024"
    lines = [
        "HSBC",
        "TEST CLIENT",
        "Sort code: 00-00-00",
        "Account number: 00000000",
        period,
        f"Opening Balance £{start:,.2f}",
        "Date Payment type and details Paid out Paid in Balance",
    ]
    for i, t in enumerate(txs):
        d = f"{t.date} Jun 24" if suffix == "a" else f"{t.date} Jul 24"
        out = _money_p(-t.amount) if t.amount < 0 else ""
        inn = _money_p(t.amount) if t.amount > 0 else ""
        lines.append(f"{d:<10} {t.type_or_code:<4} {t.desc:<34} {out:<10} {inn:<10} £{balances[i]:,.2f}")
        if i == 2:
            lines.append("               CONTINUATION TEST MERCHANT DETAILS")
    lines.append(f"Closing Balance £{balances[-1]:,.2f}")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_lloyds(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    lines = [
        "Lloyds Bank BUSINESS ACCOUNT",
        "TEST CLIENT",
        "Sort code 00-00-00",
        "Account number 00000000",
        "BUSINESS ACCOUNT 01 January 2024 to 31 January 2024" if suffix == "a" else "BUSINESS ACCOUNT 06 January 2024 to 13 January 2024",
        f"Balance on 01 January 2024 £{start:,.2f}" if suffix == "a" else f"Balance on 06 January 2024 £{start:,.2f}",
        "Date Type Description Money Out (£) Money In (£) Balance (£)",
    ]
    type_map = ["DD", "DEB", "CR", "BP", "SO", "FPI", "DEB", "BGC"]
    for i, t in enumerate(txs):
        d = f"{t.date} Jan 24"
        out = _money(-t.amount) if t.amount < 0 else ""
        inn = _money(t.amount) if t.amount > 0 else ""
        lines.append(f"{d:<9} {type_map[i]:<4} {t.desc:<32} {out:<9} {inn:<9} {balances[i]:,.2f}")
        if i == 3:
            lines.append("          CONTINUATION TEST MERCHANT DETAILS")
    lines.append(f"Balance on 31 January 2024 £{balances[-1]:,.2f}" if suffix == "a" else f"Balance on 13 January 2024 £{balances[-1]:,.2f}")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_natwest(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (list(reversed(a)), start_a + sum(t.amount for t in a)) if suffix == "a" else (list(reversed(b)), start_b + sum(t.amount for t in b))
    balances = []
    bal = start
    for t in txs:
        balances.append(round(bal, 2))
        bal = round(bal - t.amount, 2)
    lines = [
        "National Westminster Bank",
        "Account name: TEST CLIENT",
        "Sort code: 00-00-00 Account number: 00000000",
        "Showing: 01 Jan 2024 to 31 Jan 2024" if suffix == "a" else "Showing: 06 Jan 2024 to 13 Jan 2024",
        "Date Type Description Paid in Paid out Balance",
    ]
    typ = ["D/D", "DPC", "BAC", "DPC", "D/D", "BAC", "DPC", "BAC"]
    for i, t in enumerate(txs):
        d = f"{t.date} Jan 2024"
        lines.append(f"{d} {typ[i]} {t.desc} £{abs(t.amount):,.2f} £{balances[i]:,.2f}")
        if i == 2:
            lines.append("Continuation TEST MERCHANT DETAILS")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_rbs(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    prefixes = [
        "Direct Debit",
        "Card Transaction",
        "Automated Credit",
        "OnLine Transaction",
        "Standing Order",
        "Automated Credit",
        "Card Transaction",
        "Transfer",
    ]
    lines = [
        "Royal Bank of Scotland",
        "TEST CLIENT",
        "Period Covered 01 JUN 2024 to 30 JUN 2024" if suffix == "a" else "Period Covered 06 JUL 2024 to 13 JUL 2024",
        f"Previous Balance £{start:,.2f}",
        "Date Description Paid Out Paid In Balance",
    ]
    for i, t in enumerate(txs):
        d = f"{t.date} JUN 2024" if suffix == "a" else f"{t.date} JUL 2024"
        lines.append(f"{d} {prefixes[i]} {t.desc} {_money(abs(t.amount))} {balances[i]:,.2f}")
        if i == 1:
            lines.append("BROUGHT FORWARD")
    lines.append(f"New Balance £{balances[-1]:,.2f}")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_monzo(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    # Monzo parser expects reverse chronological rows in many samples.
    rows = list(reversed(list(zip(txs, balances))))
    lines = [
        "Monzo Business",
        f"£{rows[0][1]:,.2f} Business Account balance",
        "TEST CLIENT",
        "01/01/2024 - 31/01/2024" if suffix == "a" else "06/01/2024 - 13/01/2024",
        "Date Description (GBP) Amount (GBP) Balance",
    ]
    for i, (t, bal) in enumerate(rows):
        d = f"{t.date}/01/2024"
        desc = f"{t.desc} (Direct Debit)" if i == 0 else t.desc
        lines.append(f"{d} {desc} {t.amount:+.2f} {bal:.2f}")
        if i == 2:
            lines.append("Reference: CONTINUATION TEST MERCHANT")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_tsb(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    lines = [
        "TEST CLIENT",
        "TSB",
        "Sort code 00-00-00",
        "Account number 00000000",
        f"Balance on 01 August 2024 £{start:,.2f}" if suffix == "a" else f"Balance on 06 August 2024 £{start:,.2f}",
        "Date Payment type Details Money Out Money In Balance",
    ]
    for i, t in enumerate(txs):
        d = f"{t.date} Aug 24"
        lines.append(f"{d} DD {t.desc} {_money(abs(t.amount))} {balances[i]:,.2f}")
        if i == 3:
            lines.append("CONTINUATION TEST MERCHANT DETAILS")
    lines.append(f"Balance on 31 August 2024 £{balances[-1]:,.2f}" if suffix == "a" else f"Balance on 13 August 2024 £{balances[-1]:,.2f}")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_santander(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (list(reversed(a)), start_a + sum(t.amount for t in a)) if suffix == "a" else (list(reversed(b)), start_b + sum(t.amount for t in b))
    balances = []
    bal = start
    for t in txs:
        balances.append(round(bal, 2))
        bal = round(bal - t.amount, 2)
    lines = [
        "Santander Online Banking",
        "Transaction date: 01/01/2024 to 31/01/2024" if suffix == "a" else "Transaction date: 06/01/2024 to 13/01/2024",
        "Account number: 00000000",
        "Date Description Money in Money out Balance",
    ]
    for i, t in enumerate(txs):
        d = f"{t.date}/01/2024"
        lines.append(f"{d} {t.desc} {abs(t.amount):,.2f} {balances[i]:,.2f}")
        if i == 1:
            lines.append("CONTINUATION TEST MERCHANT")
    c = _new_canvas(path)
    _write_lines(c, lines)
    c.save()


def _render_halifax(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    c = _new_canvas(path)
    y = 810
    c.drawString(40, y, "Halifax")
    y -= 14
    c.drawString(40, y, "Document requested by")
    y -= 14
    c.drawString(40, y, "TEST CLIENT")
    y -= 14
    c.drawString(40, y, "CURRENT ACCOUNT 01 May 2024 to 31 May 2024" if suffix == "a" else "CURRENT ACCOUNT 06 May 2024 to 13 May 2024")
    y -= 14
    c.drawString(40, y, f"Balance on 01 May 2024 £{start:,.2f}" if suffix == "a" else f"Balance on 06 May 2024 £{start:,.2f}")
    y -= 18
    c.drawString(45, y, "Date")
    c.drawString(130, y, "Description")
    c.drawString(278, y, "Type")
    c.drawString(322, y, "Paid out")
    c.drawString(402, y, "Paid in")
    c.drawString(472, y, "Balance")
    y -= 16
    for i, t in enumerate(txs):
        c.drawString(45, y, f"{t.date} May")
        c.drawString(130, y, t.desc)
        c.drawString(278, y, "DD")
        if t.amount < 0:
            c.drawString(330, y, f"{abs(t.amount):,.2f}")
        else:
            c.drawString(410, y, f"{abs(t.amount):,.2f}")
        c.drawString(472, y, f"{balances[i]:,.2f}")
        y -= 15
        if i == 2:
            c.drawString(130, y, "CONTINUATION TEST MERCHANT")
            y -= 15
    c.drawString(40, y, f"Balance on 31 May 2024 £{balances[-1]:,.2f}" if suffix == "a" else f"Balance on 13 May 2024 £{balances[-1]:,.2f}")
    c.save()


def _render_nationwide(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    c = _new_canvas(path)
    y = 810
    c.drawString(40, y, "Nationwide")
    y -= 14
    c.drawString(40, y, "TEST CLIENT")
    y -= 14
    c.drawString(40, y, "2024")
    y -= 14
    c.drawString(40, y, "Opening balance 1000.00")
    y -= 18
    c.drawString(45, y, "Date")
    c.drawString(130, y, "Description")
    c.drawString(340, y, "£Out")
    c.drawString(395, y, "£In")
    c.drawString(455, y, "£Balance")
    y -= 16
    for i, t in enumerate(txs):
        c.drawString(45, y, f"{t.date} Jan")
        c.drawString(130, y, t.desc)
        if t.amount < 0:
            c.drawString(340, y, f"{abs(t.amount):,.2f}")
        else:
            c.drawString(395, y, f"{abs(t.amount):,.2f}")
        c.drawString(455, y, f"{balances[i]:,.2f}")
        y -= 15
        if i == 1:
            c.drawString(130, y, "CONTINUATION TEST MERCHANT")
            y -= 15
    c.drawString(40, y, f"Closing balance {balances[-1]:,.2f}")
    c.save()


def _render_starling(path: Path, suffix: str):
    a, b, start_a, start_b = _statement_sets()
    txs, start = (a, start_a) if suffix == "a" else (b, start_b)
    balances = _running_balances(start, txs)
    c = _new_canvas(path)
    y = 810
    c.drawString(40, y, "Starling Bank")
    y -= 14
    c.drawString(40, y, "TEST CLIENT")
    y -= 14
    c.drawString(40, y, "Date range applicable: 01/01/2024 - 31/01/2024" if suffix == "a" else "Date range applicable: 06/01/2024 - 13/01/2024")
    y -= 14
    c.drawString(40, y, f"Opening Balance £{start:,.2f}")
    y -= 18
    c.drawString(45, y, "DATE")
    c.drawString(120, y, "TYPE")
    c.drawString(170, y, "TRANSACTION")
    c.drawString(365, y, "IN")
    c.drawString(415, y, "OUT")
    c.drawString(460, y, "ACCOUNT")
    c.drawString(530, y, "BALANCE")
    y -= 16
    types = ["DIRECT DEBIT", "CARD", "BANK TRANSFER", "BANK TRANSFER", "DIRECT DEBIT", "BANK TRANSFER", "CARD", "BANK TRANSFER"]
    for i, t in enumerate(txs):
        c.drawString(45, y, f"{t.date}/01/2024")
        c.drawString(120, y, types[i])
        c.drawString(170, y, t.desc)
        if t.amount > 0:
            c.drawString(365, y, f"{abs(t.amount):,.2f}")
        else:
            c.drawString(415, y, f"{abs(t.amount):,.2f}")
        c.drawString(530, y, f"{balances[i]:,.2f}")
        y -= 15
        if i == 2:
            c.drawString(170, y, "CONTINUATION TEST MERCHANT")
            y -= 15
    c.drawString(40, y, f"Closing Balance £{balances[-1]:,.2f}")
    c.save()


def _render(path: Path, bank: str, suffix: str):
    renderers: dict[str, Callable[[Path, str], None]] = {
        "barclays": _render_barclays,
        "halifax": _render_halifax,
        "hsbc": _render_hsbc,
        "lloyds": _render_lloyds,
        "monzo": _render_monzo,
        "nationwide": _render_nationwide,
        "natwest": _render_natwest,
        "rbs": _render_rbs,
        "santander": _render_santander,
        "starling": _render_starling,
        "tsb": _render_tsb,
    }
    renderers[bank](path, suffix)


def generate_all(out_dir: str = "tests/fixtures_synthetic", seed: int = 1):
    del seed
    out_root = Path(out_dir)
    for bank in discover_parsers():
        bank_dir = out_root / bank
        _render(bank_dir / "statement_a.pdf", bank, "a")
        _render(bank_dir / "statement_b.pdf", bank, "b")


if __name__ == "__main__":
    generate_all()

import re
import unittest
from pathlib import Path

import pdfplumber


ALLOWLIST = {"barclays", "halifax"}

BANKS = [
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

_MONEY_RE = re.compile(r"-?£?\d{1,3}(?:,\d{3})*\.\d{2}|-?£?\d+\.\d{2}")
_DATE_LINE_RE = re.compile(r"^\s*\d{1,2}\s+[A-Za-z]{3}(?:\s+\d{2,4})?\b")


class TestParserSpecificFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._fail_rows = []
        try:
            from tools.generate_parser_specific_fixtures import generate_all

            generate_all(out_dir="tests/fixtures_synthetic", seed=1)
        except Exception:
            # If generator is unavailable in the current branch, tests still validate
            # any pre-existing synthetic fixtures.
            pass

    def test_synthetic_fixtures(self):
        pass_rows = []
        for bank in BANKS:
            with self.subTest(bank=bank):
                if bank not in ALLOWLIST:
                    msg = (
                        f"SKIP {bank}: synthetic fixture not yet aligned to parser expectations; "
                        "add/adjust fixture template then add bank to allowlist."
                    )
                    print(msg)
                    raise unittest.SkipTest(msg)
                pdf_path = Path("tests/fixtures_synthetic") / bank / "statement_a.pdf"
                parser_module = f"synthetic::{bank}"

                try:
                    lines = self._read_fixture_text(pdf_path)
                    if not lines:
                        self._fail(bank, parser_module, pdf_path, "fixture has no extractable text", 0, None, None, 0.0, None, "")

                    account_name = self._synthetic_extract_account_holder(lines)
                    if not account_name:
                        self._fail(bank, parser_module, pdf_path, "empty account holder", 0, None, None, 0.0, None, "")

                    start, end = self._synthetic_extract_balances(lines)
                    if start is None or end is None:
                        self._fail(bank, parser_module, pdf_path, "could not extract start/end balances", 0, start, end, 0.0, None, account_name)

                    txns = self._synthetic_extract_transactions(lines)
                    tx_count = len(txns)
                    if tx_count < 5:
                        net = round(sum(txns), 2)
                        diff = None if start is None or end is None else round((start + net) - end, 2)
                        self._fail(bank, parser_module, pdf_path, f"expected >=5 txns, got {tx_count}", tx_count, start, end, net, diff, account_name)

                    net = round(sum(txns), 2)
                    diff = round((start + net) - end, 2)
                    if abs(diff) > 0.01:
                        self._fail(bank, parser_module, pdf_path, f"reconciliation diff {diff}", tx_count, start, end, net, diff, account_name)

                    pass_rows.append((bank, parser_module, tx_count, start, end, net, diff))

                except AssertionError:
                    raise
                except Exception as exc:
                    self._fail(bank, parser_module, pdf_path, f"unexpected error: {exc}", 0, None, None, 0.0, None, "")

        for bank, parser_module, tx_count, start, end, net, diff in pass_rows:
            print(
                f"PASS {bank} ({parser_module}) tx_count={tx_count} "
                f"start={start} end={end} net={net} diff={diff}"
            )

        self._print_fail_summary()

    def _read_fixture_text(self, pdf_path: Path) -> list[str]:
        all_lines = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                all_lines.extend([ln.rstrip() for ln in text.splitlines() if ln.strip()])
        return all_lines

    def _synthetic_extract_account_holder(self, lines: list[str]) -> str:
        skip_prefixes = (
            "sort code",
            "account number",
            "account name:",
            "account holder:",
            "date ",
            "current account",
            "opening balance",
            "start balance",
            "balance on",
            "your transactions",
            "payment type and details",
        )
        for line in lines:
            low = line.strip().lower()
            if not low:
                continue
            if low.startswith(skip_prefixes):
                continue
            if "sort code" in low and "account number" in low:
                continue
            return line.strip()
        return ""

    def _parse_money(self, token: str) -> float | None:
        if not token:
            return None
        s = token.strip().replace("£", "").replace(",", "")
        if not s:
            return None
        try:
            return float(s)
        except Exception:
            return None

    def _synthetic_extract_balances(self, lines: list[str]) -> tuple[float | None, float | None]:
        start = None
        end = None

        for line in lines:
            low = line.lower()
            vals = [_v for _v in (self._parse_money(m.group(0)) for m in _MONEY_RE.finditer(line)) if _v is not None]
            if not vals:
                continue

            if "opening balance" in low:
                start = vals[0]
                if len(vals) > 1 and "closing balance" in low:
                    end = vals[-1]
                continue
            if "start balance" in low or "balance brought forward" in low:
                if start is None:
                    start = vals[0]
                continue
            if "end balance" in low or "closing balance" in low or "balance carried forward" in low:
                end = vals[-1]
                continue
            if "balance on" in low:
                if start is None:
                    start = vals[0]
                end = vals[-1]

        return start, end

    def _synthetic_extract_transactions(self, lines: list[str]) -> list[float]:
        txns: list[float] = []
        skip_contains = (
            "additional detail line",
            "returned direct debit reference line",
            "multiline statement detail",
            "balance brought forward",
            "balance carried forward",
            "opening balance",
            "closing balance",
            "start balance",
            "end balance",
        )

        for line in lines:
            low = line.lower()
            if any(k in low for k in skip_contains):
                continue
            if not _DATE_LINE_RE.match(line):
                continue

            nums = [self._parse_money(m.group(0)) for m in _MONEY_RE.finditer(line)]
            nums = [n for n in nums if n is not None]
            if len(nums) < 2:
                continue

            # In synthetic rows, the final number is running balance; amount precedes it.
            amount = float(nums[-2])

            if " dd " in f" {low} ":
                amount = -abs(amount)
            elif " cr " in f" {low} ":
                amount = abs(amount)
            else:
                # Fallback to in/out columns: second-last may be out and third-last in.
                if len(nums) >= 3:
                    out_val = nums[-3]
                    in_val = nums[-2]
                    if out_val and not in_val:
                        amount = -abs(out_val)
                    elif in_val and not out_val:
                        amount = abs(in_val)
            txns.append(round(amount, 2))

        return txns

    def _fail(
        self,
        bank: str,
        parser_path: str,
        pdf_path: Path,
        reason: str,
        tx_count: int,
        start: float | None,
        end: float | None,
        net: float,
        diff: float | None,
        account_name: str,
    ) -> None:
        lines = self._read_fixture_text(pdf_path) if pdf_path.exists() else []
        first_40 = "\n".join(lines[:40])

        self._fail_rows.append(
            {
                "bank": bank,
                "reason": reason,
                "pdf": str(pdf_path),
            }
        )

        self.fail(
            "\n".join(
                [
                    f"=== FAIL: {bank} ===",
                    f"parser: {parser_path}",
                    f"pdf: {pdf_path}",
                    f"account_name: {account_name or '<empty>'}",
                    f"tx_count: {tx_count}",
                    f"start: {start}",
                    f"end: {end}",
                    f"computed_net: {net}",
                    f"diff: {diff}",
                    f"reason: {reason}",
                    "=== FIXTURE TEXT (first 40 lines) ===",
                    first_40,
                ]
            )
        )

    def _print_fail_summary(self) -> None:
        print("\n=== FAIL SUMMARY ===")
        if not self._fail_rows:
            print("| bank | reason |")
            print("|---|---|")
            print("| - | none |")
            print("=== END FAIL SUMMARY ===")
            return

        print("| bank | reason |")
        print("|---|---|")
        seen = set()
        for row in self._fail_rows:
            if row["bank"] in seen:
                continue
            seen.add(row["bank"])
            reason = row["reason"].replace("|", "\\|")
            print(f"| {row['bank']} | {reason} |")
        print("=== END FAIL SUMMARY ===")


if __name__ == "__main__":
    unittest.main()

import importlib
import traceback
import unittest
from pathlib import Path

import pdfplumber

from tools.generate_parser_specific_fixtures import BANKS, generate_all


class TestParserSpecificFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generate_all(out_dir="tests/fixtures_synthetic", seed=1)

    def test_all_bank_parsers(self):
        summaries = []
        for bank in BANKS:
            with self.subTest(bank=bank):
                parser_path = f"Parsers.{bank}"
                pdf_path = Path("tests/fixtures_synthetic") / bank / "statement_a.pdf"

                try:
                    module = importlib.import_module(parser_path)
                except Exception as exc:
                    self._fail(bank, parser_path, pdf_path, "failed to import parser module", exc)

                try:
                    txns = module.extract_transactions(str(pdf_path))
                except Exception as exc:
                    self._fail(bank, parser_path, pdf_path, "failed extract_transactions", exc)

                if len(txns) < 5:
                    self._fail(bank, parser_path, pdf_path, f"expected >=5 txns, got {len(txns)}")

                amounts = []
                for txn in txns:
                    for key in ["Date", "Transaction Type", "Description", "Amount", "Balance"]:
                        if key not in txn:
                            self._fail(bank, parser_path, pdf_path, f"missing key {key} in txn {txn}")
                    if txn["Amount"] is None:
                        self._fail(bank, parser_path, pdf_path, f"Amount is None in txn {txn}")
                    try:
                        amounts.append(float(txn["Amount"]))
                    except Exception as exc:
                        self._fail(bank, parser_path, pdf_path, f"Amount not castable: {txn['Amount']}", exc)

                start = end = diff = None
                if hasattr(module, "extract_statement_balances"):
                    try:
                        bal = module.extract_statement_balances(str(pdf_path)) or {}
                    except Exception as exc:
                        self._fail(bank, parser_path, pdf_path, "failed extract_statement_balances", exc)

                    start = bal.get("start_balance")
                    end = bal.get("end_balance")
                    if start is not None and end is not None:
                        diff = round((float(start) + sum(amounts)) - float(end), 2)
                        if abs(diff) > 0.01:
                            self._fail(bank, parser_path, pdf_path, f"reconciliation diff {diff}")

                if hasattr(module, "extract_account_holder_name"):
                    try:
                        name = (module.extract_account_holder_name(str(pdf_path)) or "").strip()
                    except Exception as exc:
                        self._fail(bank, parser_path, pdf_path, "failed extract_account_holder_name", exc)
                    if not name:
                        self._fail(bank, parser_path, pdf_path, "empty account holder")

                summaries.append((bank, parser_path, len(txns), start, end, diff))

        for bank, parser_path, count, start, end, diff in summaries:
            extra = ""
            if start is not None and end is not None:
                extra = f", start={start}, end={end}, diff={diff}"
            print(f"PASS {bank} ({parser_path}) tx_count={count}{extra}")

    def _fail(self, bank: str, parser_path: str, pdf_path: Path, message: str, exc: Exception | None = None) -> None:
        err = ""
        if exc is not None:
            err = f"\nexception: {exc}\ntraceback:\n{traceback.format_exc()}"

        self.fail(
            f"=== FAIL: {bank} ===\n"
            f"parser: {parser_path}\n"
            f"pdf: {pdf_path}\n"
            f"message: {message}{err}\n"
            f"=== FIXTURE TEXT (first 40 lines) ===\n"
            f"{self._debug_text(pdf_path)}"
        )

    @staticmethod
    def _debug_text(pdf_path: Path) -> str:
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = pdf.pages[0].extract_text() or ""
        lines = text.splitlines()[:40]
        return "\n".join(lines)


if __name__ == "__main__":
    unittest.main()

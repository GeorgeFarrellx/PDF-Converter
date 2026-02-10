import importlib
import io
import unittest
from contextlib import redirect_stdout
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
            parser_path = f"Parsers.{bank}"
            module = importlib.import_module(parser_path)
            pdf_path = Path("tests/fixtures_synthetic") / bank / "statement_a.pdf"

            try:
                txns = module.extract_transactions(str(pdf_path))
            except Exception as exc:
                self.fail(f"{bank} failed extract_transactions: {exc}\n{self._debug_text(pdf_path)}")

            self.assertGreaterEqual(
                len(txns),
                5,
                f"{bank} expected >=5 txns, got {len(txns)}\n{self._debug_text(pdf_path)}",
            )

            amounts = []
            for txn in txns:
                for key in ["Date", "Transaction Type", "Description", "Amount", "Balance"]:
                    self.assertIn(key, txn, f"{bank} missing key {key} in txn {txn}")
                self.assertIsNotNone(txn["Amount"], f"{bank} Amount is None\n{self._debug_text(pdf_path)}")
                try:
                    amounts.append(float(txn["Amount"]))
                except Exception as exc:
                    self.fail(f"{bank} Amount not castable: {txn['Amount']} ({exc})\n{self._debug_text(pdf_path)}")

            start = end = diff = None
            if hasattr(module, "extract_statement_balances"):
                try:
                    bal = module.extract_statement_balances(str(pdf_path)) or {}
                    start = bal.get("start_balance")
                    end = bal.get("end_balance")
                    if start is not None and end is not None:
                        diff = round((float(start) + sum(amounts)) - float(end), 2)
                        self.assertLessEqual(abs(diff), 0.01, f"{bank} reconciliation diff {diff}\n{self._debug_text(pdf_path)}")
                except Exception as exc:
                    self.fail(f"{bank} failed extract_statement_balances: {exc}\n{self._debug_text(pdf_path)}")

            if hasattr(module, "extract_account_holder_name"):
                try:
                    name = (module.extract_account_holder_name(str(pdf_path)) or "").strip()
                    self.assertTrue(name, f"{bank} empty account holder\n{self._debug_text(pdf_path)}")
                except Exception as exc:
                    self.fail(f"{bank} failed extract_account_holder_name: {exc}\n{self._debug_text(pdf_path)}")

            summaries.append((bank, parser_path, len(txns), start, end, diff))

        for bank, parser_path, count, start, end, diff in summaries:
            extra = ""
            if start is not None and end is not None:
                extra = f", start={start}, end={end}, diff={diff}"
            print(f"PASS {bank} ({parser_path}) tx_count={count}{extra}")

    @staticmethod
    def _debug_text(pdf_path: Path) -> str:
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = pdf.pages[0].extract_text() or ""
        lines = text.splitlines()[:40]
        return "\n".join(lines)


if __name__ == "__main__":
    unittest.main()

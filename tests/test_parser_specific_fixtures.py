import importlib.util
import traceback
import unittest
from pathlib import Path

from tools.generate_parser_specific_fixtures import TARGET_BANKS, generate_all


class TestParserSpecificFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generate_all(out_dir="tests/fixtures_synthetic", seed=1)

    def _load_parser(self, bank: str):
        parser_path = Path("Parsers") / f"{bank}.py"
        spec = importlib.util.spec_from_file_location(f"parser_{bank}", parser_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod, parser_path

    def test_each_bank_fixture_parses(self):
        for bank in TARGET_BANKS:
            with self.subTest(bank=bank):
                module, parser_path = self._load_parser(bank)
                pdf_path = Path("tests/fixtures_synthetic") / bank / "statement_a.pdf"

                try:
                    txs = module.extract_transactions(str(pdf_path))
                except Exception as exc:
                    self.fail(f"{bank} failed at {parser_path}: {exc}\n{traceback.format_exc()}")

                self.assertIsInstance(txs, list)
                self.assertGreaterEqual(len(txs), 5, f"{bank} returned too few transactions")
                for tx in txs:
                    for k in ["Date", "Transaction Type", "Description", "Amount", "Balance"]:
                        self.assertIn(k, tx, f"{bank} missing key {k}")

                if hasattr(module, "extract_statement_balances"):
                    try:
                        bal = module.extract_statement_balances(str(pdf_path))
                    except Exception as exc:
                        self.fail(f"{bank} balance extraction failed at {parser_path}: {exc}\n{traceback.format_exc()}")
                    self.assertIsNotNone(bal.get("start_balance"), f"{bank} missing start balance")
                    self.assertIsNotNone(bal.get("end_balance"), f"{bank} missing end balance")
                    total = round(sum(float(t["Amount"]) for t in txs), 2)
                    self.assertLessEqual(abs((float(bal["start_balance"]) + total) - float(bal["end_balance"])), 0.01)

                if hasattr(module, "extract_account_holder_name"):
                    try:
                        name = module.extract_account_holder_name(str(pdf_path))
                    except Exception as exc:
                        self.fail(f"{bank} name extraction failed at {parser_path}: {exc}\n{traceback.format_exc()}")
                    self.assertTrue(name == "TEST CLIENT" or bool(str(name).strip()), f"{bank} empty account holder")


if __name__ == "__main__":
    unittest.main()

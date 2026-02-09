import importlib.util
import traceback
import unittest
from pathlib import Path

import pdfplumber

from tools.generate_parser_specific_fixtures import TARGET_BANKS, generate_all


class TestParserSpecificFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generate_all(out_dir="tests/fixtures_synthetic", seed=1)

    def _load_parser(self, bank: str):
        parser_path = Path("Parsers") / f"{bank}.py"
        spec = importlib.util.spec_from_file_location(f"parser_{bank}", parser_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module, str(parser_path)

    def _fixture_excerpt(self, pdf_path: Path) -> str:
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                page = pdf.pages[0] if pdf.pages else None
                text = page.extract_text() if page else ""
        except Exception as exc:
            return f"<unable to extract fixture text: {exc}>"

        lines = (text or "").splitlines()
        excerpt = "\n".join(lines[:40])
        return excerpt or "<empty first-page text>"

    def _fail_with_evidence(self, bank: str, parser_path: str, pdf_path: Path, message: str):
        excerpt = self._fixture_excerpt(pdf_path)
        self.fail(
            f"{bank} failed | parser={parser_path} | pdf={pdf_path}\n"
            f"{message}\n"
            f"fixture_first_page_lines:\n{excerpt}"
        )

    def test_each_bank_fixture_parses(self):
        summaries = []

        for bank in TARGET_BANKS:
            module, parser_path = self._load_parser(bank)
            pdf_path = Path("tests/fixtures_synthetic") / bank / "statement_a.pdf"

            try:
                txs = module.extract_transactions(str(pdf_path))
            except Exception as exc:
                self._fail_with_evidence(
                    bank,
                    parser_path,
                    pdf_path,
                    f"extract_transactions exception: {exc}\n{traceback.format_exc()}",
                )

            if not isinstance(txs, list):
                self._fail_with_evidence(bank, parser_path, pdf_path, f"extract_transactions did not return list: {type(txs)}")

            if len(txs) < 5:
                self._fail_with_evidence(bank, parser_path, pdf_path, f"too few transactions: tx_count={len(txs)}")

            for idx, tx in enumerate(txs):
                for key in ["Date", "Transaction Type", "Description", "Amount", "Balance"]:
                    if key not in tx:
                        self._fail_with_evidence(bank, parser_path, pdf_path, f"missing key '{key}' in tx index={idx}: {tx}")

                amount = tx.get("Amount")
                if amount is None:
                    self._fail_with_evidence(bank, parser_path, pdf_path, f"Amount is None in tx index={idx}: {tx}")
                try:
                    float(amount)
                except Exception:
                    self._fail_with_evidence(bank, parser_path, pdf_path, f"Amount not castable to float in tx index={idx}: {tx}")

            balances_supported = False
            start = None
            end = None
            diff = None

            if hasattr(module, "extract_statement_balances"):
                try:
                    bal = module.extract_statement_balances(str(pdf_path))
                except Exception as exc:
                    self._fail_with_evidence(
                        bank,
                        parser_path,
                        pdf_path,
                        f"extract_statement_balances exception: {exc}\n{traceback.format_exc()}",
                    )

                start = bal.get("start_balance") if isinstance(bal, dict) else None
                end = bal.get("end_balance") if isinstance(bal, dict) else None

                if start is not None and end is not None:
                    balances_supported = True
                    try:
                        start_f = float(start)
                        end_f = float(end)
                    except Exception:
                        self._fail_with_evidence(
                            bank,
                            parser_path,
                            pdf_path,
                            f"non-numeric balances returned: start={start}, end={end}",
                        )

                    total = round(sum(float(t["Amount"]) for t in txs), 2)
                    diff = round((start_f + total) - end_f, 2)
                    if abs(diff) > 0.01:
                        self._fail_with_evidence(
                            bank,
                            parser_path,
                            pdf_path,
                            f"reconciliation mismatch: start={start_f:.2f} total={total:.2f} end={end_f:.2f} diff={diff:.2f}",
                        )

            if hasattr(module, "extract_account_holder_name"):
                try:
                    name = module.extract_account_holder_name(str(pdf_path))
                except Exception as exc:
                    self._fail_with_evidence(
                        bank,
                        parser_path,
                        pdf_path,
                        f"extract_account_holder_name exception: {exc}\n{traceback.format_exc()}",
                    )

                if not (name == "TEST CLIENT" or bool(str(name).strip())):
                    self._fail_with_evidence(bank, parser_path, pdf_path, f"empty account holder name: {name!r}")

            if balances_supported:
                summaries.append(
                    f"{bank}: PASS | parser={parser_path} | tx={len(txs)} | start={float(start):.2f} end={float(end):.2f} diff={float(diff):.2f}"
                )
            else:
                summaries.append(f"{bank}: PASS | parser={parser_path} | tx={len(txs)} | balances=not_supported")

        print("\n" + "\n".join(summaries))


if __name__ == "__main__":
    unittest.main()

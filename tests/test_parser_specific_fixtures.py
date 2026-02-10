import importlib
import os
import traceback
import unittest
from pathlib import Path

import pdfplumber

from tools.generate_parser_specific_fixtures import BANKS, generate_all


class TestParserSpecificFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generate_all(out_dir="tests/fixtures_synthetic", seed=1)
        cls._failure_rows: list[dict[str, str]] = []

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

        self._emit_compact_summary(total=len(BANKS), passed=len(summaries), failures=self._failure_rows)

    def _emit_compact_summary(self, total: int, passed: int, failures: list[dict[str, str]]) -> None:
        failed = len(failures)

        print("\n=== COMPACT FAILURE SUMMARY ===")
        print(f"total={total} passed={passed} failed={failed}")
        if failures:
            for row in failures:
                print(f"- {row['bank']}: {row['message']} ({row['pdf']})")
        else:
            print("- No failing banks")
        print("=== END COMPACT FAILURE SUMMARY ===")

        md_lines = [
            "## Parser-specific Fixtures (compact summary)",
            "",
            f"**Totals:** total={total} passed={passed} failed={failed}",
            "",
            "| Bank | Message | Fixture PDF | Parser Module |",
            "|---|---|---|---|",
        ]

        for row in failures:
            md_lines.append(
                f"| {self._md(row['bank'])} | {self._md(row['message'])} | {self._md(row['pdf'])} | {self._md(row['parser'])} |"
            )

            if row.get("fixture_text"):
                md_lines.extend(
                    [
                        f"<details><summary>{self._md(row['bank'])} fixture text (first 40 lines)</summary>",
                        "",
                        "```text",
                        row["fixture_text"],
                        "```",
                        "</details>",
                        "",
                    ]
                )

        if not failures:
            md_lines.append("| - | - | - | - |")

        summary_md = "\n".join(md_lines) + "\n"
        Path("ci_parser_fixture_summary.md").write_text(summary_md, encoding="utf-8")

        step_summary = os.environ.get("GITHUB_STEP_SUMMARY", "").strip()
        if step_summary:
            with open(step_summary, "a", encoding="utf-8") as f:
                f.write("\n")
                f.write(summary_md)

    @staticmethod
    def _md(value: str) -> str:
        return (value or "").replace("|", "\\|").replace("\n", "<br>").strip()

    def _record_failure(self, bank: str, parser_path: str, pdf_path: Path, message: str, fixture_text: str) -> None:
        for row in self._failure_rows:
            if row["bank"] == bank:
                return
        self._failure_rows.append(
            {
                "bank": bank,
                "parser": parser_path,
                "pdf": str(pdf_path),
                "message": message,
                "fixture_text": fixture_text,
            }
        )

    def _fail(self, bank: str, parser_path: str, pdf_path: Path, message: str, exc: Exception | None = None) -> None:
        fixture_text = self._debug_text(pdf_path)
        self._record_failure(bank, parser_path, pdf_path, message, fixture_text)

        err = ""
        if exc is not None:
            err = f"\nexception: {exc}\ntraceback:\n{traceback.format_exc()}"

        self.fail(
            "========================================\n"
            f"=== FAIL: {bank} ===\n"
            f"parser: {parser_path}\n"
            f"pdf: {pdf_path}\n"
            f"message: {message}{err}\n"
            "=== FIXTURE TEXT (first 40 lines) ===\n"
            f"{fixture_text}\n"
            "========================================"
        )

    @staticmethod
    def _debug_text(pdf_path: Path) -> str:
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = pdf.pages[0].extract_text() or ""
        lines = text.splitlines()[:40]
        return "\n".join(lines)


if __name__ == "__main__":
    unittest.main()

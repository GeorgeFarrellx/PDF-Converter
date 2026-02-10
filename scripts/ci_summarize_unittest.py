from __future__ import annotations

import re
import sys
from pathlib import Path

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

FAIL_RE = re.compile(r"^=== FAIL: ([a-z]+) ===\s*$")
PASS_RE = re.compile(r"^PASS ([a-z]+) \(([^)]*)\).*$")
END_BLOCK_RE = re.compile(r"^=+\s*$")


def _md_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>").strip()


def summarize(log_text: str) -> str:
    data: dict[str, dict[str, str | bool]] = {
        bank: {"status": "UNKNOWN", "message": "", "fixture": "", "parser": "", "multiple": False}
        for bank in BANKS
    }

    current_fail_bank: str | None = None

    for raw_line in log_text.splitlines():
        line = raw_line.rstrip("\n")

        fail_match = FAIL_RE.match(line)
        if fail_match:
            bank = fail_match.group(1)
            current_fail_bank = bank
            if bank not in data:
                data[bank] = {"status": "UNKNOWN", "message": "", "fixture": "", "parser": "", "multiple": False}
            if data[bank]["status"] == "FAIL":
                data[bank]["multiple"] = True
            data[bank]["status"] = "FAIL"
            continue

        pass_match = PASS_RE.match(line)
        if pass_match:
            bank = pass_match.group(1)
            parser_path = pass_match.group(2)
            if bank not in data:
                data[bank] = {"status": "UNKNOWN", "message": "", "fixture": "", "parser": "", "multiple": False}
            if data[bank]["status"] != "FAIL":
                data[bank]["status"] = "PASS"
            if not data[bank]["parser"]:
                data[bank]["parser"] = parser_path
            current_fail_bank = None
            continue

        if current_fail_bank:
            bank_data = data[current_fail_bank]
            if line.startswith("parser:") and not bank_data["parser"]:
                bank_data["parser"] = line.split(":", 1)[1].strip()
            elif line.startswith("pdf:") and not bank_data["fixture"]:
                bank_data["fixture"] = line.split(":", 1)[1].strip()
            elif line.startswith("message:") and not bank_data["message"]:
                bank_data["message"] = line.split(":", 1)[1].strip()
            elif line.startswith("=== FIXTURE TEXT") or END_BLOCK_RE.match(line):
                current_fail_bank = None

    passed = sum(1 for bank in BANKS if data.get(bank, {}).get("status") == "PASS")
    failed = sum(1 for bank in BANKS if data.get(bank, {}).get("status") == "FAIL")
    total = len(BANKS)

    out: list[str] = []
    out.append("## Parser-specific Fixtures Test Summary")
    out.append("")
    out.append(f"**Totals:** passed={passed} failed={failed} total={total}")
    out.append("")
    out.append("| Bank | Status | Message | Fixture | Parser |")
    out.append("|---|---|---|---|---|")

    for bank in BANKS:
        item = data[bank]
        message = str(item["message"] or "")
        if item["status"] == "FAIL" and item["multiple"]:
            message = f"{message} (multiple failures)".strip()
        row = [
            bank,
            str(item["status"]),
            _md_escape(message),
            _md_escape(str(item["fixture"] or "")),
            _md_escape(str(item["parser"] or "")),
        ]
        out.append(f"| {row[0]} | {row[1]} | {row[2]} | {row[3]} | {row[4]} |")

    return "\n".join(out) + "\n"


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python scripts/ci_summarize_unittest.py <unittest.log>", file=sys.stderr)
        return 2

    log_path = Path(sys.argv[1])
    if not log_path.exists():
        print(f"Log file not found: {log_path}", file=sys.stderr)
        return 2

    print(summarize(log_path.read_text(encoding="utf-8", errors="replace")), end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

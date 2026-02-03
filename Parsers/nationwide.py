# Version: nationwide.py
import re
from datetime import datetime
import pdfplumber

YEAR_RE = re.compile(r"\b(20\d{2})\b")
MONEY_RE = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")
MONEY_POUND_RE = re.compile(r"£?\d{1,3}(?:,\d{3})*\.\d{2}")

# Base Nationwide “transaction type” prefixes (as they appear at the start of the description column)
TYPE_PATTERNS = [
    ("Contactless Payment", re.compile(r"^contactless\s+payment\b", re.IGNORECASE)),
    ("Visa purchase", re.compile(r"^visa\s+purchase\b", re.IGNORECASE)),
    ("Card payment", re.compile(r"^card\s+payment\b", re.IGNORECASE)),
    ("Payment to", re.compile(r"^payment\s+to\b", re.IGNORECASE)),
    ("Transfer to", re.compile(r"^transfer\s+to\b", re.IGNORECASE)),
    ("Transfer from", re.compile(r"^transfer\s+from\b", re.IGNORECASE)),
    ("Bank credit", re.compile(r"^bank\s+credit\b", re.IGNORECASE)),
    ("Cash credit", re.compile(r"^cash\s+credit\b", re.IGNORECASE)),
    ("Direct debit", re.compile(r"^direct\s+debit\b", re.IGNORECASE)),
    ("ATM Withdrawal", re.compile(r"^atm\s+withdrawal\b", re.IGNORECASE)),
]

def _to_float(s: str):
    try:
        s = s.replace("£", "").replace(",", "")
        return float(s)
    except Exception:
        return None

def _cluster_lines_by_y(words, y_tol=2.5):
    """
    Group words into visual “rows” based on their y-position (top).
    """
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    cur = []
    cur_y = None

    for w in words:
        y = w["top"]
        if cur_y is None:
            cur_y = y
            cur = [w]
            continue

        if abs(y - cur_y) <= y_tol:
            cur.append(w)
        else:
            lines.append(sorted(cur, key=lambda a: a["x0"]))
            cur_y = y
            cur = [w]

    if cur:
        lines.append(sorted(cur, key=lambda a: a["x0"]))
    return lines

def _line_text(line_words):
    return " ".join(w["text"] for w in line_words).strip()

def _find_table_header(lines):
    """
    Find header row: Date | Description | £Out | £In | £Balance
    Return x positions + header y.
    """
    for line in lines:
        toks = [(w["text"].replace(" ", "").lower(), w) for w in line]

        def find(substr):
            for t, w in toks:
                if substr in t:
                    return w
            return None

        w_date = find("date")
        w_desc = find("description")
        w_out  = find("£out")
        w_in   = find("£in")
        w_bal  = find("£balance")

        if w_date and w_desc and w_out and w_in and w_bal:
            return {
                "header_y": w_date["top"],
                "x_desc": w_desc["x0"],
                "x_out": w_out["x0"],
                "x_in": w_in["x0"],
                "x_bal": w_bal["x0"],
            }
    return None

def _infer_year(all_text: str):
    m = YEAR_RE.search(all_text)
    return int(m.group(1)) if m else None


def _update_year_from_left_zone(left_zone_words, current_year):
    """Update active year using year markers that appear in the Date column.

    Nationwide statements can contain year markers inside the table, e.g.
    a line starting with '2024' (Balance from statement...) and later a
    standalone '2025' line. This updates the active year as we parse.
    """
    if not left_zone_words:
        return current_year

    first = (left_zone_words[0].get("text") or "").strip()
    if len(first) == 4 and first.isdigit() and first.startswith("20"):
        y = int(first)
        if 2000 <= y <= 2099:
            return y

    return current_year

def _try_parse_date(left_zone_words, year):
    """
    Nationwide date in the Date column usually appears as: '31 Oct' (two words).
    """
    if year is None or len(left_zone_words) < 2:
        return None

    d = left_zone_words[0]["text"]
    m = left_zone_words[1]["text"][:3]

    if not (len(d) == 2 and d.isdigit()):
        return None

    try:
        return datetime.strptime(f"{d} {m} {year}", "%d %b %Y").date()
    except Exception:
        return None

def _extract_type_and_clean_first_line(first_desc_line: str):
    """
    Returns (base_type, cleaned_first_line_description_without_type_prefix)

    SPECIAL CASE REQUIRED BY GEORGE:
      Returned Direct Debit -> Transaction Type should be "Direct Debit"
      AND Description should KEEP "Returned Direct Debit ..." at the start.
    """
    s = (first_desc_line or "").strip()
    if not s:
        return "Other", ""

    # Special case: keep phrase in Description, but treat as Direct Debit for typing
    if re.search(r"^returned\s+direct\s+debit\b", s, flags=re.IGNORECASE):
        # base_type is "Direct debit" so it becomes Transaction Type "Direct Debit"
        return "Direct debit", s

    for type_name, rx in TYPE_PATTERNS:
        m = rx.search(s)
        if m:
            cleaned = s[m.end():].strip(" -\t")
            return type_name, cleaned

    return "Other", s

def _normalise_type_titlecase(base_type: str) -> str:
    """
    The user wants Title Case in the Transaction Type column.
    Keep wording consistent with Nationwide, just nicer casing.
    """
    bt = (base_type or "").strip().lower()
    mapping = {
        "direct debit": "Direct Debit",
        "bank credit": "Bank Credit",
        "cash credit": "Cash Credit",
        "payment to": "Payment To",
        "transfer to": "Transfer To",
        "transfer from": "Transfer From",
        "contactless payment": "Contactless Payment",
        "visa purchase": "Visa Purchase",
        "card payment": "Card Payment",
        "atm withdrawal": "ATM Withdrawal",
        "other": "Other",
    }
    return mapping.get(bt, base_type)

def _apply_type_overrides(base_type: str, description: str) -> str:
    """
    George's rules (priority order):

    3.1 Returned direct debit -> Direct Debit
    3.2 Mentions applepay -> Card Payment
    3.3 Mentions clearpay -> Card Payment
    3.4 All contactless payments -> Card Payment
    3.5 If description ends with "GB" -> Card Payment

    Otherwise: keep the (Title Cased) base type.
    """
    desc = (description or "").strip()
    low = desc.lower()

    # 3.1 (we also support it even if the line got through without the special case)
    if low.startswith("returned direct debit"):
        return "Direct Debit"

    # 3.2 + 3.3
    if "applepay" in low or "clearpay" in low:
        return "Card Payment"

    # 3.4
    if (base_type or "").strip().lower() == "contactless payment":
        return "Card Payment"

    # 3.5
    if re.search(r"\bGB\b\s*$", desc):
        return "Card Payment"

    # default: Title Case the base type
    return _normalise_type_titlecase(base_type or "Other")


def extract_account_holder_name(pdf_path: str) -> str:
    """Best-effort: extract the account holder name from page 1.

    Used by Main for Excel page headers and output filename.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""

    titles = ("Mr ", "Mrs ", "Ms ", "Miss ", "Dr ")
    for raw in (text.splitlines() if text else []):
        line = (raw or "").strip()
        if not line:
            continue
        if any(ch.isdigit() for ch in line):
            continue
        if line.startswith(titles):
            if "statement" in line.lower():
                continue
            return line

    return ""


def extract_statement_balances(pdf_path: str):
    """
    Extract statement start/end balances (best-effort).

    George notes these appear on page 1 for his statements.

    Strategy:
      1) Page 1 text regex using several label variants.
      2) Page 1 words/coordinates fallback: find a line that looks like a start/end balance label,
         then grab the money on that line (or the next line).

    Returns: {"start_balance": float|None, "end_balance": float|None}
    """
    start_balance = None
    end_balance = None

    start_labels = [
        "Start balance",
        "Opening balance",
        "Balance brought forward",
        "Balance b/f",
        "Balance at start",
        "Balance at start of statement",
    ]

    end_labels = [
        "End balance",
        "Closing balance",
        "Balance carried forward",
        "Balance c/f",
        "Balance at end",
        "Balance at end of statement",
    ]
    def _find_in_text(text: str, labels: list[str]):
        if not text:
            return None

        # Normalise whitespace
        t = " ".join(text.split())
        # Also keep a no-space variant to catch labels like 'Startbalance'
        t_nospace = "".join(t.split())

        for lbl in labels:
            # 1) Normal spaced label match
            rx = re.compile(
                re.escape(lbl) + r"[: -]*£? *(" + MONEY_RE.pattern + r")",
                flags=re.IGNORECASE
            )
            m = rx.search(t)
            if m:
                return _to_float(m.group(1))

            # 2) No-space label match (e.g. 'Startbalance')
            lbl_ns = "".join(lbl.split())
            rx2 = re.compile(
                re.escape(lbl_ns) + r"[: -]*£? *(" + MONEY_RE.pattern + r")",
                flags=re.IGNORECASE
            )
            m2 = rx2.search(t_nospace)
            if m2:
                return _to_float(m2.group(1))

        return None

    def _tokens(s: str) -> set[str]:
        if not s:
            return set()

        low = s.lower()

        # Split common concatenations seen in extract_words/extract_text
        low = low.replace("startbalance", "start balance")
        low = low.replace("endbalance", "end balance")
        low = low.replace("openingbalance", "opening balance")
        low = low.replace("closingbalance", "closing balance")

        cleaned = re.sub(r"[^a-z0-9 ]+", " ", low)
        return {x for x in cleaned.split() if x}

    def _money_from_line_or_next(lines, idx: int):
        line_text = _line_text(lines[idx])
        m = MONEY_POUND_RE.findall(line_text)
        if m:
            return _to_float(m[-1])
        if idx + 1 < len(lines):
            next_text = _line_text(lines[idx + 1])
            m2 = MONEY_POUND_RE.findall(next_text)
            if m2:
                return _to_float(m2[-1])
        return None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return {"start_balance": None, "end_balance": None}

            page0 = pdf.pages[0]

            # Pass 1: regex over extracted text (page 1)
            text = page0.extract_text() or ""
            start_balance = _find_in_text(text, start_labels)
            end_balance = _find_in_text(text, end_labels)

            if start_balance is not None and end_balance is not None:
                return {"start_balance": start_balance, "end_balance": end_balance}

            # Pass 2: words/coordinates fallback (page 1)
            words = page0.extract_words(use_text_flow=True, keep_blank_chars=False) or []
            if not words:
                return {"start_balance": start_balance, "end_balance": end_balance}

            lines = _cluster_lines_by_y(words, y_tol=2.5)

            for i, line in enumerate(lines):
                lt = _line_text(line)
                toks = _tokens(lt)

                # Start-balance heuristics
                if start_balance is None:
                    if ("start" in toks and "balance" in toks) or ("opening" in toks and "balance" in toks) or ("brought" in toks and "forward" in toks):
                        start_balance = _money_from_line_or_next(lines, i)

                # End-balance heuristics
                if end_balance is None:
                    if ("end" in toks and "balance" in toks) or ("closing" in toks and "balance" in toks) or ("carried" in toks and "forward" in toks):
                        end_balance = _money_from_line_or_next(lines, i)

                if start_balance is not None and end_balance is not None:
                    break

    except Exception:
        pass

    return {"start_balance": start_balance, "end_balance": end_balance}


def extract_transactions(pdf_path: str):
    """
    Output rows are dicts with CAPITALISED headings:
      Date, Transaction Type, Description, Amount, Balance
    """
    # Infer year once (fallback). Some statements switch year mid-table (e.g. Dec -> Jan).
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            all_text += (p.extract_text() or "") + "\n"

    year = _infer_year(all_text)
    current_year = year

    rows = []
    current_date = None
    current_txn = None  # {date, base_type, desc_lines, amount, balance}

    def commit():
        nonlocal current_txn
        if not current_txn:
            return

        desc_lines = [x for x in current_txn["desc_lines"] if x]
        description = " - ".join(desc_lines).strip()

        tx_type = _apply_type_overrides(current_txn["base_type"], description)

        rows.append({
            "Date": current_txn["date"],
            "Transaction Type": tx_type,
            "Description": description,
            "Amount": current_txn["amount"],
            "Balance": "" if current_txn["balance"] is None else current_txn["balance"],
        })
        current_txn = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            if not words:
                continue

            # Build initial line clusters to locate header positions
            lines = _cluster_lines_by_y(words, y_tol=2.5)
            header = _find_table_header(lines)
            if not header:
                continue

            header_y = header["header_y"]
            x_desc = header["x_desc"]
            x_out  = header["x_out"]
            x_in   = header["x_in"]
            x_bal  = header["x_bal"]

            # CRITICAL: cut off the right-side sidebar text so it doesn't contaminate rows
            table_right = x_bal + 55
            table_words = [w for w in words if w["x0"] <= table_right]

            # Re-cluster using only table words
            lines = _cluster_lines_by_y(table_words, y_tol=2.5)

            for line in lines:
                if line[0]["top"] <= header_y + 1:
                    continue

                row_text = _line_text(line)
                if not row_text:
                    continue

                # Skip effective date metadata rows
                if row_text.lower().strip().startswith("effective date"):
                    continue

                # Split into column zones by x
                left_zone = [w for w in line if w["x0"] < x_desc - 5]
                desc_zone = [w for w in line if (x_desc - 5) <= w["x0"] < x_out - 5]
                out_zone  = [w for w in line if x_out - 5 <= w["x0"] < x_in - 5]
                in_zone   = [w for w in line if x_in - 5 <= w["x0"] < x_bal - 5]
                bal_zone  = [w for w in line if (x_bal - 5) <= w["x0"] <= table_right]

                # Update active year if a year marker appears in the Date column
                current_year = _update_year_from_left_zone(left_zone, current_year)

                # Skip pure year marker lines (e.g. a standalone '2025') so they don't pollute descriptions
                if (
                    len(left_zone) == 1
                    and (left_zone[0].get("text") or "").strip().isdigit()
                    and len((left_zone[0].get("text") or "").strip()) == 4
                    and not _line_text(out_zone).strip()
                    and not _line_text(in_zone).strip()
                    and not _line_text(bal_zone).strip()
                    and not _line_text(desc_zone).strip()
                ):
                    continue

                # Update date if present
                possible_date = _try_parse_date(left_zone, current_year)
                if possible_date:
                    current_date = possible_date
                if current_date is None:
                    continue

                # Amounts by column
                out_text = _line_text(out_zone)
                in_text  = _line_text(in_zone)
                bal_text = _line_text(bal_zone)

                out_m = MONEY_RE.search(out_text)
                in_m  = MONEY_RE.search(in_text)
                bal_m = MONEY_RE.search(bal_text)

                out_amt = _to_float(out_m.group(0)) if out_m else None
                in_amt  = _to_float(in_m.group(0)) if in_m else None
                bal_amt = _to_float(bal_m.group(0)) if bal_m else None

                desc_text = _line_text(desc_zone).strip()

                # New transaction row if it has £Out or £In
                if out_amt is not None or in_amt is not None:
                    commit()

                    if in_amt is not None and out_amt is None:
                        signed = +in_amt
                    elif out_amt is not None and in_amt is None:
                        signed = -out_amt
                    else:
                        signed = (in_amt or 0) - (out_amt or 0)

                    base_type, cleaned_first = _extract_type_and_clean_first_line(desc_text)

                    # Build description lines:
                    # - Usually cleaned_first is without the type prefix
                    # - For Returned Direct Debit we intentionally keep the full phrase
                    desc_lines = []
                    if cleaned_first:
                        desc_lines.append(cleaned_first)

                    current_txn = {
                        "date": current_date,
                        "base_type": base_type,
                        "desc_lines": desc_lines,
                        "amount": signed,
                        "balance": bal_amt,
                    }
                    continue

                # Continuation line: add to description
                if current_txn and desc_text:
                    current_txn["desc_lines"].append(desc_text)
                    if current_txn["balance"] is None and bal_amt is not None:
                        current_txn["balance"] = bal_amt

    commit()
    return rows

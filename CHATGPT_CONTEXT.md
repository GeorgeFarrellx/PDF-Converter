# CHATGPT_CONTEXT.md (Authoritative Entry Point)

## Rules (Single Source of Truth)
1. `CHATGPT_CONTEXT.md` is the single source of truth for repository file links used for ChatGPT context.
2. Any new, renamed, or version-bumped file MUST be added here in the same commit.
3. For `Parsers/`, every `*.py` file in the folder MUST be listed as a raw GitHub link.

This file is the single source of truth for ChatGPT context in this repository and must be used as the primary entry point.

## Repo Manifest

Root-level structure and key files/folders visible in this repo:

- `Logs/` — runtime output folder; may not always be complete or committed, but exists in the repository structure.
- `Parsers/` — bank-specific parser modules.
- `CHATGPT_CONTEXT.md` — authoritative context and raw-link index.
- `README.md` — project overview and usage guidance (reference only unless explicitly requested).
- `VERSION.txt` — app/version metadata.
- `core.py` — core conversion/processing orchestration logic.
- `gui.py` — user interface layer and UI behavior.
- `launcher.py` — startup/bootstrap entry wrapper.
- `main.py` — primary app entry point and high-level flow control.
- `gitignore` — ignore rules (reference only).

## Responsibility Map

- `main.py`
  - Primary execution entry and top-level app flow coordination.
  - Delegates processing/UI operations to appropriate modules.

- `core.py`
  - Core conversion pipeline logic and shared processing utilities.
  - Central place for non-UI business logic and orchestration internals.

- `gui.py`
  - GUI construction, layout, and interaction handling.
  - Connects UI actions to core processing functions.

- `launcher.py`
  - Launch/bootstrap wrapper responsible for starting the application in the intended runtime mode.

- `Parsers/*.py`
  - Bank-specific statement parsing implementations.
  - Each parser module handles extraction/normalization rules for its own institution format.

## Repo Structure & Ownership

- **Authoritative index ownership:** `CHATGPT_CONTEXT.md` owns file visibility and raw-link coverage for ChatGPT context.
- **Application orchestration ownership:** `main.py`, `launcher.py`, and `core.py` own app startup, processing flow, and shared non-UI business logic.
- **UI ownership:** `gui.py` owns user interaction, layout, and continuity/reconciliation behavior in the interface.
- **Parser ownership:** `Parsers/*.py` own bank-specific PDF extraction and row normalization only; parser files must not define app-wide GUI or startup policy.
- **Version ownership:** `VERSION.txt` is the authoritative user-facing version marker and must align with parser-version changes.
- **Current app version:** `2.30`.

## Parser Interface Contract

Every parser module in `Parsers/` must expose the same callable contract so the core pipeline can load and execute it predictably.

Required parser functions:

1. `can_parse(...)`  
   - Returns `True/False` for parser applicability to an input statement.
2. `parse_statement(...)`  
   - Performs extraction from source statement/PDF input.
3. `normalize_transactions(...)`  
   - Returns normalized transaction rows in the schema expected by `core.py` and the Excel output path.
4. `get_parser_version(...)`  
   - Returns parser version string matching filename/versioning rules.

Contract rules:

- Parsers must return deterministic normalized output for identical input.
- Parsers must not mutate global application state in `core.py`/`gui.py`.
- Parsers must fail explicitly (structured error/clear exception) when input is unsupported or malformed.
- Parser output must preserve transaction continuity (no silent row drops/reordering unless explicitly documented by parser logic).

## Versioning Rules

- Parser filename version suffixes are authoritative (example: `bankname-1.1.py`).
- Any parser behavior change that can alter normalized output requires a parser version bump.
- Any parser addition, rename, removal, or version bump must update this file in the same commit.
- `VERSION.txt` must be updated when release-level behavior changes are introduced.
- Raw GitHub links in this document must remain intact and updated to match current filenames.

## GUI / Reconciliation Invariants

- GUI must represent the same transaction set produced by the selected parser and `core.py` pipeline (no hidden rows).
- Reconciliation/continuity views must preserve ordering semantics used by the normalized parser output.
- User-triggered re-runs must not mix data from different parser versions within a single reconciliation result.
- UI messaging for parser failures must be explicit and non-silent.
- GUI behavior must not imply successful reconciliation if parsing/normalization failed.

## What Not To Assume

- Do not assume files not listed in this document exist, are readable, or are in scope.
- Do not assume all parser versions are interchangeable; treat version suffixes as compatibility boundaries.
- Do not assume parser output schemas can drift without coordinated `core.py` handling.
- Do not assume GUI state is source-of-truth when it conflicts with parser/core normalized output.
- Do not assume logs are complete, committed, or authoritative for business logic decisions.
- Do not assume undocumented parser side effects are acceptable.

## OWNERSHIP BOUNDARIES

- `main*.py`: orchestration only (no bank-specific parsing logic).
- `core*.py`: reconciliation/continuity/support bundle logic (no bank-specific parsing).
- `gui*.py`: UI only (no parsing/reconciliation logic).
- `Parsers/*.py`: bank-specific parsing only (no categorisation logic, no reconciliation logic).

## PARSER INTERFACE CONTRACT (MANDATORY)

Each parser must expose:
- `extract_transactions(pdf_path) -> list[dict]`
- `extract_statement_balances(pdf_path) -> dict` (start_balance/end_balance or equivalent)
- `extract_account_holder_name(pdf_path) -> str`
- `extract_statement_period(pdf_path) -> (date|None, date|None)` (required where supported)

Rules:
- Text-based PDFs only (NO OCR).
- Must handle multi-page PDFs.
- Must not alter parsed raw transaction values after extraction.
- Must not perform categorisation (categorisation only populates the Category column elsewhere).

## VERSIONING RULES

- If a file is created, renamed, or version-bumped, it MUST be added to `CHATGPT_CONTEXT.md` in the same commit.
- Parser changes: bump filename version if versioned naming is used (e.g., `bank-1.6.py -> bank-1.7.py`).
- Do not delete old versions unless explicitly instructed.
- Keep imports/references consistent when versions change (or remain version-agnostic where already implemented).

## CONTINUITY & RECONCILIATION INVARIANTS

- “Continuity not checked” is an error condition (not a success).
- “Balances not found” must be treated as a detectable issue.
- Statement periods should be displayed in chronological order.
- UI text changes must not change reconciliation/continuity logic.

## FILE LINKS (single source of truth)

### Complete repository file list (all tracked files) + raw links

- `CHATGPT_CONTEXT.md`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/CHATGPT_CONTEXT.md
- `README.md`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/README.md
- `VERSION.txt`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/VERSION.txt
- `core.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/core.py
- `Global Categorisation Rules.csv`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Global%20Categorisation%20Rules.csv
- `gitignore`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gitignore
- `gui.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gui.py
- `launcher.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/launcher.py
- `main.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/main.py
- `Parsers/barclays.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/barclays.py
- `Parsers/halifax.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/halifax.py
- `Parsers/hsbc.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/hsbc.py
- `Parsers/lloyds.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/lloyds.py
- `Parsers/monzo.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/monzo.py
- `Parsers/nationwide.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/nationwide.py
- `Parsers/natwest.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/natwest.py
- `Parsers/rbs.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/rbs.py
- `Parsers/santander.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/santander.py
- `Parsers/starling.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/starling.py
- `Parsers/tsb.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/tsb.py
- `Parsers/zempler.py`  
  https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/zempler.py

### Root app files

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/main.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/core.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gui.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/launcher.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/VERSION.txt
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/CHATGPT_CONTEXT.md

### Parsers

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/barclays.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/halifax.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/hsbc.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/lloyds.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/monzo.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/nationwide.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/natwest.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/rbs.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/santander.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/starling.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/tsb.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/zempler.py

### Optional / Non-editable awareness (reference only)

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/README.md
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gitignore

## Notes / Rules

- Only files listed in this document may be assumed/read.
- If a file is not listed, ChatGPT must respond exactly:
  - “I cannot confirm this from the current GitHub code.”

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

## Raw Links

### Root app files

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/main.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/core.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gui.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/launcher.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/VERSION.txt

### Parsers

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/barclays-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/halifax-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/hsbc-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/lloyds-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/lloyds-1.2.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/monzo-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/nationwide-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/natwest-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/rbs-1.1.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/santander-1.6.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/starling.py
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/Parsers/tsb-1.1.py

### Optional / Non-editable awareness (reference only)

https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/README.md
https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gitignore

## Notes / Rules

- Only files listed in this document may be assumed/read.
- If a file is not listed, ChatGPT must respond exactly:
  - “I cannot confirm this from the current GitHub code.”

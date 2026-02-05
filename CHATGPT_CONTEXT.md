# CHATGPT_CONTEXT.md (Authoritative Entry Point)

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

## Raw Links

### Root app files

- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/main.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/core.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/gui.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/launcher.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/VERSION.txt

### Parsers

- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/barclays.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/halifax.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/hsbc.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/lloyds.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/monzo.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/nationwide.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/natwest.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/rbs.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/santander.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/starling.py
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/main/Parsers/tsb.py

### Optional / Non-editable awareness (reference only)

- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/README.md
- https://raw.githubusercontent.com/GeorgeFarrellx/PDF-Converter/refs/heads/main/gitignore

## Notes / Rules

- Only files listed in this document may be assumed/read.
- If a file is not listed, ChatGPT must respond exactly:
  - “I cannot confirm this from the current GitHub code.”

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

# PDF-Converter
Converts PDFs.

## ChatGPT context links
See [CHATGPT_CONTEXT.md](CHATGPT_CONTEXT.md).

## Parsers
- `Parsers/` contains exactly one active parser per bank named `<bank>.py`.
- Older versions live under `archive/Parsers/<bank>/`.

## Windows (no Python) quick start
1. Download the repository ZIP from GitHub.
2. Extract the ZIP.
3. Open the extracted folder and double-click `SETUP_AND_RUN_WINDOWS.bat`.

The script will try to find Python first (`py -3`, then `python`). If Python is missing, it will attempt to install Python via `winget`. If `winget` is unavailable, you will be prompted to install Python manually. Some locked-down corporate PCs may block `winget` installs.

The script creates `.venv` and installs dependencies automatically before starting the app.

Manual fallback steps:
1. Install Python 3.10 or newer.
2. In the repository folder, run:
   - `python -m venv .venv`
   - `.venv\Scripts\python -m pip install -r requirements.txt`
   - `.venv\Scripts\python main.py`

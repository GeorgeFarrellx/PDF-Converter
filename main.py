# Version: 2.02
import os
import sys
import traceback
from datetime import datetime

import core
from gui import App

# Re-export the core symbols main relies on
PARSERS_DIR = core.PARSERS_DIR
LOGS_DIR = core.LOGS_DIR
ensure_folder = core.ensure_folder
_run_self_tests = core._run_self_tests


def main():
    if not os.path.isdir(PARSERS_DIR):
        raise FileNotFoundError(f"Missing Parsers folder: {PARSERS_DIR}")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _run_self_tests()
        raise SystemExit(0)

    try:
        main()
    except Exception as e:
        try:
            ensure_folder(LOGS_DIR)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            crash_path = os.path.join(LOGS_DIR, f"startup_crash_{ts}.txt")
            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            with open(crash_path, "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        raise

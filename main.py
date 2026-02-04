# Version: 2.02
import os
import sys
import traceback
from datetime import datetime

import importlib


def _show_startup_error(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Startup Error", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)


def check_dependencies() -> None:
    missing = []
    installable = []

    if sys.version_info < (3, 10):
        missing.append("Python 3.10+ is required.")

    for module, package in [
        ("tkinterdnd2", "tkinterdnd2"),
        ("pandas", "pandas"),
        ("openpyxl", "openpyxl"),
    ]:
        try:
            importlib.import_module(module)
        except Exception:
            missing.append(f"Missing dependency: {package}")
            installable.append(package)

    if missing:
        message = "The application cannot start because required dependencies are missing:\n\n"
        message += "\n".join(f"- {item}" for item in missing)
        if installable:
            message += "\n\nInstall them with:\n  python -m pip install " + " ".join(installable)
        _show_startup_error(message)
        raise SystemExit(1)


def main():
    check_dependencies()

    import core
    from gui import App

    if not os.path.isdir(core.PARSERS_DIR):
        raise FileNotFoundError(f"Missing Parsers folder: {core.PARSERS_DIR}")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        check_dependencies()
        import core
        core._run_self_tests()
        raise SystemExit(0)

    try:
        main()
    except Exception as e:
        try:
            import core

            core.ensure_folder(core.LOGS_DIR)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            crash_path = os.path.join(core.LOGS_DIR, f"startup_crash_{ts}.txt")
            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            with open(crash_path, "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        raise

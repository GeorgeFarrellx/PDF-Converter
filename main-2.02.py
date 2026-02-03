# Version: main-2.02
import os
import sys
import glob
import re
import traceback
import importlib
import importlib.util
from datetime import datetime


def _import_latest_core_module():
    """Import core from the only core_*.py file present (version-agnostic)."""
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = glob.glob(os.path.join(here, "core_*.py"))
    if not candidates:
        raise ModuleNotFoundError(f"No core_*.py found in {here}")

    def ver_key(path: str):
        base = os.path.splitext(os.path.basename(path))[0]  # e.g. core_1_0
        m = re.match(r"^core_([0-9]+)(?:_([0-9]+))?(?:_([0-9]+))?$", base)
        if not m:
            return (-1, -1, -1)
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) is not None else 0
        c = int(m.group(3)) if m.group(3) is not None else 0
        return (a, b, c)

    best_path = sorted(candidates, key=ver_key)[-1]
    mod_name = os.path.splitext(os.path.basename(best_path))[0]

    try:
        module = importlib.import_module(mod_name)
    except ModuleNotFoundError:
        spec = importlib.util.spec_from_file_location(mod_name, best_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Unable to load module spec for {best_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[mod_name] = module
        spec.loader.exec_module(module)

    return module


_core = _import_latest_core_module()

# Re-export the core symbols main relies on
PARSERS_DIR = _core.PARSERS_DIR
LOGS_DIR = _core.LOGS_DIR
ensure_folder = _core.ensure_folder
_run_self_tests = _core._run_self_tests


def _import_latest_gui_app():
    """Import App from the only gui_*.py file present (version-agnostic)."""
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = glob.glob(os.path.join(here, "gui_*.py"))
    if not candidates:
        raise ModuleNotFoundError(f"No gui_*.py found in {here}")

    def ver_key(path: str):
        base = os.path.splitext(os.path.basename(path))[0]  # e.g. gui_1_1
        m = re.match(r"^gui_([0-9]+)(?:_([0-9]+))?(?:_([0-9]+))?$", base)
        if not m:
            return (-1, -1, -1)
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) is not None else 0
        c = int(m.group(3)) if m.group(3) is not None else 0
        return (a, b, c)

    best_path = sorted(candidates, key=ver_key)[-1]
    mod_name = os.path.splitext(os.path.basename(best_path))[0]

    # Try normal import first (works when folder is on sys.path)
    try:
        module = importlib.import_module(mod_name)
    except ModuleNotFoundError:
        # Fallback: import directly from file path
        spec = importlib.util.spec_from_file_location(mod_name, best_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Unable to load module spec for {best_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[mod_name] = module
        spec.loader.exec_module(module)

    if not hasattr(module, "App"):
        raise ImportError(f"{mod_name}.py does not define App")

    return module.App


App = _import_latest_gui_app()


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

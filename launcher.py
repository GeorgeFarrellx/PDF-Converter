# Version: 2.02
import os
import sys
import traceback
from datetime import datetime


def ensure_folder(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def write_startup_log(logs_dir: str, prefix: str, content: str) -> str:
    ensure_folder(logs_dir)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(logs_dir, f"{prefix}_{ts}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return path


def main():
    # When frozen, sys.executable is the EXE path.
    if getattr(sys, "frozen", False):
        app_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))

    logs_dir = os.path.join(app_dir, "Logs")

    try:
        os.chdir(app_dir)
    except Exception:
        pass

    # Ensure imports can find sibling modules like core.py / gui.py
    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)

    target = os.path.abspath(os.path.join(app_dir, "main.py"))

    header = []
    header.append(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    header.append(f"App dir: {app_dir}")
    header.append(f"Target main: {target}")
    header.append("")
    header_txt = "\n".join(header)

    if not os.path.exists(target):
        log_path = write_startup_log(logs_dir, "startup_missing_main", header_txt + "ERROR: Target main not found.\n")
        raise FileNotFoundError(f"Target main not found. Log: {log_path}")

    try:
        with open(target, "r", encoding="utf-8") as f:
            src = f.read()
        code = compile(src, target, "exec")
    except SyntaxError as e:
        details = header_txt
        details += "SYNTAX ERROR:\n"
        details += f"{e.__class__.__name__}: {e}\n"
        details += f"File: {e.filename}\nLine: {e.lineno}\nOffset: {e.offset}\n"
        details += "\nText:\n"
        details += (e.text or "").rstrip("\n") + "\n\n"
        details += "Traceback:\n" + traceback.format_exc()
        log_path = write_startup_log(logs_dir, "startup_syntax_error", details)
        raise SyntaxError(f"Syntax error in {target}. See log: {log_path}") from e
    except Exception as e:
        details = header_txt + "ERROR compiling target:\n" + traceback.format_exc()
        log_path = write_startup_log(logs_dir, "startup_compile_error", details)
        raise RuntimeError(f"Compile error. See log: {log_path}") from e

    try:
        glb = {"__file__": target, "__name__": "__main__", "__package__": None}
        exec(code, glb, glb)
    except SystemExit:
        raise
    except Exception as e:
        details = header_txt + "RUNTIME ERROR:\n" + traceback.format_exc()
        log_path = write_startup_log(logs_dir, "startup_runtime_error", details)
        raise RuntimeError(f"Runtime error. See log: {log_path}") from e


if __name__ == "__main__":
    main()

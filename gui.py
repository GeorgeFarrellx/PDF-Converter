# Version: 2.03
import os
import traceback
import zipfile
from datetime import datetime, timedelta, date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Drag & drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception as e:
    raise RuntimeError(
        "tkinterdnd2 is not installed. Install it with:\n"
        "  python -m pip install tkinterdnd2\n\n"
        f"Original error: {e}"
    )

from core import *  # noqa: F403


def _read_app_version() -> str:
    version_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "VERSION.txt")
    try:
        with open(version_path, "r", encoding="utf-8") as f:
            return (f.read() or "").strip()
    except Exception:
        return ""


APP_VERSION = _read_app_version()



def _fmt_money(v) -> str:
    """Safely format a numeric value as GBP for logs/UI.

    Returns an empty string for None/blank. Never raises.
    """
    try:
        if v is None or v == "":
            return ""
        if isinstance(v, str):
            s = v.strip()
            if not s:
                return ""
            s = s.replace("£", "").replace(",", "")
            v = float(s)
        else:
            v = float(v)

        sign = "-" if v < 0 else ""
        v = abs(v)
        return f"{sign}£{v:,.2f}"
    except Exception:
        try:
            return str(v) if v is not None else ""
        except Exception:
            return ""


# ----------------------------
# GUI
# ----------------------------


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF Converter")
        self.geometry("760x520")
        self.minsize(720, 500)

        self.selected_files: list[str] = []
        self.bank_var = tk.StringVar(value="Select bank...")
        self.output_folder_var = tk.StringVar(value=DEFAULT_OUTPUT_FOLDER)
        self.status_var = tk.StringVar(value="Ready.")
        self.auto_detect_var = tk.BooleanVar(value=True)

        self.last_report_data = None
        self.last_excel_data = None
        self.last_saved_output_path = None

        self._build_ui()
        self._wire_dnd()

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        bank_row = ttk.Frame(root)
        bank_row.pack(fill="x")

        ttk.Label(bank_row, text="Bank:").pack(side="left")
        self.bank_combo = ttk.Combobox(
            bank_row,
            textvariable=self.bank_var,
            values=BANK_OPTIONS,
            state="readonly",
            width=20,
        )
        self.bank_combo.pack(side="left", padx=(8, 16))

        ttk.Checkbutton(
            bank_row,
            text="Auto-detect bank",
            variable=self.auto_detect_var,
        ).pack(side="left")

        ttk.Label(root, text="Drag & drop PDF statements here:").pack(anchor="w", pady=(14, 6))

        self.drop_box = tk.Listbox(root, height=10)
        self.drop_box.pack(fill="both", expand=False)
        self.drop_box.insert("end", "Drop PDFs here, or click 'Browse PDFs'.")

        btn_row = ttk.Frame(root)
        btn_row.pack(fill="x", pady=(10, 0))

        ttk.Button(btn_row, text="Browse PDFs", command=self.browse_pdfs).pack(side="left")
        ttk.Button(btn_row, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(btn_row, text="Clear List", command=self.clear_list).pack(side="left", padx=8)

        ttk.Separator(root, orient="horizontal").pack(fill="x", pady=(10, 10))

        run_row = ttk.Frame(root)
        run_row.pack(fill="x", pady=(12, 0))

        self.run_btn = ttk.Button(run_row, text="Convert", command=self.run_parser)
        self.run_btn.pack(side="left")

        self.cleanup_btn = ttk.Button(run_row, text="Clean Up", command=self.clean_up)
        self.cleanup_btn.pack(side="left", padx=10)

        ttk.Separator(root).pack(fill="x", pady=(18, 8))
        ttk.Label(root, text="Status:").pack(anchor="w")
        self.progress = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(4, 2))
        self.status_label = ttk.Label(root, textvariable=self.status_var)
        self.status_label.pack(anchor="w", pady=(2, 0))

        post_row = ttk.Frame(root)
        post_row.pack(fill="x", pady=(12, 0))

        self.show_checks_btn = ttk.Button(post_row, text="Show Checks", command=self.show_last_checks)
        self.show_checks_btn.pack(side="left")

        self.save_again_btn = ttk.Button(post_row, text="Save Output", command=self.save_last_output)
        self.save_again_btn.pack(side="left", padx=10)

    def _wire_dnd(self):
        self.drop_box.drop_target_register(DND_FILES)
        self.drop_box.dnd_bind("<<Drop>>", self.on_drop)

    def set_status(self, msg: str):
        self.status_var.set(msg)
        try:
            self.update_idletasks()
        except Exception:
            pass

    def set_progress(self, completed: int, total: int):
        """Update progress bar based on completed PDFs / total PDFs."""
        try:
            total_i = int(total) if total not in (None, "") else 0
        except Exception:
            total_i = 0
        if total_i <= 0:
            total_i = 1
        try:
            completed_i = int(completed) if completed not in (None, "") else 0
        except Exception:
            completed_i = 0

        pct = (max(0, min(completed_i, total_i)) / total_i) * 100.0

        try:
            if hasattr(self, "progress") and self.progress is not None:
                self.progress["value"] = max(0.0, min(100.0, float(pct)))
        except Exception:
            pass

        try:
            self.update_idletasks()
        except Exception:
            pass

    def browse_pdfs(self):
        filepaths = filedialog.askopenfilenames(
            title="Select PDF bank statements",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not filepaths:
            return
        self.add_files(list(filepaths))

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return
        self.output_folder_var.set(folder)

    def show_last_checks(self):
        if not self.last_report_data:
            messagebox.showwarning("Checks", "No checks to show yet. Convert first.")
            return

        data = self.last_report_data
        output_path = self.last_saved_output_path or "(Not saved yet)"
        pre_save = False if self.last_saved_output_path else True

        show_reconciliation_popup(
            self,
            output_path,
            data.get("recon_results") or [],
            coverage_period=data.get("coverage_period", "") or "",
            continuity_results=data.get("continuity_results") or [],
            pre_save=pre_save,
            support_bundle_callback=self.create_support_bundle_zip,
        )

    def save_last_output(self):
        if not self.last_excel_data:
            messagebox.showwarning("Save", "Nothing to save yet. Convert first.")
            return

        data = self.last_excel_data
        transactions = data.get("transactions") or []
        client_name = data.get("client_name") or ""
        filename = data.get("filename") or "Transactions.xlsx"

        initial_dir = data.get("initial_dir") or ""
        if not initial_dir:
            folder = self.output_folder_var.get().strip()
            if folder:
                initial_dir = folder

        output_path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialdir=initial_dir or None,
            initialfile=filename,
        )

        if not output_path:
            self.set_status("Cancelled.")
            return

        try:
            ensure_folder(os.path.dirname(output_path))
        except Exception as e:
            messagebox.showerror("Save error", "Cannot create folder for output file:\n" + str(e))
            self.set_status("Failed")
            return

        try:
            self.output_folder_var.set(os.path.dirname(output_path))
        except Exception:
            pass

        self.set_status("Writing Excel...")
        save_transactions_to_excel(transactions, output_path, client_name=client_name)

        self.last_saved_output_path = output_path
        self.set_progress(len(self.selected_files), max(1, len(self.selected_files)))

        # Determine whether the last run had warnings, so the status mirrors the main run result.
        any_warn = False
        try:
            any_warn = bool((self.last_report_data or {}).get("any_warn"))
        except Exception:
            any_warn = False

        if any_warn:
            self.set_status(f"Done with warnings. Output: {output_path}")
        else:
            self.set_status(f"Done. Output: {output_path}")

    def create_support_bundle_zip(self):
        if not self.last_report_data:
            messagebox.showwarning("Support bundle", "No run data available. Run the parser first.")
            return

        data = self.last_report_data or {}

        # Support bundle gating: treat any warnings/failures/NOT CHECKED/balances missing as issues.
        def _status_is_ok(s) -> bool:
            try:
                s = str(s or "").strip().upper()
            except Exception:
                return False
            return s.startswith("OK")

        def _text_has_issue_markers(s) -> bool:
            try:
                u = str(s or "").upper()
            except Exception:
                return False
            return (
                ("NOT CHECKED" in u)
                or ("BALANCES NOT FOUND" in u)
                or ("BALANCE NOT FOUND" in u)
                or ("MISMATCH" in u)
                or ("FAILED" in u)
            )

        def _missing_balances(rec: dict) -> bool:
            try:
                sb = rec.get("start_balance")
                eb = rec.get("end_balance")
                return (sb is None or sb == "") or (eb is None or eb == "")
            except Exception:
                return True

        has_issue = False
        issue_reasons = []

        # Fast path: if GUI already flagged warnings, allow bundle.
        try:
            has_issue = bool(data.get("any_warn"))
        except Exception:
            has_issue = False

        # Recon results: anything other than OK is an issue.
        if not has_issue:
            try:
                for r in (data.get("recon_results") or []):
                    st = (r.get("status") if isinstance(r, dict) else str(r or ""))
                    if (not _status_is_ok(st)) or _text_has_issue_markers(st):
                        has_issue = True
                        break
            except Exception:
                pass

        # Continuity results: NOT CHECKED is an issue (treat as error).
        if not has_issue:
            try:
                for c in (data.get("continuity_results") or []):
                    if isinstance(c, dict):
                        st = (c.get("display_status") or c.get("status") or "")
                        prev_pdf = c.get("prev_pdf") or ""
                        next_pdf = c.get("next_pdf") or ""
                    else:
                        st = str(c or "")
                        prev_pdf = ""
                        next_pdf = ""

                    if _text_has_issue_markers(st) or (not _status_is_ok(st)):
                        has_issue = True
                        if "NOT CHECKED" in str(st).upper():
                            issue_reasons.append(f"Continuity not checked: {prev_pdf} -> {next_pdf}")
                        break
            except Exception:
                pass

        # Explicit balance-missing gate: if any PDF lacks start/end balances, continuity may show "Not checked".
        # Treat this as an issue so support bundles include logs.
        try:
            for r in (data.get("recon_results") or []):
                if not isinstance(r, dict):
                    continue
                if _missing_balances(r):
                    has_issue = True
                    issue_reasons.append(f"{r.get('pdf') or 'PDF'}: balances not found")
        except Exception:
            pass

        if not has_issue:
            # Still allow bundle for verification/debugging, but wording is safe (no broken string literals).
            messagebox.showinfo(
                "Support bundle",
                "No reconciliation or continuity issues were detected for the last run based on stored results.\n"
                "You can still create a support bundle for verification/debugging.",
            )

        if not self.last_excel_data:
            messagebox.showerror(
                "Support bundle",
                "Cannot create support bundle because the Excel data for the last run is missing.",
            )
            return

        try:
            ensure_folder(LOGS_DIR)
        except Exception as e:
            messagebox.showerror("Support bundle", "Cannot access Logs folder:\n" + str(e))
            return

        excel_source = (data.get("output_xlsx_path") or self.last_saved_output_path or "")
        excel_source = excel_source if (excel_source and os.path.exists(excel_source)) else ""

        bundle_base = data.get("bundle_base")
        if not bundle_base:
            try:
                bundle_base = os.path.splitext(self.last_excel_data.get("filename") or "Transactions.xlsx")[0]
            except Exception:
                bundle_base = "RUN"

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"SUPPORT BUNDLE - {bundle_base} - {ts}.zip"
        zip_path = make_unique_path(os.path.join(LOGS_DIR, zip_name))

        transactions = self.last_excel_data.get("transactions") or []
        client_name = self.last_excel_data.get("client_name") or ""

        temp_excel_path = ""
        created_temp_excel = False
        zip_created = False

        try:
            if not excel_source:
                temp_excel_name = f"SUPPORT EXCEL - {bundle_base} - {ts}.xlsx"
                temp_excel_path = make_unique_path(os.path.join(LOGS_DIR, temp_excel_name))
                save_transactions_to_excel(transactions, temp_excel_path, client_name=client_name)
                excel_source = temp_excel_path
                created_temp_excel = True

            log_path = data.get("log_path") or ""
            log_exists = bool(log_path and os.path.exists(log_path))

            # If no log exists (common when continuity is NOT CHECKED), create a lightweight support log now.
            if not log_exists:
                support_log_name = f"{bundle_base} - support log - {ts}.txt"
                support_log_path = make_unique_path(os.path.join(LOGS_DIR, support_log_name))

                recon = data.get("recon_results") or []
                cont = data.get("continuity_results") or []

                lines = []
                lines.append(f"Support log generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                lines.append(f"Bank: {(data.get('bank') or '').strip()}")
                lines.append(f"GUI: {APP_VERSION}")
                lines.append(f"Core: {APP_VERSION}")
                lines.append("")

                if issue_reasons:
                    lines.append("Detected issues:")
                    for r in issue_reasons:
                        lines.append(f"- {r}")
                    lines.append("")

                lines.append("Reconciliation summary:")
                for r in recon:
                    if not isinstance(r, dict):
                        lines.append(f"- {r}")
                        continue
                    pdf = r.get("pdf") or ""
                    st = r.get("status") or ""
                    sb = _fmt_money(r.get("start_balance"))
                    eb = _fmt_money(r.get("end_balance"))
                    ps = r.get("period_start")
                    pe = r.get("period_end")
                    per = ""
                    if ps and pe and hasattr(ps, "strftime") and hasattr(pe, "strftime"):
                        per = f"{ps.strftime('%d/%m/%Y')} - {pe.strftime('%d/%m/%Y')}"
                    lines.append(
                        f"- {pdf}: {st} | start={sb or '<missing>'} | end={eb or '<missing>'} | period={per or 'None'}"
                    )
                lines.append("")

                lines.append("Continuity summary:")
                for c in cont:
                    if isinstance(c, dict):
                        prev_pdf = c.get("prev_pdf") or ""
                        next_pdf = c.get("next_pdf") or ""
                        st = c.get("display_status") or c.get("status") or ""
                        lines.append(f"- {prev_pdf} -> {next_pdf}: {st}")
                    else:
                        lines.append(f"- {c}")
                lines.append("")

                with open(support_log_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(lines).rstrip() + "\n")

                log_path = support_log_path
                log_exists = True
                try:
                    if self.last_report_data is not None:
                        self.last_report_data["log_path"] = log_path
                except Exception:
                    pass

            learning_report_path = data.get("learning_report_path") or ""
            learning_report_exists = bool(learning_report_path and os.path.exists(learning_report_path))

            pdf_paths = list(data.get("source_pdfs") or [])

            used_names = set()

            def _unique_zip_name(filename: str) -> str:
                base, ext = os.path.splitext(filename)
                if not ext:
                    ext = ".pdf"
                candidate = f"{base}{ext}"
                if candidate.lower() not in used_names:
                    used_names.add(candidate.lower())
                    return candidate
                n = 2
                while True:
                    candidate = f"{base} ({n}){ext}"
                    if candidate.lower() not in used_names:
                        used_names.add(candidate.lower())
                        return candidate
                    n += 1

            snapshot_files = []
            missing_snapshot = []

            gui_path = os.path.abspath(__file__)
            if os.path.exists(gui_path):
                snapshot_files.append((gui_path, "CODE_SNAPSHOT/gui.py"))
            else:
                missing_snapshot.append("gui.py")

            core_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "core.py")
            if os.path.exists(core_path):
                snapshot_files.append((core_path, "CODE_SNAPSHOT/core.py"))
            else:
                missing_snapshot.append("core.py")

            parser_path = data.get("parser_file") or ""
            parser_basename = os.path.basename(parser_path) if parser_path else ""
            if parser_path and os.path.exists(parser_path):
                snapshot_files.append((parser_path, f"CODE_SNAPSHOT/{parser_basename}"))
            else:
                missing_snapshot.append(parser_basename or "parser file")

            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(excel_source, arcname=os.path.basename(excel_source))

                if log_exists:
                    zf.write(log_path, arcname=os.path.basename(log_path))

                if learning_report_exists:
                    zf.write(learning_report_path, arcname=os.path.basename(learning_report_path))

                for p in pdf_paths:
                    if not p or not os.path.exists(p):
                        continue
                    base = os.path.basename(p) or "statement.pdf"
                    safe = sanitize_filename(base) or "statement.pdf"
                    unique = _unique_zip_name(safe)
                    arcname = "Source PDFs/" + unique
                    zf.write(p, arcname=arcname)

                for local_path, arcname in snapshot_files:
                    zf.write(local_path, arcname=arcname)

                if missing_snapshot:
                    zf.writestr(
                        "CODE_SNAPSHOT/_MISSING_FILES.txt",
                        "\n".join(sorted(set(missing_snapshot))) + "\n",
                    )

            zip_created = True

            if created_temp_excel and temp_excel_path and os.path.exists(temp_excel_path):
                try:
                    os.remove(temp_excel_path)
                except Exception:
                    pass

            messagebox.showinfo("Support bundle created", "Support bundle created:\n" + zip_path)

        except Exception as e:
            if created_temp_excel and temp_excel_path and zip_created and os.path.exists(temp_excel_path):
                try:
                    os.remove(temp_excel_path)
                except Exception:
                    pass

            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            messagebox.showerror("Support bundle error", f"{e}\n\nDetails:\n{err}")


    def generate_learning_report(self, reason: str | None = None, exception: Exception | None = None):
        if pdfplumber is None:
            return None

        try:
            ensure_folder(LOGS_DIR)
        except Exception:
            return None

        data = self.last_report_data or {}
        bank = (data.get("bank") or self.bank_var.get() or "").strip() or "Unknown"
        client_or_run = (
            data.get("bundle_base")
            or data.get("client_name")
            or os.path.splitext(data.get("run_filename") or "")[0]
            or "RUN"
        )
        client_or_run = sanitize_filename(str(client_or_run)) or "RUN"

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_name = f"LEARNING - {ts} - {sanitize_filename(bank)} - {client_or_run}.txt"
        report_path = make_unique_path(os.path.join(LOGS_DIR, report_name))

        source_pdfs = list(data.get("source_pdfs") or [])

        main_id = APP_VERSION

        autodetect_result = data.get("autodetect_first_pdf")
        parser_file = data.get("parser_file") or ""

        recon_results = data.get("recon_results") or []
        continuity_results = data.get("continuity_results") or []

        total_tx = 0
        date_min = None
        date_max = None
        try:
            if self.last_excel_data:
                txs = self.last_excel_data.get("transactions") or []
                total_tx = len(txs)
                dates = [t.get("Date") for t in txs if isinstance(t, dict) and t.get("Date")]
                if dates:
                    date_min = min(dates)
                    date_max = max(dates)
        except Exception:
            total_tx = 0
            date_min = None
            date_max = None

        def _norm_text_block(s: str) -> str:
            s = (s or "").replace("\r\n", "\n").replace("\r", "\n")
            lines = [ln.strip() for ln in s.split("\n")]
            out = []
            blank_run = 0
            for ln in lines:
                if not ln:
                    blank_run += 1
                    if blank_run <= 2:
                        out.append("")
                    continue
                blank_run = 0
                out.append(ln)
            return "\n".join(out).strip()

        def _page_snapshot(text: str) -> str:
            text = _norm_text_block(text)
            if not text:
                return ""
            lines = text.splitlines()
            if len(lines) > 80:
                snap = "\n".join(lines[:80])
            else:
                snap = text
            if len(snap) > 3000:
                snap = snap[:3000] + "\n...<truncated>"
            return snap

        def _fmt_date(v):
            try:
                if v is None or v == "":
                    return ""
                if hasattr(v, "to_pydatetime"):
                    v = v.to_pydatetime()
                if hasattr(v, "date") and isinstance(v, datetime):
                    v = v.date()
                if hasattr(v, "strftime"):
                    return v.strftime("%d/%m/%Y")
                return str(v)
            except Exception:
                return str(v)

        try:
            with open(report_path, "w", encoding="utf-8") as f:
                f.write("LEARNING REPORT\n")
                f.write("=" * 60 + "\n")
                f.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                if main_id:
                    f.write(f"Main: {main_id}\n")
                f.write(f"Bank (selected): {bank}\n")
                if autodetect_result:
                    f.write(f"Auto-detect (first PDF): {autodetect_result}\n")
                if parser_file:
                    f.write(f"Parser file: {parser_file}\n")
                if reason:
                    f.write(f"Reason: {reason}\n")
                f.write("\n")

                f.write("PDF SUMMARY\n")
                f.write("-" * 60 + "\n")

                for pdf_path in source_pdfs:
                    f.write(f"PDF: {os.path.basename(pdf_path)}\n")
                    f.write(f"Path: {pdf_path}\n")

                    page_count = 0
                    empty_pages = 0
                    per_page_text = []

                    try:
                        with pdfplumber.open(pdf_path) as pdf:
                            page_count = len(pdf.pages)
                            for pi, page in enumerate(pdf.pages, start=1):
                                txt = ""
                                try:
                                    txt = page.extract_text() or ""
                                except Exception:
                                    txt = ""

                                snap_base = _norm_text_block(txt)
                                if len(snap_base) < 50:
                                    empty_pages += 1
                                per_page_text.append((pi, txt))

                    except Exception as e:
                        f.write(f"Page count: (error: {e})\n\n")
                        continue

                    f.write(f"Page count: {page_count}\n")
                    mostly_empty = (page_count > 0 and (empty_pages / page_count) >= 0.7)
                    f.write(
                        f"Extracted text mostly empty: {'YES' if mostly_empty else 'NO'} ({empty_pages}/{page_count} pages low-text)\n"
                    )
                    f.write("\n")

                    f.write("Per-page text snapshot:\n")
                    for (pi, txt) in per_page_text:
                        snap = _page_snapshot(txt)
                        f.write("\n")
                        f.write(f"--- Page {pi} ---\n")
                        if snap:
                            f.write(snap + "\n")
                        else:
                            f.write("<no extracted text>\n")

                    f.write("\n" + ("-" * 60) + "\n\n")

                f.write("RUN SUMMARY\n")
                f.write("-" * 60 + "\n")
                f.write(f"Total transactions: {total_tx}\n")
                if date_min and date_max:
                    f.write(f"Date range: {_fmt_date(date_min)} - {_fmt_date(date_max)}\n")
                f.write("\n")

                f.write("Statement balances found per PDF:\n")
                for r in recon_results:
                    pdf = r.get("pdf") or ""
                    sb = r.get("start_balance")
                    eb = r.get("end_balance")
                    sb_ok = "YES" if sb is not None and sb != "" else "NO"
                    eb_ok = "YES" if eb is not None and eb != "" else "NO"
                    f.write(f"- {pdf}: start_found={sb_ok}, end_found={eb_ok}\n")
                f.write("\n")

                f.write("Reconciliation results:\n")
                for r in recon_results:
                    pdf = r.get("pdf") or ""
                    st = r.get("status") or ""
                    diff = r.get("difference")
                    f.write(f"- {pdf}: {st}")
                    if diff is not None and diff != "":
                        try:
                            f.write(f" (diff {float(diff):.2f})")
                        except Exception:
                            f.write(f" (diff {diff})")
                    f.write("\n")
                f.write("\n")

                if continuity_results:
                    f.write("Continuity results:\n")
                    for c in continuity_results:
                        prev_pdf = c.get("prev_pdf") or ""
                        next_pdf = c.get("next_pdf") or ""
                        st = c.get("status") or ""
                        diff = c.get("diff")
                        missing = ""
                        try:
                            mf = c.get("missing_from")
                            mt = c.get("missing_to")
                            if mf and mt and hasattr(mf, "strftime") and hasattr(mt, "strftime"):
                                missing = f" | missing {_fmt_date(mf)} - {_fmt_date(mt)}"
                        except Exception:
                            missing = ""

                        f.write(f"- {prev_pdf} -> {next_pdf}: {st}")
                        if diff is not None and diff != "":
                            try:
                                f.write(f" (diff {float(diff):.2f})")
                            except Exception:
                                f.write(f" (diff {diff})")
                        if missing:
                            f.write(missing)
                        f.write("\n")
                    f.write("\n")

                if exception is not None:
                    f.write("EXCEPTION\n")
                    f.write("-" * 60 + "\n")
                    f.write(f"Type: {type(exception).__name__}\n")
                    f.write(f"Message: {exception}\n\n")
                    tb = "".join(traceback.format_exception(type(exception), exception, exception.__traceback__))
                    f.write(tb)
                    f.write("\n")

            try:
                if self.last_report_data is None:
                    self.last_report_data = {}
                self.last_report_data["learning_report_path"] = report_path
            except Exception:
                pass

            return report_path

        except Exception:
            return None

    def clear_list(self):
        self.selected_files = []
        self.drop_box.delete(0, "end")
        self.drop_box.insert(0, "Drop PDFs here, or click 'Browse PDFs'.")
        self.set_status("Cleared file list.")

    def remove_selected(self):
        selected = list(self.drop_box.curselection())
        if not selected:
            return

        displayed = list(self.drop_box.get(0, "end"))

        if len(displayed) == 1 and displayed[0].startswith("Drop PDFs here"):
            return

        for idx in sorted(selected, reverse=True):
            if 0 <= idx < len(self.selected_files):
                self.selected_files.pop(idx)

        self.drop_box.delete(0, "end")
        if not self.selected_files:
            self.drop_box.insert("end", "Drop PDFs here, or click 'Browse PDFs'.")
        else:
            for p in self.selected_files:
                self.drop_box.insert("end", os.path.basename(p))

        self.set_status("Removed selected item(s).")

    def on_drop(self, event):
        files = parse_dnd_event_files(event.data)
        self.add_files(files)

    def add_files(self, files: list[str]):
        pdfs = []
        for f in files:
            f = f.strip()
            if not f:
                continue
            if not os.path.exists(f):
                continue
            if not is_pdf(f):
                continue
            pdfs.append(f)

        if not pdfs:
            self.set_status("No valid PDFs added.")
            return

        if self.auto_detect_var.get():
            detected = auto_detect_bank_from_pdf(pdfs[0])
            if detected and detected in BANK_OPTIONS:
                self.bank_var.set(detected)

        for p in pdfs:
            if p not in self.selected_files:
                self.selected_files.append(p)

        self.drop_box.delete(0, "end")
        for p in self.selected_files:
            self.drop_box.insert("end", os.path.basename(p))

        self.set_status(f"Added {len(pdfs)} PDF(s). Total: {len(self.selected_files)}.")

    def clean_up(self):
        if not self.selected_files:
            messagebox.showwarning("No files", "Please add at least one PDF statement.")
            return

        bank = self.bank_var.get().strip()
        out_folder = self.output_folder_var.get().strip()

        if not bank or bank == "Select bank...":
            messagebox.showwarning(
                "Bank",
                "Please select a bank (or enable auto-detect and add a PDF).",
            )
            return

        try:
            self.set_status(f"Loading parser for {bank}...")
            self.set_status("Starting...")
            self.set_progress(0, max(1, len(self.selected_files)))
            parser = load_parser_module(bank)

            client_name = ""
            try:
                if hasattr(parser, "extract_account_holder_name"):
                    client_name = parser.extract_account_holder_name(self.selected_files[0]) or ""
            except Exception:
                client_name = ""
            if not client_name:
                client_name = get_client_name_from_pdf(self.selected_files[0])

            client_folder = sanitize_filename((client_name or "").strip().upper()) or "CLIENT"

            initial_dir = ""
            if out_folder:
                initial_dir = out_folder
            else:
                try:
                    initial_dir = os.path.dirname(self.selected_files[0])
                except Exception:
                    initial_dir = ""

            zip_path = filedialog.asksaveasfilename(
                title="Save ZIP file",
                defaultextension=".zip",
                filetypes=[("ZIP file", "*.zip")],
                initialdir=initial_dir or None,
                initialfile=f"{client_folder}.zip",
            )

            if not zip_path:
                self.set_status("Clean Up cancelled.")
                return

            try:
                ensure_folder(os.path.dirname(zip_path))
            except Exception as e:
                messagebox.showerror("Save error", f"Cannot create folder for ZIP file:\n{e}")
                self.set_status("Error.")
                return

            try:
                self.output_folder_var.set(os.path.dirname(zip_path))
            except Exception:
                pass

            items = []
            failures = []

            for i, pdf_path in enumerate(self.selected_files, start=1):
                self.set_status(
                    f"Reading statement dates {i}/{len(self.selected_files)}: {os.path.basename(pdf_path)}"
                )

                dmin = None
                dmax = None

                if bank == "Barclays":
                    try:
                        dmin, dmax = extract_barclays_statement_period(pdf_path)
                    except Exception:
                        dmin, dmax = None, None

                if not (dmin and dmax):
                    try:
                        txns = parser.extract_transactions(pdf_path) or []
                        _dates = [t.get("Date") for t in (txns or []) if t.get("Date")]
                        dmin = min(_dates) if _dates else None
                        dmax = max(_dates) if _dates else None
                    except Exception:
                        failures.append(os.path.basename(pdf_path))
                        dmin, dmax = None, None

                period = ""
                try:
                    if dmin and dmax:
                        period = f"{dmin.strftime('%d.%m.%y')} - {dmax.strftime('%d.%m.%y')}"
                except Exception:
                    period = ""

                items.append(
                    {
                        "path": pdf_path,
                        "pdf": os.path.basename(pdf_path),
                        "date_min": dmin,
                        "date_max": dmax,
                        "period": period,
                    }
                )

            def _sort_key(it):
                d = it.get("date_min")
                try:
                    if hasattr(d, "to_pydatetime"):
                        d = d.to_pydatetime()
                except Exception:
                    pass
                try:
                    if isinstance(d, datetime):
                        d = d.date()
                except Exception:
                    pass

                if d is None:
                    return (1, datetime.max.date(), it.get("pdf", ""))
                return (0, d, it.get("pdf", ""))

            items.sort(key=_sort_key)

            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for idx, it in enumerate(items, start=1):
                    period = it.get("period") or ""
                    if period:
                        arc_base = f"{idx} {period}.pdf"
                    else:
                        base = sanitize_filename(os.path.splitext(it.get("pdf", "statement"))[0]) or "statement"
                        arc_base = f"{idx} {base}.pdf"

                    arcname = arc_base
                    zf.write(it["path"], arcname=arcname)

            self.set_status(f"Clean Up complete: {zip_path}")

            if failures:
                messagebox.showwarning(
                    "Clean Up complete (some periods unknown)",
                    "ZIP created successfully, but I could not extract dates for some PDFs.\n"
                    "They were included using the original filename as a fallback:\n\n"
                    + "\n".join(failures)
                    + f"\n\nZIP: {zip_path}",
                )
            else:
                messagebox.showinfo("Clean Up complete", f"ZIP created:\n{zip_path}")

        except Exception as e:
            self.set_status("Error.")
            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            messagebox.showerror("Clean Up error", f"{e}\n\nDetails:\n{err}")

    # (run_parser unchanged)
    def run_parser(self):
        if not self.selected_files:
            messagebox.showwarning("No files", "Please add at least one PDF statement.")
            return

        bank = self.bank_var.get().strip()
        out_folder = self.output_folder_var.get().strip()

        if not bank or bank == "Select bank...":
            messagebox.showwarning(
                "Bank",
                "Please select a bank (or enable auto-detect and add a PDF).",
            )
            return
        if out_folder:
            try:
                ensure_folder(out_folder)
            except Exception:
                out_folder = ""

        try:
            self.set_status(f"Loading parser for {bank}...")
            parser = load_parser_module(bank)

            client_name = ""
            try:
                if hasattr(parser, "extract_account_holder_name"):
                    client_name = parser.extract_account_holder_name(self.selected_files[0]) or ""
            except Exception:
                client_name = ""
            if not client_name:
                client_name = get_client_name_from_pdf(self.selected_files[0])

            all_transactions = []
            recon_results = []
            per_pdf_txns: dict[str, list[dict]] = {}
            remove_txn_ids: set[int] = set()
            pdf_by_name: dict[str, str] = {os.path.basename(p): p for p in (self.selected_files or [])}

            if self.auto_detect_var.get() and len(self.selected_files) > 1:
                self.set_status("Detecting bank...")
                mismatches = []
                unknowns = []

                for p in self.selected_files:
                    detected = auto_detect_bank_from_pdf(p)
                    if detected is None:
                        unknowns.append(os.path.basename(p))
                    elif detected != bank:
                        mismatches.append(f"{os.path.basename(p)} → {detected}")

                if mismatches:
                    NL = chr(10)
                    msg = (
                        "Auto-detect thinks some PDFs may be from a different bank than the one selected." + NL + NL
                        + f"Selected bank: {bank}" + NL + NL
                        + NL.join(mismatches)
                        + NL + NL
                        + "This can happen if a statement contains another bank name/BIC in a transaction description, "
                        "or if the PDF has a cover/summary page." + NL + NL
                        + f"Do you want to continue and parse ALL PDFs as {bank}?"
                    )

                    proceed_anyway = messagebox.askyesno("Bank not confirmed", msg)
                    if not proceed_anyway:
                        self.set_status("Cancelled: bank not confirmed.")
                        return

                if unknowns:
                    NL = chr(10)
                    msg = (
                        "Auto-detect could not confirm the bank for some PDFs." + NL
                        + "I will continue, but results may be wrong if a different bank is included:" + NL + NL
                        + NL.join(unknowns)
                    )
                    messagebox.showwarning("Bank not confirmed", msg)

            def _status_startswith(v, prefix: str) -> bool:
                try:
                    return str(v or "").startswith(prefix)
                except Exception:
                    return False

            def _get_pdf_period(_parser, _pdf_path: str, _rec: dict):
                """Best-effort (period_start, period_end) from parser; else None/None.

                Expected parser hook (if present): extract_statement_period(pdf_path) -> (date|None, date|None)
                Also tries common alternative names.
                """
                fn_names = [
                    "extract_statement_period",
                    "extract_statement_period_dates",
                    "extract_period",
                    "extract_statement_date_range",
                    "extract_date_range",
                ]
                for nm in fn_names:
                    try:
                        fn = getattr(_parser, nm, None)
                    except Exception:
                        fn = None
                    if not callable(fn):
                        continue
                    try:
                        ps, pe = fn(_pdf_path)
                        return ps, pe
                    except TypeError:
                        # Some implementations may take (path, rec)
                        try:
                            ps, pe = fn(_pdf_path, _rec)
                            return ps, pe
                        except Exception:
                            continue
                    except Exception:
                        continue
                return None, None

            run_log_lines = []
            run_log_lines.append(f"Run time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            run_log_lines.append(f"Bank: {bank}")
            run_log_lines.append(f"PDF count: {len(self.selected_files)}")
            run_log_lines.append("")

            for i, pdf_path in enumerate(self.selected_files, start=1):
                self.set_status(f"Parsing: {os.path.basename(pdf_path)} ({i}/{len(self.selected_files)})")
                self.set_progress(i, max(1, len(self.selected_files)))
                txns = parser.extract_transactions(pdf_path)
                per_pdf_txns[pdf_path] = txns
                all_transactions.extend(txns)
                rec = reconcile_statement(parser, pdf_path, txns)

                # Provide txns to core continuity logic (overlap resolution uses these when present).
                rec["transactions"] = txns
                try:
                    _dates = [t.get("Date") for t in (txns or []) if t.get("Date")]
                    rec["date_min"] = min(_dates) if _dates else None
                    rec["date_max"] = max(_dates) if _dates else None
                except Exception:
                    rec["date_min"] = None
                    rec["date_max"] = None

                # Capture per-PDF statement period ONLY if parser exposes it; otherwise leave None.
                # (Starling and others may provide: extract_statement_period(pdf_path) -> (start, end))
                try:
                    if hasattr(parser, "extract_statement_period") and callable(getattr(parser, "extract_statement_period")):
                        ps, pe = parser.extract_statement_period(pdf_path)
                    else:
                        ps, pe = (None, None)
                except Exception:
                    ps, pe = (None, None)

                rec["period_start"] = ps
                rec["period_end"] = pe

                if bank == "Lloyds":
                    try:
                        start_d = rec.get("date_min")
                        opening = None
                        if start_d:
                            first_balance_idx = None
                            for j, t in enumerate(txns or []):
                                if t.get("Date") != start_d:
                                    continue
                                bal = t.get("Balance")
                                if bal is None or bal == "":
                                    continue
                                first_balance_idx = j
                                break

                            if first_balance_idx is not None:
                                bal_val = float((txns[first_balance_idx] or {}).get("Balance"))
                                net = 0.0
                                for j in range(0, first_balance_idx + 1):
                                    tt = txns[j]
                                    if tt.get("Date") == start_d:
                                        amt = tt.get("Amount")
                                        if amt is None or amt == "":
                                            continue
                                        net += float(amt)
                                opening = round(bal_val - net, 2)

                        rec["continuity_start_balance"] = opening if opening is not None else rec.get("start_balance")
                    except Exception:
                        rec["continuity_start_balance"] = rec.get("start_balance")
                else:
                    rec["continuity_start_balance"] = rec.get("start_balance")

                try:
                    rec["txn_count"] = len(txns or [])
                except Exception:
                    rec["txn_count"] = None
                try:
                    rec["fingerprint"] = compute_statement_fingerprint(txns)
                except Exception:
                    rec["fingerprint"] = None

                recon_results.append(rec)

                run_log_lines.append(f"{os.path.basename(pdf_path)}")
                run_log_lines.append(f"  Transactions: {len(txns)}")
                run_log_lines.append(f"  Reconciliation: {rec.get('status')}")

                # Period visibility for debugging (core overlap/chronology uses period_start/period_end).
                try:
                    ps = rec.get("period_start")
                    pe = rec.get("period_end")
                    if ps and pe and hasattr(ps, "strftime") and hasattr(pe, "strftime"):
                        run_log_lines.append(f"  Period: {ps.strftime('%d/%m/%Y')} - {pe.strftime('%d/%m/%Y')}")
                    else:
                        run_log_lines.append("  Period: None")
                except Exception:
                    run_log_lines.append("  Period: None")

                if rec.get("status") in ("OK", "Mismatch"):
                    run_log_lines.append(
                        f"  Start: {_fmt_money(rec.get('start_balance'))} | Net: {_fmt_money(rec.get('sum_amounts'))} | End: {_fmt_money(rec.get('end_balance'))}"
                    )
                    if rec.get("status") == "Mismatch":
                        run_log_lines.append(f"  Diff: {_fmt_money(rec.get('difference'))}")
                run_log_lines.append("")

            self.set_status("Running reconciliation checks...")
            duplicate_groups = find_duplicate_statements(recon_results)
            if duplicate_groups:
                def _fmt_date(v):
                    try:
                        if v is None or v == "":
                            return ""
                        if hasattr(v, "to_pydatetime"):
                            v = v.to_pydatetime()
                        if hasattr(v, "strftime"):
                            return v.strftime("%d/%m/%Y")
                        return str(v)
                    except Exception:
                        return str(v)

                msg_lines = []
                msg_lines.append("Duplicate statements detected. Please remove the duplicates and run again.")
                msg_lines.append("")

                for gi, grp in enumerate(duplicate_groups, start=1):
                    msg_lines.append(f"Group {gi} ({len(grp)} files):")
                    for r in grp:
                        msg_lines.append(f"  - {r.get('pdf')}")

                    try:
                        dmin = _fmt_date(grp[0].get("date_min"))
                        dmax = _fmt_date(grp[0].get("date_max"))
                        dr = (f"{dmin} - {dmax}").strip(" -")
                    except Exception:
                        dr = ""

                    try:
                        start = _fmt_money(grp[0].get("start_balance"))
                        end = _fmt_money(grp[0].get("end_balance"))
                    except Exception:
                        start, end = "", ""

                    try:
                        txc = grp[0].get("txn_count")
                    except Exception:
                        txc = ""

                    summary_parts = []
                    if dr:
                        summary_parts.append(f"Dates {dr}")
                    if start or end:
                        summary_parts.append(f"Start {start} / End {end}")
                    if txc not in (None, ""):
                        summary_parts.append(f"Txns {txc}")

                    if summary_parts:
                        msg_lines.append("  Summary: " + " | ".join(summary_parts))

                    msg_lines.append("")

                try:
                    ensure_folder(LOGS_DIR)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    log_path = os.path.join(LOGS_DIR, f"duplicate_statements_{ts}.txt")
                    with open(log_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(msg_lines).strip() + "\n")
                except Exception:
                    pass

                messagebox.showerror("Duplicate statements detected", "\n".join(msg_lines).strip())

                self.set_status("Error: duplicate statements detected.")
                return

            if not all_transactions:
                raise ValueError("No transactions found!")

            def _date_key(x):
                d = x.get("Date")
                return d if d is not None else datetime.min.date()

            all_transactions.sort(key=_date_key)

            date_values = []
            for t in all_transactions:
                d = t.get("Date")
                if d:
                    date_values.append(d)
            date_min = min(date_values) if date_values else None
            date_max = max(date_values) if date_values else None

            filename = build_output_filename(client_name, date_min, date_max)

            self.set_status("Running continuity checks...")
            continuity_results = compute_statement_continuity(recon_results)

            # Apply overlap de-duplication results produced by core continuity logic.
            # Core will populate overlap_* fields on each continuity link when applicable.
            rec_by_pdfname: dict[str, dict] = {}
            try:
                for r in (recon_results or []):
                    nm = r.get("pdf")
                    if nm:
                        rec_by_pdfname[str(nm)] = r
            except Exception:
                rec_by_pdfname = {}

            for link in (continuity_results or []):
                # Prefer core's display_status for UI/popup/logging.
                try:
                    ds = link.get("display_status")
                    if ds:
                        link["status"] = str(ds)
                except Exception:
                    pass

                prev_pdf = str(link.get("prev_pdf") or "")
                next_pdf = str(link.get("next_pdf") or "")

                applied = False
                try:
                    applied = bool(link.get("applied_overlap_resolution"))
                except Exception:
                    applied = False

                dup_idx = link.get("duplicates_to_remove_from_B")
                if not isinstance(dup_idx, list):
                    dup_idx = []

                removed_n_effective = 0
                expected_removed_for_link = 0

                if applied and dup_idx:
                    # Apply removal plan: remove indices from B ONLY, in descending order.
                    # IMPORTANT: ensure we remove from the SAME B txn list that will be used to rebuild the combined export.
                    b_rec = rec_by_pdfname.get(next_pdf)
                    b_txns = None
                    try:
                        if b_rec is not None:
                            b_txns = b_rec.get("transactions")
                    except Exception:
                        b_txns = None

                    if not isinstance(b_txns, list):
                        b_txns = []

                    # Locate B's original pdf_path so we can also mutate per_pdf_txns (if it is a different list object).
                    b_path = None
                    try:
                        b_path = pdf_by_name.get(next_pdf)
                    except Exception:
                        b_path = None
                    if not b_path:
                        try:
                            for p in (self.selected_files or []):
                                if os.path.basename(p) == next_pdf:
                                    b_path = p
                                    break
                        except Exception:
                            b_path = None

                    b_path_txns = None
                    try:
                        if b_path:
                            b_path_txns = per_pdf_txns.get(b_path)
                    except Exception:
                        b_path_txns = None

                    if not isinstance(b_path_txns, list):
                        b_path_txns = None

                    # Normalise & sort indices descending to avoid shifting.
                    idxs = []
                    for x in dup_idx:
                        try:
                            idxs.append(int(x))
                        except Exception:
                            continue
                    idxs = sorted(set(idxs), reverse=True)

                    expected_removed_for_link = len(idxs)

                    # Remove from b_txns
                    for ii in idxs:
                        if 0 <= ii < len(b_txns):
                            try:
                                removed_txn = b_txns.pop(ii)
                                remove_txn_ids.add(id(removed_txn))  # fallback safety
                                removed_n_effective += 1
                            except Exception:
                                pass

                    # If per_pdf_txns uses a different list object, remove there too so rebuild matches.
                    if b_path_txns is not None and b_path_txns is not b_txns:
                        for ii in idxs:
                            if 0 <= ii < len(b_path_txns):
                                try:
                                    removed_txn = b_path_txns.pop(ii)
                                    remove_txn_ids.add(id(removed_txn))
                                except Exception:
                                    pass

                # Logging / verification for each continuity link
                # Track expected vs effective removals for output verification.
                try:
                    expected_total_removed_from_core = expected_total_removed_from_core
                except Exception:
                    expected_total_removed_from_core = 0
                try:
                    effective_total_removed_in_lists = effective_total_removed_in_lists
                except Exception:
                    effective_total_removed_in_lists = 0

                if applied and dup_idx:
                    try:
                        expected_total_removed_from_core += int(link.get("removed_count") or expected_removed_for_link or 0)
                    except Exception:
                        expected_total_removed_from_core += int(expected_removed_for_link or 0)
                    effective_total_removed_in_lists += int(removed_n_effective or 0)

                try:
                    ow = link.get("overlap_window")
                    win = ""
                    if isinstance(ow, dict):
                        ws = ow.get("start")
                        we = ow.get("end")
                        if ws and we and hasattr(ws, "strftime") and hasattr(we, "strftime"):
                            win = f"{ws.strftime('%d/%m/%Y')} - {we.strftime('%d/%m/%Y')}"
                except Exception:
                    win = ""

                removed_count = link.get("removed_count")
                if removed_count in (None, ""):
                    removed_count = removed_n_effective
                dupe_sum = link.get("dupe_sum")

                chrono_applied = link.get("chronology_gate_applied")
                chrono_note = link.get("chronology_gate_note")

                if applied:
                    run_log_lines.append(
                        f"Continuity: {prev_pdf} -> {next_pdf} | overlap YES"
                        + (f" | window {win}" if win else "")
                        + f" | removed {removed_count}"
                        + (f" | dupe_sum {_fmt_money(dupe_sum)}" if dupe_sum not in (None, "") else "")
                    )
                else:
                    first_overlap_line = ""
                    try:
                        lines = link.get("overlap_log_lines")
                        if isinstance(lines, list) and lines:
                            first_overlap_line = str(lines[0])
                    except Exception:
                        first_overlap_line = ""

                    run_log_lines.append(
                        f"Continuity: {prev_pdf} -> {next_pdf} | overlap NO"
                        + (f" | note {first_overlap_line}" if first_overlap_line else "")
                    )

                # Chain/chronology debug info (lightweight)
                try:
                    if chrono_applied is not None:
                        run_log_lines.append(
                            f"  Chronology gate applied: {'YES' if chrono_applied else 'NO'}"
                            + (f" | {chrono_note}" if chrono_note else "")
                        )
                except Exception:
                    pass

                try:
                    cc_total = link.get("chain_candidates_total")
                    if cc_total is not None:
                        run_log_lines.append(
                            "  Chain candidates: "
                            + f"total={link.get('chain_candidates_total')} "
                            + f"known={link.get('chain_candidates_known_period_start')} "
                            + f"pass={link.get('chain_candidates_chrono_pass')} "
                            + f"fail={link.get('chain_candidates_chrono_fail')} "
                            + f"unknown={link.get('chain_candidates_chrono_unknown')}"
                        )
                except Exception:
                    pass

                run_log_lines.append("")

            # Rebuild combined transactions AFTER applying de-dupe (so Excel output reflects the plan).
            # Rebuild whenever any overlap removal was applied, even if ids couldn't be captured.
            try:
                _need_rebuild = bool(effective_total_removed_in_lists)
            except Exception:
                _need_rebuild = bool(remove_txn_ids)

            if _need_rebuild:
                rebuilt = []
                for p in (self.selected_files or []):
                    rebuilt.extend(per_pdf_txns.get(p) or [])

                all_transactions = list(rebuilt)

                # Log verification: expected (core) vs effective removals vs output delta.
                try:
                    expected_n = int(expected_total_removed_from_core)
                except Exception:
                    expected_n = 0
                try:
                    effective_n = int(effective_total_removed_in_lists)
                except Exception:
                    effective_n = 0

                run_log_lines.append(
                    f"Overlap de-duplication applied: removed {effective_n} transactions from output"
                    + (f" (core expected {expected_n})" if expected_n else "")
                )
                run_log_lines.append("")

            any_gap = any(_status_startswith((r.get('status') or ''), 'Mismatch') for r in (continuity_results or []))

            coverage_period = ""
            try:
                if date_min and date_max:
                    coverage_period = f"{date_min.strftime('%d/%m/%Y')} to {date_max.strftime('%d/%m/%Y')}"
            except Exception:
                coverage_period = ""
            any_issue = any(
                (r.get("status") not in ("OK", "Not checked"))
                for r in (recon_results or [])
            )

            # Continuity: anything not starting with OK (including NOT CHECKED / balances not found)
            # should be treated as an issue so logs/support bundles are produced.
            any_cont_issue = any(
                not _status_startswith((c.get("status") or c.get("display_status") or ""), "OK")
                for c in (continuity_results or [])
            )

            any_issue = any_issue or any_cont_issue or any_gap

            recon_log_path = None

            initial_dir = ""
            if out_folder:
                initial_dir = out_folder
            else:
                try:
                    initial_dir = os.path.dirname(self.selected_files[0])
                except Exception:
                    initial_dir = ""

            self.last_saved_output_path = None
            any_warn = any((r.get("status") or "") != "OK" for r in (recon_results or [])) or any(
                not _status_startswith((c.get("display_status") or c.get("status") or ""), "OK")
                for c in (continuity_results or [])
                if isinstance(c, dict)
            )

            autodetect_first_pdf = None
            try:
                if self.auto_detect_var.get() and self.selected_files:
                    autodetect_first_pdf = auto_detect_bank_from_pdf(self.selected_files[0])
            except Exception:
                autodetect_first_pdf = None

            parser_file = ""
            try:
                parser_file = getattr(parser, "__file__", "") or ""
            except Exception:
                parser_file = ""

            self.last_report_data = {
                "recon_results": recon_results,
                "continuity_results": continuity_results,
                "coverage_period": coverage_period,
                "source_pdfs": list(self.selected_files or []),
                "any_warn": bool(any_warn),
                "log_path": recon_log_path,
                "learning_report_path": None,
                "output_xlsx_path": None,
                "bundle_base": os.path.splitext(filename)[0],
                "bank": bank,
                "autodetect_first_pdf": autodetect_first_pdf,
                "parser_file": parser_file,
                "client_name": client_name,
                "run_filename": filename,
            }
            self.last_excel_data = {
                "transactions": all_transactions,
                "client_name": client_name,
                "filename": filename,
                "initial_dir": initial_dir,
            }

            if any_issue:
                any_recon_mismatch = any((r.get("status") or "") == "Mismatch" for r in (recon_results or []))
                any_cont_mismatch = any(
                    _status_startswith((c.get("display_status") or c.get("status") or ""), "Mismatch")
                    for c in (continuity_results or [])
                    if isinstance(c, dict)
                )
                issue_reason = "Mismatch" if (any_recon_mismatch or any_cont_mismatch) else "Issue"
                try:
                    self.generate_learning_report(reason=issue_reason)
                except Exception:
                    pass

            # Auto-create support bundle when we have an issue OR warnings.
            try:
                if any_issue or any_warn:
                    self.create_support_bundle_zip()
            except Exception as e:
                messagebox.showwarning(
                    "Support bundle",
                    f"Could not create support bundle automatically:\n{e}",
                )

            proceed = show_reconciliation_popup(
                self,
                "(Not saved yet)",
                recon_results,
                coverage_period=coverage_period,
                continuity_results=continuity_results,
                pre_save=True,
                support_bundle_callback=self.create_support_bundle_zip,
            )

            if not proceed:
                self.set_status("Cancelled.")
                return

            output_path = filedialog.asksaveasfilename(
                title="Save Excel file",
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx")],
                initialdir=initial_dir or None,
                initialfile=filename,
            )

            if not output_path:
                self.set_status("Cancelled.")
                return

            try:
                ensure_folder(os.path.dirname(output_path))
            except Exception as e:
                messagebox.showerror("Save error", f"Cannot create folder for output file:\n{e}")
                self.set_status("Error.")
                return

            try:
                self.output_folder_var.set(os.path.dirname(output_path))
            except Exception:
                pass

            self.set_status("Writing Excel...")
            save_transactions_to_excel(all_transactions, output_path, client_name=client_name)

            self.last_saved_output_path = output_path

            try:
                if self.last_report_data is not None:
                    self.last_report_data["output_xlsx_path"] = output_path
            except Exception:
                pass

            self.set_progress(len(self.selected_files), max(1, len(self.selected_files)))
            if any_warn:
                self.set_status(f"Done with warnings. Output: {output_path}")
            else:
                self.set_status(f"Done. Output: {output_path}")

        except Exception as e:
            self.set_status("Error.")
            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))

            try:
                if self.last_report_data is None:
                    self.last_report_data = {
                        "recon_results": [],
                        "continuity_results": [],
                        "coverage_period": "",
                        "source_pdfs": list(self.selected_files or []),
                        "any_warn": True,
                        "log_path": None,
                        "learning_report_path": None,
                        "output_xlsx_path": None,
                        "bundle_base": "RUN",
                        "bank": bank,
                        "autodetect_first_pdf": None,
                        "parser_file": "",
                        "client_name": "",
                        "run_filename": "",
                    }
                self.generate_learning_report(reason="Exception", exception=e)
            except Exception:
                pass

            try:
                ensure_folder(LOGS_DIR)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                crash_name = f"crash_{ts}.txt"
                crash_path = os.path.join(LOGS_DIR, crash_name)
                with open(crash_path, "w", encoding="utf-8") as f:
                    f.write(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Bank: {bank}\n")
                    f.write(f"Output folder: {out_folder}\n")
                    f.write("PDFs:\n")
                    for p in (self.selected_files or []):
                        f.write(f"  - {p}\n")
                    f.write("\nException:\n")
                    f.write(err)
            except Exception:
                pass

            messagebox.showerror("Error", f"{e}\n\nDetails:\n{err}")


def _self_tests():
    # Minimal, optional sanity checks (no GUI launched).
    assert _fmt_money(None) == ""
    assert _fmt_money(0) == "£0.00"
    assert _fmt_money(12.3) == "£12.30"
    assert _fmt_money(-12.3) == "-£12.30"
    assert _fmt_money("£1,234.50") == "£1,234.50"


if __name__ == "__main__":
    # Set GUI_SELFTEST=1 to run quick format tests without launching the app.
    if os.environ.get("GUI_SELFTEST") == "1":
        _self_tests()
    else:
        app = App()
        app.mainloop()

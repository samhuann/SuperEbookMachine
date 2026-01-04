import shutil
import subprocess
import threading
import webbrowser
import sys
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# Common formats Calibre typically supports well as inputs
DEFAULT_INPUT_EXTS = [
    ".pdf", ".epub", ".mobi", ".azw", ".azw3", ".fb2", ".rtf", ".txt",
    ".docx", ".doc", ".html", ".htm", ".cbz", ".cbr"
]


def find_ebook_convert(user_path: str) -> str:
    """
    Find Calibre's ebook-convert executable.
    - If user_path is provided, validate it.
    - Else try PATH.
    """
    user_path = (user_path or "").strip()
    if user_path:
        p = Path(user_path)
        if p.exists():
            return str(p)
        raise FileNotFoundError(f"ebook-convert not found at: {user_path}")

    which = shutil.which("ebook-convert")
    if which:
        return which

    # Try common Windows path silently as a convenience
    common = Path(r"C:\Program Files\Calibre2\ebook-convert.exe")
    if common.exists():
        return str(common)

    raise FileNotFoundError(
        "Could not find 'ebook-convert'.\n\n"
        "Install Calibre (required), then either:\n"
        "  • Ensure ebook-convert is on PATH, OR\n"
        r"  • Browse to C:\Program Files\Calibre2\ebook-convert.exe"
    )


def normalize_ext_list(s: str) -> set[str]:
    """
    Parse comma-separated extensions into canonical format: {'.pdf', '.epub'}
    Accepts: 'pdf, epub' or '.pdf, .epub'
    """
    out = set()
    for part in s.split(","):
        part = part.strip().lower()
        if not part:
            continue
        if not part.startswith("."):
            part = "." + part
        out.add(part)
    return out


def iter_selected_files(in_root: Path, exts: set[str]):
    for p in in_root.rglob("*"):
        if p.is_file() and p.suffix.lower() in exts:
            yield p


def _sanitize_filename_component(s: str) -> str:
    # Windows-invalid filename characters: < > : " / \ | ? *
    # Also avoid trailing dots/spaces.
    bad = '<>:"/\\|?*'
    out = "".join((ch if ch not in bad else "_") for ch in s)
    out = out.strip().strip(".")
    return out or "_"


def build_out_path(
    in_root: Path,
    out_root: Path,
    in_path: Path,
    out_fmt: str,
    *,
    flatten: bool,
    keep_input_ext: bool,
) -> Path:
    """Build an output path.

    - If flatten=False: preserve folder structure under out_root.
    - If flatten=True: write everything directly into out_root.
      To avoid collisions, include the relative parent folders in the filename.
    - If keep_input_ext=True: keep the input file extension (copy mode).
      Else: use out_fmt.
    """
    rel = in_path.relative_to(in_root)
    out_suffix = in_path.suffix if keep_input_ext else ("." + out_fmt.lower().lstrip("."))

    if not flatten:
        return (out_root / rel).with_suffix(out_suffix)

    base = _sanitize_filename_component(in_path.stem)
    parent_parts = [_sanitize_filename_component(p) for p in rel.parts[:-1]]
    if parent_parts:
        tag = "__".join(parent_parts)
        base = f"{base}__{tag}"
    filename = f"{base}{out_suffix}"
    return out_root / filename


class SuperEbookMachine(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SuperEbookMachine")
        self.geometry("1020x700")

        self._set_window_icon()

        # Thread control
        self.stop_event = threading.Event()
        self.worker_thread = None

        # Paths
        self.in_root_var = tk.StringVar()
        self.out_root_var = tk.StringVar()
        self.calibre_path_var = tk.StringVar()

        # Target and conversion options
        self.target_var = tk.StringVar(value="app")  # "app" or "device"
        self.out_fmt_var = tk.StringVar(value="epub")  # set by target on start
        self.profile_var = tk.StringVar(value="kindle")
        self.workers_var = tk.IntVar(value=6)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.flatten_output_var = tk.BooleanVar(value=False)
        self.no_convert_var = tk.BooleanVar(value=False)

        # Input scanning options
        self.ext_vars = {ext: tk.BooleanVar(value=(ext == ".pdf")) for ext in DEFAULT_INPUT_EXTS}
        self.use_custom_exts_var = tk.BooleanVar(value=False)
        self.custom_exts_var = tk.StringVar(value=".pdf")

        # Progress & counters
        self.progress_var = tk.IntVar(value=0)
        self.count_var = tk.StringVar(value="0/0")
        self.status_var = tk.StringVar(value="Ready.")

        self.total = 0
        self.done = 0
        self.ok = 0
        self.skip = 0
        self.fail = 0

        self._build_ui()
        self.on_target_change()  # set initial format hint

    def _resource_path(self, relative_path: str) -> Path:
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        return base / relative_path

    def _set_window_icon(self):
        icon_path = self._resource_path("SuperEbookMachine.ico")
        try:
            if icon_path.exists():
                self.iconbitmap(str(icon_path))
        except Exception:
            # If icon loading fails (or on non-Windows), keep default.
            pass

    # ---------------- UI ----------------
    def _build_ui(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        # Row 0: Input root
        ttk.Label(frm, text="Input root folder (contains subfolders):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.in_root_var, width=82).grid(row=0, column=1, sticky="we", padx=(8, 8))
        ttk.Button(frm, text="Browse…", command=self.browse_in).grid(row=0, column=2, sticky="e")

        # Row 1: Output root
        ttk.Label(frm, text="Output root folder:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.out_root_var, width=82).grid(row=1, column=1, sticky="we", padx=(8, 8))
        ttk.Button(frm, text="Browse…", command=self.browse_out).grid(row=1, column=2, sticky="e")

        # Row 2: Calibre path
        ttk.Label(frm, text="Calibre ebook-convert path (optional):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.calibre_path_var, width=82).grid(row=2, column=1, sticky="we", padx=(8, 8))
        ttk.Button(frm, text="Browse…", command=self.browse_calibre).grid(row=2, column=2, sticky="e")

        # Row 3: Target group
        target = ttk.LabelFrame(frm, text="Target")
        target.grid(row=3, column=0, columnspan=3, sticky="we", pady=(10, 6))
        target.columnconfigure(0, weight=1)

        ttk.Radiobutton(
            target,
            text="Kindle App (Phone / Desktop / Cloud) — use EPUB + Send-to-Kindle",
            variable=self.target_var,
            value="app",
            command=self.on_target_change
        ).grid(row=0, column=0, sticky="w", padx=8, pady=4)

        ttk.Radiobutton(
            target,
            text="Physical Kindle (USB) — use AZW3 + copy to Kindle/documents/",
            variable=self.target_var,
            value="device",
            command=self.on_target_change
        ).grid(row=1, column=0, sticky="w", padx=8, pady=4)

        ttk.Button(target, text="Open Send-to-Kindle", command=self.open_send_to_kindle).grid(
            row=0, column=1, rowspan=2, padx=12, pady=4
        )

        # Row 4: Scan input types
        types_frame = ttk.LabelFrame(frm, text="Scan input file types")
        types_frame.grid(row=4, column=0, columnspan=3, sticky="we", pady=6)

        # Checkboxes
        r = 0
        c = 0
        for ext in DEFAULT_INPUT_EXTS:
            ttk.Checkbutton(types_frame, text=ext, variable=self.ext_vars[ext]).grid(
                row=r, column=c, sticky="w", padx=8, pady=3
            )
            c += 1
            if c >= 6:
                c = 0
                r += 1

        ttk.Checkbutton(
            types_frame,
            text="Use custom extensions (comma-separated):",
            variable=self.use_custom_exts_var,
            command=self._refresh_custom_state
        ).grid(row=r + 1, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 2))

        self.custom_entry = ttk.Entry(types_frame, textvariable=self.custom_exts_var, width=60)
        self.custom_entry.grid(row=r + 1, column=2, columnspan=4, sticky="we", padx=(8, 8), pady=(8, 2))

        quick = ttk.Frame(types_frame)
        quick.grid(row=r + 2, column=0, columnspan=6, sticky="w", padx=8, pady=(4, 8))
        ttk.Button(quick, text="Only PDFs", command=self.select_only_pdfs).pack(side="left")
        ttk.Button(quick, text="Common (PDF+EPUB+MOBI)", command=self.select_common).pack(side="left", padx=8)
        ttk.Button(quick, text="Select none", command=self.select_none).pack(side="left")

        self._refresh_custom_state()

        # Row 5: Conversion options
        opt = ttk.LabelFrame(frm, text="Conversion options")
        opt.grid(row=5, column=0, columnspan=3, sticky="we", pady=6)

        ttk.Label(opt, text="Output format:").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        self.out_fmt_combo = ttk.Combobox(
            opt, textvariable=self.out_fmt_var, width=10, state="readonly",
            values=["azw3", "epub", "mobi"]
        )
        self.out_fmt_combo.grid(row=0, column=1, sticky="w", padx=(6, 18), pady=8)

        ttk.Label(opt, text="Output profile:").grid(row=0, column=2, sticky="w", pady=8)
        self.profile_combo = ttk.Combobox(
            opt, textvariable=self.profile_var, width=12, state="readonly",
            values=["kindle", "kindle_pw3", "tablet", "default"]
        )
        self.profile_combo.grid(row=0, column=3, sticky="w", padx=(6, 18), pady=8)

        ttk.Label(opt, text="Workers:").grid(row=0, column=4, sticky="w", pady=8)
        ttk.Spinbox(opt, from_=1, to=32, textvariable=self.workers_var, width=6).grid(
            row=0, column=5, sticky="w", padx=(6, 18), pady=8
        )

        ttk.Checkbutton(opt, text="Overwrite existing outputs", variable=self.overwrite_var).grid(
            row=0, column=6, sticky="w", padx=(6, 8), pady=8
        )

        ttk.Checkbutton(
            opt,
            text="Flatten output (ignore subfolders)",
            variable=self.flatten_output_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 8))

        ttk.Checkbutton(
            opt,
            text="Copy originals (no conversion / keep file type)",
            variable=self.no_convert_var,
            command=self._refresh_convert_state,
        ).grid(row=1, column=3, columnspan=4, sticky="w", padx=8, pady=(0, 8))

        # Row 6: Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, sticky="we", pady=6)

        self.start_btn = ttk.Button(btns, text="Start", command=self.on_start)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(btns, text="Stop", command=self.on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=8)

        ttk.Button(btns, text="Test ebook-convert", command=self.on_test_calibre).pack(side="left", padx=8)
        ttk.Button(btns, text="Clear log", command=self.clear_log).pack(side="right")

        # Row 7: Progress
        prog_row = ttk.Frame(frm)
        prog_row.grid(row=7, column=0, columnspan=3, sticky="we", pady=(10, 0))
        prog_row.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            prog_row, orient="horizontal", mode="determinate",
            maximum=1, variable=self.progress_var
        )
        self.progress.grid(row=0, column=0, sticky="we")
        ttk.Label(prog_row, textvariable=self.count_var, width=12, anchor="e").grid(row=0, column=1, padx=(10, 0))

        ttk.Label(frm, textvariable=self.status_var).grid(row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # Row 9: Log
        self.log = tk.Text(frm, height=18, wrap="word")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        scroll = ttk.Scrollbar(frm, orient="vertical", command=self.log.yview)
        scroll.grid(row=9, column=3, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(9, weight=1)

        self._refresh_convert_state()

    def _refresh_convert_state(self):
        copying = bool(self.no_convert_var.get())
        self.out_fmt_combo.config(state="disabled" if copying else "readonly")
        self.profile_combo.config(state="disabled" if copying else "readonly")

    # ---------------- Target behavior ----------------
    def on_target_change(self):
        """
        Set recommended output format based on target.
        """
        if self.target_var.get() == "app":
            self.out_fmt_var.set("epub")
            self.status_var.set("Target: Kindle App → EPUB (use Send-to-Kindle).")
        else:
            self.out_fmt_var.set("azw3")
            self.status_var.set("Target: Physical Kindle → AZW3 (USB transfer).")

    def open_send_to_kindle(self):
        webbrowser.open("https://www.amazon.com/sendtokindle")

    # ---------------- Scan type behavior ----------------
    def _refresh_custom_state(self):
        self.custom_entry.config(state="normal" if self.use_custom_exts_var.get() else "disabled")

    def select_only_pdfs(self):
        for ext, var in self.ext_vars.items():
            var.set(ext == ".pdf")
        self.use_custom_exts_var.set(False)
        self._refresh_custom_state()

    def select_common(self):
        common = {".pdf", ".epub", ".mobi"}
        for ext, var in self.ext_vars.items():
            var.set(ext in common)
        self.use_custom_exts_var.set(False)
        self._refresh_custom_state()

    def select_none(self):
        for var in self.ext_vars.values():
            var.set(False)

    def get_selected_exts(self) -> set[str]:
        if self.use_custom_exts_var.get():
            exts = normalize_ext_list(self.custom_exts_var.get())
            if not exts:
                raise ValueError("Custom extensions enabled, but none were provided.")
            return exts

        exts = {ext for ext, var in self.ext_vars.items() if var.get()}
        if not exts:
            raise ValueError("No input types selected. Choose at least one extension.")
        return exts

    # ---------------- UI helpers ----------------
    def browse_in(self):
        d = filedialog.askdirectory(title="Select input root folder")
        if d:
            self.in_root_var.set(d)

    def browse_out(self):
        d = filedialog.askdirectory(title="Select output root folder")
        if d:
            self.out_root_var.set(d)

    def browse_calibre(self):
        f = filedialog.askopenfilename(
            title="Select ebook-convert executable",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")]
        )
        if f:
            self.calibre_path_var.set(f)

    def log_line(self, s: str):
        self.log.insert("end", s + "\n")
        self.log.see("end")

    def clear_log(self):
        self.log.delete("1.0", "end")

    def set_running(self, running: bool):
        self.start_btn.config(state="disabled" if running else "normal")
        self.stop_btn.config(state="normal" if running else "disabled")

    def ui_call(self, fn):
        self.after(0, fn)

    def refresh_counters(self):
        self.count_var.set(f"{self.done}/{self.total}")
        self.status_var.set(
            f"Progress: {self.done}/{self.total}  |  OK: {self.ok}  SKIP: {self.skip}  FAIL: {self.fail}"
        )

    # ---------------- Actions ----------------
    def on_test_calibre(self):
        try:
            calibre = find_ebook_convert(self.calibre_path_var.get())
            out = subprocess.run([calibre, "--version"], capture_output=True, text=True, check=True)
            msg = out.stdout.strip() or out.stderr.strip() or "OK"
            messagebox.showinfo("ebook-convert OK", msg)
        except Exception as e:
            messagebox.showerror("ebook-convert error", str(e))

    def on_stop(self):
        self.stop_event.set()
        self.status_var.set("Stopping… (waiting for in-flight conversions to finish)")

    def on_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        in_root = Path(self.in_root_var.get().strip())
        out_root = Path(self.out_root_var.get().strip())

        if not in_root.exists() or not in_root.is_dir():
            messagebox.showerror("Invalid input", "Please select a valid input root folder.")
            return
        if not out_root:
            messagebox.showerror("Invalid output", "Please select a valid output root folder.")
            return

        # Calibre required
        try:
            calibre = find_ebook_convert(self.calibre_path_var.get())
        except Exception as e:
            messagebox.showerror("Calibre not found", str(e))
            return

        # Selected scan types
        try:
            selected_exts = self.get_selected_exts()
        except Exception as e:
            messagebox.showerror("Input types", str(e))
            return

        copying = bool(self.no_convert_var.get())

        # Target sanity check (help users avoid AZW3-in-app confusion)
        if (not copying) and self.target_var.get() == "app" and self.out_fmt_var.get().lower() != "epub":
            if messagebox.askyesno(
                "Format mismatch",
                "Kindle apps require EPUB via Send-to-Kindle.\n\nSwitch output format to EPUB?"
            ):
                self.out_fmt_var.set("epub")
            else:
                return

        out_fmt = self.out_fmt_var.get().strip().lower()
        profile = self.profile_var.get().strip()
        workers = max(1, int(self.workers_var.get()))
        overwrite = bool(self.overwrite_var.get())
        flatten = bool(self.flatten_output_var.get())

        # Reset state
        self.stop_event.clear()
        self.set_running(True)

        self.total = self.done = self.ok = self.skip = self.fail = 0
        self.progress_var.set(0)
        self.progress.configure(maximum=1)
        self.count_var.set("0/0")
        self.status_var.set("Scanning…")

        self.log_line(f"SuperEbookMachine starting…")
        self.log_line(f"ebook-convert: {calibre}")
        self.log_line(f"Input:  {in_root}")
        self.log_line(f"Output: {out_root}")
        self.log_line(f"Scan types: {', '.join(sorted(selected_exts))}")
        self.log_line(f"Target: {'Kindle App' if self.target_var.get() == 'app' else 'Physical Kindle'}")
        if copying:
            self.log_line(f"Mode: COPY (no conversion) | Workers: {workers} | Overwrite: {overwrite} | Flatten: {flatten}")
        else:
            self.log_line(
                f"Convert -> {out_fmt} | Profile: {profile} | Workers: {workers} | Overwrite: {overwrite} | Flatten: {flatten}"
            )
        self.log_line("----")

        def background():
            try:
                files = list(iter_selected_files(in_root, selected_exts))
                total = len(files)

                def init_total():
                    self.total = total
                    self.progress.configure(maximum=max(1, total))
                    self.progress_var.set(0)
                    self.done = 0
                    self.refresh_counters()

                self.ui_call(init_total)

                if total == 0:
                    self.ui_call(lambda: self.status_var.set("No matching files found."))
                    return

                extra_args = [] if copying else ["--output-profile", profile]

                with ThreadPoolExecutor(max_workers=workers) as ex:
                    futures = []
                    for f in files:
                        if self.stop_event.is_set():
                            break
                        outp = build_out_path(
                            in_root,
                            out_root,
                            f,
                            out_fmt,
                            flatten=flatten,
                            keep_input_ext=copying,
                        )
                        futures.append(
                            ex.submit(self._process_one, calibre, f, outp, extra_args, overwrite, copying)
                        )

                    for fut in as_completed(futures):
                        if self.stop_event.is_set():
                            break
                        status, msg = fut.result()

                        def apply_result():
                            self.done += 1
                            if status == "ok":
                                self.ok += 1
                            elif status == "skip":
                                self.skip += 1
                            else:
                                self.fail += 1

                            self.log_line(msg)
                            self.progress_var.set(self.done)
                            self.refresh_counters()

                        self.ui_call(apply_result)

                if self.stop_event.is_set():
                    self.ui_call(lambda: self.status_var.set("Stopped by user (some conversions may still complete)."))
                else:
                    self.ui_call(self.refresh_counters)

            except Exception as e:
                self.ui_call(lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.ui_call(lambda: self.set_running(False))

        self.worker_thread = threading.Thread(target=background, daemon=True)
        self.worker_thread.start()

    def _process_one(
        self,
        calibre: str,
        in_path: Path,
        out_path: Path,
        extra_args,
        overwrite: bool,
        copying: bool,
    ):
        out_path.parent.mkdir(parents=True, exist_ok=True)

        if out_path.exists() and not overwrite:
            return "skip", f"SKIP exists: {out_path}"

        if copying:
            try:
                shutil.copy2(in_path, out_path)
                return "ok", f"OK   {out_path}"
            except Exception as e:
                return "fail", f"FAIL {in_path} -> {out_path} :: {e}"

        cmd = [calibre, str(in_path), str(out_path)] + list(extra_args)

        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
            return "ok", f"OK   {out_path}"
        except subprocess.CalledProcessError as e:
            err = e.stderr.decode(errors="ignore").strip()
            last_line = err.splitlines()[-1] if err else "Unknown error"
            return "fail", f"FAIL {in_path} -> {out_path} :: {last_line}"


if __name__ == "__main__":
    SuperEbookMachine().mainloop()

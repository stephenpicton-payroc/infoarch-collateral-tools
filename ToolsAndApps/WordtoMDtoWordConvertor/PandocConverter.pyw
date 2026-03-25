"""
Pandoc Converter — lightweight GUI for docx <-> markdown conversion.
Requires Python 3.7+ with tkinter and a locally installed pandoc binary.
"""

import os
import shutil
import subprocess
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {".docx", ".md"}

PANDOC_FALLBACK_PATHS = [
    r"C:\Program Files\Pandoc\pandoc.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Pandoc", "pandoc.exe"),
]

CONVERSION_TIMEOUT = 120  # seconds

# ---------------------------------------------------------------------------
# Pure helper functions
# ---------------------------------------------------------------------------


def find_pandoc() -> str | None:
    """Return the path to the pandoc binary, or None if not found."""
    path = shutil.which("pandoc")
    if path:
        return path
    for candidate in PANDOC_FALLBACK_PATHS:
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


def detect_direction(input_path: str) -> str | None:
    """Return 'docx_to_md', 'md_to_docx', or None for unsupported extensions."""
    ext = Path(input_path).suffix.lower()
    if ext == ".docx":
        return "docx_to_md"
    if ext == ".md":
        return "md_to_docx"
    return None


def derive_output_filename(input_path: str, direction: str) -> str:
    """Return the output filename (stem preserved, extension changed)."""
    stem = Path(input_path).stem
    return stem + (".md" if direction == "docx_to_md" else ".docx")


def build_pandoc_command(
    pandoc_path: str,
    input_path: str,
    output_dir: str,
    output_filename: str,
    direction: str,
) -> list:
    """Assemble the pandoc argument list for the given conversion direction."""
    output_file = os.path.join(output_dir, output_filename)
    if direction == "docx_to_md":
        images_dir = os.path.join(output_dir, "images")
        return [
            pandoc_path,
            input_path,
            "-t", "markdown",
            "-o", output_file,
            f"--extract-media={images_dir}",
        ]
    else:  # md_to_docx
        return [
            pandoc_path,
            input_path,
            "-o", output_file,
        ]


def run_conversion(
    pandoc_path: str,
    input_path: str,
    output_dir: str,
    direction: str,
) -> tuple:
    """
    Run the pandoc conversion. Returns (success: bool, message: str).
    Never raises — all errors are captured and returned as messages.
    """
    output_filename = derive_output_filename(input_path, direction)
    cmd = build_pandoc_command(
        pandoc_path, input_path, output_dir, output_filename, direction
    )
    output_path = os.path.join(output_dir, output_filename)
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=CONVERSION_TIMEOUT,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode != 0:
            err = result.stderr.strip() or "Pandoc returned a non-zero exit code."
            return False, f"Pandoc error:\n{err}"
        return True, f"Output written to:\n{output_path}"
    except subprocess.TimeoutExpired:
        return False, f"Conversion timed out after {CONVERSION_TIMEOUT} seconds."
    except FileNotFoundError:
        return False, "Could not launch pandoc. Please check your installation."
    except Exception as exc:
        return False, f"Unexpected error: {exc}"


# ---------------------------------------------------------------------------
# GUI application
# ---------------------------------------------------------------------------


class PandocConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Pandoc Converter")
        self.root.minsize(540, 420)
        self.root.resizable(True, True)

        self.input_path: str = ""
        self.direction: str | None = None
        self.pandoc_path: str | None = None
        self.pandoc_missing: bool = False

        self._check_pandoc()
        self._build_ui()
        self._bind_events()

        # Show error after UI is fully built so the window appears first
        if self.pandoc_missing:
            self.root.after(
                100,
                lambda: messagebox.showerror(
                    "Pandoc Not Found",
                    "Pandoc could not be found on your system.\n\n"
                    "Please install pandoc (https://pandoc.org/installing.html) "
                    "and restart this application.",
                ),
            )

    # ------------------------------------------------------------------
    # Initialisation helpers
    # ------------------------------------------------------------------

    def _check_pandoc(self):
        self.pandoc_path = find_pandoc()
        self.pandoc_missing = self.pandoc_path is None

    def _build_ui(self):
        PAD = 12

        # ── Root grid configuration ──────────────────────────────────────
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # ── Outer frame ──────────────────────────────────────────────────
        outer = ttk.Frame(self.root, padding=PAD)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(6, weight=1)  # log area expands

        # ── Pandoc status badge (top-right) ──────────────────────────────
        status_text = "pandoc: found" if not self.pandoc_missing else "pandoc: NOT FOUND"
        status_colour = "#1a7a1a" if not self.pandoc_missing else "#c0392b"
        self._status_label = ttk.Label(
            outer, text=status_text, foreground=status_colour, font=("Segoe UI", 8)
        )
        self._status_label.grid(row=0, column=0, sticky="e", pady=(0, 6))

        # ── Input file section ────────────────────────────────────────────
        ttk.Label(outer, text="Input File", font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w"
        )

        file_frame = ttk.Frame(outer)
        file_frame.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        file_frame.columnconfigure(0, weight=1)

        self._file_entry = ttk.Entry(file_frame, state="readonly")
        self._file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self._browse_btn = ttk.Button(file_frame, text="Browse…", command=self._on_browse)
        self._browse_btn.grid(row=0, column=1)

        # ── Conversion direction label ────────────────────────────────────
        self._direction_label = ttk.Label(
            outer,
            text="Conversion:  —",
            font=("Segoe UI", 9),
            foreground="#555555",
        )
        self._direction_label.grid(row=3, column=0, sticky="w", pady=(10, 0))

        # ── Convert button ────────────────────────────────────────────────
        self._convert_btn = ttk.Button(
            outer,
            text="Convert",
            command=self._on_convert,
            state="disabled",
            width=18,
        )
        self._convert_btn.grid(row=4, column=0, pady=(12, 0))

        # ── Log section ───────────────────────────────────────────────────
        ttk.Label(outer, text="Log", font=("Segoe UI", 9, "bold")).grid(
            row=5, column=0, sticky="w", pady=(14, 2)
        )

        self._log_area = scrolledtext.ScrolledText(
            outer,
            state="disabled",
            font=("Consolas", 9),
            height=10,
            wrap="word",
            relief="flat",
            borderwidth=1,
        )
        self._log_area.grid(row=6, column=0, sticky="nsew")

        # Colour tags for log entries
        self._log_area.tag_config("success", foreground="#1a7a1a")
        self._log_area.tag_config("error", foreground="#c0392b")
        self._log_area.tag_config("info", foreground="#555555")

    def _bind_events(self):
        self.root.bind("<Return>", lambda _e: self._on_convert() if str(self._convert_btn["state"]) != "disabled" else None)

    # ------------------------------------------------------------------
    # File selection
    # ------------------------------------------------------------------

    def _on_browse(self):
        path = filedialog.askopenfilename(
            title="Select a file to convert",
            filetypes=[
                ("Supported files", "*.docx *.md"),
                ("Word documents", "*.docx"),
                ("Markdown files", "*.md"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        self.input_path = path
        self._file_entry.config(state="normal")
        self._file_entry.delete(0, "end")
        self._file_entry.insert(0, path)
        self._file_entry.config(state="readonly")
        self._update_direction_label()
        self._clear_log()

    def _update_direction_label(self):
        self.direction = detect_direction(self.input_path)
        if self.direction == "docx_to_md":
            text = "Conversion:  docx  →  markdown"
            colour = "#1a3a6b"
            can_convert = True
        elif self.direction == "md_to_docx":
            text = "Conversion:  markdown  →  docx"
            colour = "#1a3a6b"
            can_convert = True
        else:
            text = "Conversion:  Unsupported file type"
            colour = "#d68910"
            can_convert = False

        self._direction_label.config(text=text, foreground=colour)

        if can_convert and not self.pandoc_missing:
            self._convert_btn.config(state="normal")
        else:
            self._convert_btn.config(state="disabled")

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def _on_convert(self):
        if not self.input_path or not self.direction:
            return

        output_dir = filedialog.askdirectory(title="Choose output folder")
        if not output_dir:
            return  # user cancelled — do nothing

        self._convert_btn.config(state="disabled")
        self._log("Converting…", "info")

        thread = threading.Thread(
            target=self._run_conversion_thread,
            args=(output_dir,),
            daemon=True,
        )
        thread.start()

    def _run_conversion_thread(self, output_dir: str):
        success, message = run_conversion(
            self.pandoc_path,
            self.input_path,
            output_dir,
            self.direction,
        )
        self.root.after(0, lambda: self._on_conversion_done(success, message))

    def _on_conversion_done(self, success: bool, message: str):
        if success:
            self._log(message, "success")
        else:
            self._log(message, "error")

        if not self.pandoc_missing and self.direction:
            self._convert_btn.config(state="normal")

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def _log(self, text: str, tag: str = "info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        entry = f"[{timestamp}] {text}\n"
        self._log_area.config(state="normal")
        self._log_area.insert("end", entry, tag)
        self._log_area.see("end")
        self._log_area.config(state="disabled")

    def _clear_log(self):
        self._log_area.config(state="normal")
        self._log_area.delete("1.0", "end")
        self._log_area.config(state="disabled")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = PandocConverterApp(root)
    root.mainloop()

#!/usr/bin/env python3
"""Toyo Schedule Maker (Light) - GUI Installer

Light edition: no OCR / Import-from-Photo support, so this installer
only sets up the core venv with openpyxl + fpdf2. Launched automatically
by run.sh on first run or when deps are missing.
"""

import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path

VENV_DIR = Path.home() / ".toyo_scheduler" / "venv"
PYTHON = VENV_DIR / "bin" / "python3"
PIP = VENV_DIR / "bin" / "pip"


def check_venv():
    return PYTHON.exists()


def check_core_deps():
    """Check if openpyxl and fpdf2 are importable."""
    if not check_venv():
        return False
    try:
        result = subprocess.run(
            [str(PYTHON), "-c", "import openpyxl; from fpdf import FPDF"],
            capture_output=True, timeout=10)
        return result.returncode == 0
    except Exception:
        return False


class InstallerWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Toyo Schedule Maker (Light) - Setup")
        self.root.geometry("520x340")
        self.root.resizable(False, False)

        # Center on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 520) // 2
        y = (self.root.winfo_screenheight() - 340) // 2
        self.root.geometry(f"+{x}+{y}")

        self.cancelled = False
        self.install_success = False

        self._build_ui()
        self._check_status()

    def _build_ui(self):
        ttk.Label(self.root, text="Toyo Schedule Maker (Light)",
                  font=("Helvetica", 16, "bold")).pack(pady=(20, 5))
        ttk.Label(self.root, text="First-Time Setup",
                  font=("Helvetica", 11)).pack(pady=(0, 15))

        status_frame = ttk.LabelFrame(self.root, text="Components", padding=10)
        status_frame.pack(fill=tk.X, padx=20, pady=5)

        row1 = ttk.Frame(status_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Python Environment", width=25, anchor="w").pack(side=tk.LEFT)
        self.venv_status = ttk.Label(row1, text="Checking...", width=15)
        self.venv_status.pack(side=tk.RIGHT)

        row2 = ttk.Frame(status_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Core Dependencies", width=25, anchor="w").pack(side=tk.LEFT)
        self.core_status = ttk.Label(row2, text="Checking...", width=15)
        self.core_status.pack(side=tk.RIGHT)

        ttk.Label(self.root,
                  text="Light edition — Photo Import (OCR) is not included.",
                  font=("Helvetica", 9, "italic")).pack(pady=(10, 0), padx=20, anchor="w")

        self.progress_label = ttk.Label(self.root, text="", font=("Helvetica", 10))
        self.progress_label.pack(pady=(10, 3), padx=20, anchor="w")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=460)
        self.progress.pack(padx=20, pady=3)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=20, pady=(15, 10))

        self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self._cancel)
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)

        self.install_btn = ttk.Button(btn_frame, text="Install", command=self._start_install)
        self.install_btn.pack(side=tk.RIGHT, padx=5)

    def _set_status(self, label, installed):
        if installed:
            label.configure(text="Installed", foreground="green")
        else:
            label.configure(text="Not installed", foreground="red")

    def _check_status(self):
        has_venv = check_venv()
        has_core = check_core_deps()

        self._set_status(self.venv_status, has_venv)
        self._set_status(self.core_status, has_core)

        if has_venv and has_core:
            self.progress_label.configure(text="Everything is installed!")
            self.install_btn.configure(text="Launch App", command=self._launch)
            self.install_success = True

    def _cancel(self):
        self.cancelled = True
        self.root.destroy()

    def _launch(self):
        self.install_success = True
        self.root.destroy()

    def _start_install(self):
        self.install_btn.configure(state="disabled")
        self.cancel_btn.configure(state="disabled")
        self.progress.start(15)

        thread = threading.Thread(target=self._install_thread, daemon=True)
        thread.start()

    def _install_thread(self):
        try:
            if not check_venv():
                self._update_status("Creating Python environment...")
                VENV_DIR.parent.mkdir(parents=True, exist_ok=True)
                result = subprocess.run(
                    [sys.executable, "-m", "venv", str(VENV_DIR)],
                    capture_output=True, text=True)
                if result.returncode != 0:
                    self._show_error(f"Failed to create venv:\n{result.stderr}")
                    return
                self._update_label(self.venv_status, "Installed", "green")

            if not check_core_deps():
                self._update_status("Installing core dependencies...")
                result = subprocess.run(
                    [str(PIP), "install", "--quiet", "openpyxl", "fpdf2"],
                    capture_output=True, text=True, timeout=300)
                if result.returncode != 0:
                    self._show_error(f"Failed to install core deps:\n{result.stderr}")
                    return
                self._update_label(self.core_status, "Installed", "green")

            self._finish_success()

        except subprocess.TimeoutExpired:
            self._show_error("Installation timed out. Check your internet connection.")
        except Exception as e:
            self._show_error(str(e))

    def _update_status(self, text):
        self.root.after(0, lambda: self.progress_label.configure(text=text))

    def _update_label(self, label, text, color):
        self.root.after(0, lambda: label.configure(text=text, foreground=color))

    def _show_error(self, msg):
        def _do():
            self.progress.stop()
            self.progress_label.configure(text="Installation failed!", foreground="red")
            self.install_btn.configure(state="normal", text="Retry", command=self._start_install)
            self.cancel_btn.configure(state="normal")
            from tkinter import messagebox
            messagebox.showerror("Installation Error", msg)
        self.root.after(0, _do)

    def _finish_success(self):
        def _do():
            self.progress.stop()
            self.progress_label.configure(text="Installation complete!", foreground="green")
            self.install_btn.configure(state="normal", text="Launch App", command=self._launch)
            self.cancel_btn.configure(state="normal")
            self.install_success = True
        self.root.after(0, _do)

    def run(self):
        self.root.mainloop()
        return self.install_success


def needs_install():
    return not check_core_deps()


if __name__ == "__main__":
    app = InstallerWindow()
    if app.run():
        print("Ready to launch app.")
    else:
        print("Installation cancelled.")

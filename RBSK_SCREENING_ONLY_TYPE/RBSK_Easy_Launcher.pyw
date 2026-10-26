import argparse
import base64
import hashlib
import contextlib
import csv
import ctypes
import ctypes.wintypes
import io
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import uuid
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk


# ── Auto-Install Required Packages ────────────────────────────
def _auto_install_packages():
    """Check for all required third-party packages and install any
    that are missing.  Shows a tkinter dialog so the user knows
    what is happening.  Called once at startup before any third-party
    imports.

    Handles fresh-PC scenarios:
      - pip itself may be missing  → runs ensurepip first
      - .pyw runs under pythonw.exe (no console) → hides subprocess windows
      - os.execv is unreliable on Windows → uses subprocess.Popen + sys.exit
    """
    required = {
        "selenium": "selenium",
        "openpyxl": "openpyxl",
        "requests": "requests",
        "urllib3": "urllib3",
        "cryptography": "cryptography",
    }
    missing = []
    for import_name, pip_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        return  # everything already installed

    # Hide console windows when running under pythonw.exe
    _creation_flags = 0
    _startupinfo = None
    if sys.executable.lower().endswith("pythonw.exe"):
        _creation_flags = 0x08000000  # CREATE_NO_WINDOW
        _startupinfo = subprocess.STARTUPINFO()
        _startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        _startupinfo.wShowWindow = 0  # SW_HIDE

    def _pip_run(pip_args: list[str], timeout: int = 120) -> subprocess.CompletedProcess:
        """Run a pip command, hiding console windows on .pyw."""
        cmd = [sys.executable, "-m", "pip"] + pip_args
        return subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout,
            creationflags=_creation_flags, startupinfo=_startupinfo,
        )

    # ── Ensure pip itself is available ──
    try:
        _pip_run(["--version"], timeout=15)
    except Exception:
        # pip is missing — try ensurepip
        try:
            subprocess.run(
                [sys.executable, "-m", "ensurepip", "--upgrade"],
                capture_output=True, text=True, timeout=120,
                creationflags=_creation_flags, startupinfo=_startupinfo,
            )
        except Exception:
            pass
        # Verify pip is now available
        try:
            _pip_run(["--version"], timeout=15)
        except Exception:
            # pip still not available — show error and exit
            err_root = tk.Tk()
            err_root.withdraw()
            messagebox.showerror(
                "pip Not Available",
                "Python 'pip' is not installed and could not be set up automatically.\n\n"
                "Please run this command in Command Prompt:\n\n"
                "    python -m ensurepip --upgrade\n\n"
                "Then restart the application.",
            )
            err_root.destroy()
            sys.exit(1)

    # ── Show a small tkinter window while installing ──
    root = tk.Tk()
    root.title("RBSK - Installing Packages")
    root.geometry("480x220")
    root.resizable(False, False)
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    frame = tk.Frame(root, bg="#f4f8fb")
    frame.pack(fill="both", expand=True, padx=20, pady=20)
    tk.Label(
        frame, text="First-time setup — installing packages...",
        font=("Segoe UI", 13, "bold"), bg="#f4f8fb", fg="#18212f",
    ).pack(pady=(10, 5))
    status_var = tk.StringVar(value=f"Packages to install: {', '.join(missing)}")
    tk.Label(
        frame, textvariable=status_var,
        font=("Segoe UI", 10), bg="#f4f8fb", fg="#555555",
        wraplength=440, justify="center",
    ).pack(pady=5)
    progress_var = tk.StringVar(value="Please wait, this may take a minute...")
    tk.Label(
        frame, textvariable=progress_var,
        font=("Segoe UI", 10), bg="#f4f8fb", fg="#245eea",
    ).pack(pady=5)
    root.update()

    # ── Upgrade pip quietly first ──
    try:
        _pip_run(["install", "--upgrade", "pip", "--quiet"], timeout=60)
    except Exception:
        pass

    # ── Install each missing package ──
    installed = []
    failed = []
    for i, pkg in enumerate(missing, 1):
        status_var.set(f"Installing {pkg}... ({i}/{len(missing)})")
        root.update()
        try:
            result = _pip_run(["install", pkg, "--quiet"])
            if result.returncode == 0:
                installed.append(pkg)
            else:
                failed.append(pkg)
        except Exception:
            failed.append(pkg)

    root.destroy()

    if failed:
        err_root = tk.Tk()
        err_root.withdraw()
        messagebox.showerror(
            "Package Installation Failed",
            f"Could not install: {', '.join(failed)}\n\n"
            f"Please run this command manually in Command Prompt:\n\n"
            f"    python -m pip install {' '.join(failed)}\n\n"
            f"Then restart the application.",
        )
        err_root.destroy()
        sys.exit(1)

    if installed:
        # Restart the app so freshly installed packages can be imported.
        # os.execv is unreliable on Windows, so use Popen + sys.exit instead.
        subprocess.Popen(
            [sys.executable] + sys.argv,
            creationflags=_creation_flags,
            startupinfo=_startupinfo,
        )
        sys.exit(0)


_auto_install_packages()

import Students

try:
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.primitives.padding import PKCS7
except Exception:
    Cipher = algorithms = modes = PKCS7 = None


# ── DPI Awareness ──────────────────────────────────────────────
def _enable_dpi_awareness():
    """Enable per-monitor DPI awareness on Windows so the app
    sees the real screen resolution instead of a scaled one."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

_enable_dpi_awareness()


# ── Screen Scaling Helpers ─────────────────────────────────────
REFERENCE_WIDTH = 1920
REFERENCE_HEIGHT = 1080


def compute_scale_factor(screen_w: int, screen_h: int) -> float:
    """Compute a UI scale factor relative to 1920x1080.
    Clamp between 0.6 and 2.0 so the UI stays usable."""
    sx = screen_w / REFERENCE_WIDTH
    sy = screen_h / REFERENCE_HEIGHT
    factor = min(sx, sy)
    return max(0.6, min(factor, 2.0))


def scaled_geometry(base_w: int, base_h: int, factor: float,
                    screen_w: int, screen_h: int) -> str:
    """Return a centered WxH+X+Y geometry string scaled to the screen."""
    w = min(int(base_w * factor), screen_w - 40)
    h = min(int(base_h * factor), screen_h - 80)
    x = max(0, (screen_w - w) // 2)
    y = max(0, (screen_h - h) // 2)
    return f"{w}x{h}+{x}+{y}"


def scaled_font_size(base_size: int, factor: float) -> int:
    """Scale a font size, keeping a minimum of 7."""
    return max(7, int(base_size * factor))


# ── Auto-Install Helper ────────────────────────────────────────
def _try_install_cryptography() -> bool:
    """Attempt to pip-install cryptography. Returns True on success."""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "cryptography", "--quiet"],
            capture_output=True, text=True, timeout=120,
        )
        return result.returncode == 0
    except Exception:
        return False


def check_cryptography_at_startup():
    """Show a user-friendly dialog if cryptography is missing,
    with an option to auto-install it."""
    if Cipher is not None:
        return True  # already available
    root = tk.Tk()
    root.withdraw()
    answer = messagebox.askyesnocancel(
        "Missing Package: cryptography",
        "The Python 'cryptography' package is required for\n"
        "licence activation but is not installed on this PC.\n\n"
        "Click YES to install it automatically now.\n"
        "Click NO to see the manual install command.\n"
        "Click CANCEL to exit.",
        icon="warning",
    )
    if answer is True:
        messagebox.showinfo(
            "Installing...",
            "Installing cryptography package...\n\n"
            "This may take a minute. Click OK and please wait.",
        )
        if _try_install_cryptography():
            messagebox.showinfo(
                "Installed Successfully",
                "cryptography package installed!\n\n"
                "The application will now restart.",
            )
            root.destroy()
            os.execv(sys.executable, [sys.executable] + sys.argv)
        else:
            messagebox.showerror(
                "Installation Failed",
                "Could not install cryptography automatically.\n\n"
                "Please run this command manually in Command Prompt:\n\n"
                "    python -m pip install cryptography\n\n"
                "Then restart the application.",
            )
            root.destroy()
            sys.exit(1)
    elif answer is False:
        messagebox.showinfo(
            "Manual Install",
            "Open Command Prompt and run:\n\n"
            "    python -m pip install cryptography\n\n"
            "Then restart the application.",
        )
        root.destroy()
        sys.exit(1)
    else:
        root.destroy()
        sys.exit(0)


APP_DIR = Path(__file__).resolve().parent
SAVED_LOGIN_FILE = APP_DIR / "rbsk_saved_login.json"
LICENSE_FILE = APP_DIR / "rbsk_license.json"
BROWSER_PROFILE_DIR = APP_DIR / "rbsk_browser_profile"
BROWSER_PROFILE_ROOT = APP_DIR / "rbsk_browser_profiles"
LICENSE_AES_KEY = b"RBSKScreen2026SecureKey@#1234567"
LICENSE_AES_IV = b"InitVector123456"
DEFAULT_LICENSE_LIMIT = 10000
SAFE_UPDATE_PROGRAM_FILES = (
    "RBSK_Easy_Launcher.pyw",
    "Students.py",
    "RUN_EXCEL_SCREENING.py",
    "RUN_SCREENING_NOW.py",
    "OPEN_RBSK_SCREENING_APP.bat",
    "RUN_SCREENING_ONLY.bat",
    "ONE_CLICK_INSTALL_AND_RUN.bat",
    "SETUP_RBSK.bat",
    "INSTALL_REQUIRED_FILES.bat",
    "CREATE_DESKTOP_SHORTCUT.bat",
    "create_shortcut.ps1",
    "README_1CLICK_SETUP.txt",
    "CREATE_SAFE_UPDATE_FOLDER.bat",
    "create_safe_update_folder.ps1",
    "SAFE_UPDATE_INSTRUCTIONS.txt",
)


def destination_license_summary(target_dir: Path) -> dict:
    license_path = Path(target_dir) / "rbsk_license.json"
    if not license_path.is_file():
        raise Students.ImportErrorWithHelp(
            "The selected folder has no existing rbsk_license.json. Select the installed RBSK folder on the other PC."
        )
    try:
        data = json.loads(license_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise Students.ImportErrorWithHelp("The destination PC licence file cannot be read. Update was not started.") from exc
    if not Students.clean(data.get("activation_key")):
        raise Students.ImportErrorWithHelp("The destination PC does not have an activated licence. Update was not started.")
    limit = int(data.get("entry_limit") or DEFAULT_LICENSE_LIMIT)
    used = max(int(data.get("successful_entries") or 0), 0)
    return {
        "machine_id": Students.clean(data.get("machine_id")) or "Existing PC",
        "expires_on": Students.clean(data.get("expires_on")) or "Stored in activation key",
        "entry_limit": limit,
        "successful_entries": used,
        "remaining_entries": max(limit - used, 0),
    }


def safe_update_installation(source_dir: Path, target_dir: Path) -> dict:
    source = Path(source_dir).resolve()
    target = Path(target_dir).resolve()
    if source == target:
        raise Students.ImportErrorWithHelp("Select the RBSK installation folder on the other PC, not this current folder.")
    if not target.is_dir():
        raise Students.ImportErrorWithHelp("The selected destination folder does not exist.")

    summary = destination_license_summary(target)
    license_path = target / "rbsk_license.json"
    original_license = license_path.read_bytes()
    copied: list[str] = []
    try:
        for name in SAFE_UPDATE_PROGRAM_FILES:
            source_file = source / name
            if not source_file.is_file():
                continue
            temporary = target / f".{name}.update_tmp"
            shutil.copy2(source_file, temporary)
            os.replace(temporary, target / name)
            copied.append(name)
    finally:
        for temporary in target.glob(".*.update_tmp"):
            try:
                temporary.unlink()
            except Exception:
                pass

    if license_path.read_bytes() != original_license:
        license_path.write_bytes(original_license)
        raise Students.ImportErrorWithHelp("Destination licence changed during update and was restored. Update stopped.")
    if not copied:
        raise Students.ImportErrorWithHelp("No application program files were found to update.")
    return {**summary, "copied_files": copied}
DEFAULT_BATCH_CONTROL = APP_DIR / "sample_screening_health_students.xlsx"
if not DEFAULT_BATCH_CONTROL.exists():
    DEFAULT_BATCH_CONTROL = APP_DIR / "final_formatted_data_for_form.xlsx"
if not DEFAULT_BATCH_CONTROL.exists():
    DEFAULT_BATCH_CONTROL = APP_DIR / "batch_control_real_data.csv"
if not DEFAULT_BATCH_CONTROL.exists():
    DEFAULT_BATCH_CONTROL = APP_DIR / "batch_control_sample.csv"


def screening_month_options() -> tuple[str, ...]:
    today = date.today()
    start_year = today.year if today.month >= 4 else today.year - 1
    month_names = (
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    )
    values = [""]
    for offset in range(12):
        month_number = ((3 + offset) % 12) + 1
        year = start_year if month_number >= 4 else start_year + 1
        values.append(f"{month_names[month_number - 1]} {year}")
    return tuple(values)


def app_machine_id() -> str:
    raw = f"{uuid.getnode():012X}|{Path.home()}".upper()
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()
    return f"RBSK-{digest[:4]}-{digest[4:8]}-{digest[8:12]}-{digest[12:16]}"


def decrypt_license_key(activation_key: str) -> dict:
    if Cipher is None or algorithms is None or modes is None or PKCS7 is None:
        raise Students.ImportErrorWithHelp("Python cryptography package is required for licence activation.")
    try:
        encrypted = base64.b64decode(Students.clean(activation_key))
        cipher = Cipher(algorithms.AES(LICENSE_AES_KEY), modes.CBC(LICENSE_AES_IV))
        decryptor = cipher.decryptor()
        padded = decryptor.update(encrypted) + decryptor.finalize()
        unpadder = PKCS7(128).unpadder()
        payload = (unpadder.update(padded) + unpadder.finalize()).decode("utf-8")
    except Exception as exc:
        raise Students.ImportErrorWithHelp("Activation key is invalid or damaged.") from exc
    parts = [part.strip() for part in payload.split("|")]
    if len(parts) < 2:
        raise Students.ImportErrorWithHelp("Activation key payload is invalid.")
    limit = DEFAULT_LICENSE_LIMIT
    if len(parts) >= 3 and parts[2]:
        try:
            limit = int(parts[2])
        except ValueError as exc:
            raise Students.ImportErrorWithHelp("Activation key entry limit is invalid.") from exc
    return {
        "machine_id": parts[0].upper(),
        "expires_on": parts[1],
        "entry_limit": limit,
    }


class TextRedirector(io.StringIO):
    def __init__(self, output_queue: queue.Queue[str]):
        super().__init__()
        self.output_queue = output_queue

    def write(self, text: str) -> int:
        if text:
            self.output_queue.put(text)
        return len(text)

    def flush(self) -> None:
        pass


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def protect_text_for_windows_user(text: str) -> str:
    data = text.encode("utf-8")
    in_buffer = ctypes.create_string_buffer(data)
    in_blob = DATA_BLOB(len(data), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
    out_blob = DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        "RBSK Student Entry Helper",
        None,
        None,
        None,
        0x01,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    return base64.b64encode(encrypted).decode("ascii")


def unprotect_text_for_windows_user(token: str) -> str:
    encrypted = base64.b64decode(token.encode("ascii"))
    in_buffer = ctypes.create_string_buffer(encrypted)
    in_blob = DATA_BLOB(len(encrypted), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
    out_blob = DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0x01,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise ctypes.WinError()
    try:
        plain = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    return plain.decode("utf-8")


class EasyLauncher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RBSK Screening Health Only")

        # ── Screen-aware scaling ──
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        self.ui_scale = compute_scale_factor(screen_w, screen_h)
        self.geometry(scaled_geometry(1180, 820, self.ui_scale, screen_w, screen_h))
        self.minsize(
            min(int(980 * self.ui_scale), screen_w - 40),
            min(int(700 * self.ui_scale), screen_h - 80),
        )

        self.output_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.dashboard_refresh_job = None
        self.dashboard_rows: list[dict] = []
        self.dashboard_completed: set[str] = set()
        self.dashboard_failed_this_run: set[str] = set()
        self.dashboard_current_awc = "-"
        self.dashboard_mapping_status = "Not checked"
        self.dashboard_error = ""
        self.current_run_only_sr: str | None = None
        self.current_run_label = "Current run: all pending rows"
        self.current_run_snapshot: dict | None = None
        self.current_awc_done = 0
        self.current_awc_total = 0
        self.run_timing_started_at: float | None = None
        self.run_timing_done = 0
        self.run_timing_total = 0
        self.run_timing_row_seconds: list[float] = []
        self.run_timing_text = "-"
        self.log_rows: list[dict] = []
        self.active_browser_drivers = []
        self.popup_browser_driver = None
        self.saved_login_data = self.load_saved_login_data()
        self.license_lock = threading.Lock()

        self._configure_style()
        self._build_ui()
        self.after(300, self.refresh_dashboard)
        self.after(450, self.refresh_logs)
        self.after(150, self._drain_output)

    def load_saved_login_data(self) -> dict:
        if not SAVED_LOGIN_FILE.exists():
            return {"last_login_id": "", "accounts": {}, "profile_overrides": {}}
        try:
            data = json.loads(SAVED_LOGIN_FILE.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return {"last_login_id": "", "accounts": {}, "profile_overrides": {}}
            data.setdefault("last_login_id", "")
            data.setdefault("accounts", {})
            data["profile_overrides"] = {}
            return data
        except Exception:
            return {"last_login_id": "", "accounts": {}, "profile_overrides": {}}

    def load_license_data(self) -> dict:
        if not LICENSE_FILE.exists():
            return {}
        try:
            data = json.loads(LICENSE_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def save_license_data(self, data: dict) -> None:
        LICENSE_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def parse_license_date(self, value: str) -> date:
        try:
            return datetime.strptime(Students.clean(value), "%Y-%m-%d").date()
        except Exception as exc:
            raise Students.ImportErrorWithHelp("Stored licence expiry date is invalid. Activate a new key.") from exc

    def validate_license_data(self, data: dict) -> dict:
        if not data.get("activation_key"):
            raise Students.ImportErrorWithHelp(f"Licence is not activated. Machine ID: {app_machine_id()}")
        payload = decrypt_license_key(data.get("activation_key", ""))
        machine_id = app_machine_id()
        if payload["machine_id"] != machine_id:
            raise Students.ImportErrorWithHelp(f"Licence key is for another machine.\nThis Machine ID: {machine_id}")
        expires_on = self.parse_license_date(payload["expires_on"])
        if date.today() > expires_on:
            raise Students.ImportErrorWithHelp(f"Licence expired on {expires_on:%Y-%m-%d}. Activate a new monthly key.")
        limit = int(payload.get("entry_limit") or DEFAULT_LICENSE_LIMIT)
        used = int(data.get("successful_entries") or 0)
        if used >= limit:
            raise Students.ImportErrorWithHelp(
                f"Monthly licence limit reached: {used}/{limit} successful entries. Activate a new key to continue."
            )
        data["machine_id"] = machine_id
        data["expires_on"] = expires_on.strftime("%Y-%m-%d")
        data["entry_limit"] = limit
        data["successful_entries"] = used
        return data

    def license_status_text(self) -> str:
        try:
            data = self.validate_license_data(self.load_license_data())
            remaining = max(int(data["entry_limit"]) - int(data["successful_entries"]), 0)
            return f"Licence active until {data['expires_on']} | {remaining}/{data['entry_limit']} entries left"
        except Exception as exc:
            return f"Licence needed: {exc}"

    def refresh_license_status(self) -> None:
        if hasattr(self, "license_status"):
            self.license_status.set(self.license_status_text())

    def copy_machine_id(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(app_machine_id())
        self.show_guided_dialog("Machine ID Copied", app_machine_id(), wait=False)

    def activate_license_from_ui(self) -> None:
        key = Students.clean(self.license_key.get() if hasattr(self, "license_key") else "")
        if not key:
            messagebox.showwarning("Activation key required", "Paste the monthly activation key first.")
            return
        try:
            payload = decrypt_license_key(key)
            machine_id = app_machine_id()
            if payload["machine_id"] != machine_id:
                raise Students.ImportErrorWithHelp(f"Key is for {payload['machine_id']}.\nThis Machine ID is {machine_id}.")
            expires_on = self.parse_license_date(payload["expires_on"])
            if date.today() > expires_on:
                raise Students.ImportErrorWithHelp(f"This key expired on {expires_on:%Y-%m-%d}.")
            limit = int(payload.get("entry_limit") or DEFAULT_LICENSE_LIMIT)
            self.save_license_data(
                {
                    "activation_key": key,
                    "machine_id": machine_id,
                    "activated_on": date.today().strftime("%Y-%m-%d"),
                    "expires_on": expires_on.strftime("%Y-%m-%d"),
                    "entry_limit": limit,
                    "successful_entries": 0,
                    "last_success_on": "",
                }
            )
            self.license_key.set("")
            self.refresh_license_status()
            self.show_guided_dialog(
                "Licence Activated",
                f"Licence activated until {expires_on:%Y-%m-%d}.\nSuccessful entry limit: {limit}.",
                wait=False,
            )
        except Exception as exc:
            messagebox.showerror("Licence Activation Failed", str(exc))

    def ensure_license_available(self) -> None:
        if self.dry_run.get() if hasattr(self, "dry_run") else False:
            return
        with self.license_lock:
            self.validate_license_data(self.load_license_data())

    def record_license_success(self, count: int = 1) -> None:
        if self.dry_run.get() if hasattr(self, "dry_run") else False:
            return
        with self.license_lock:
            data = self.validate_license_data(self.load_license_data())
            data["successful_entries"] = int(data.get("successful_entries") or 0) + int(count or 1)
            data["last_success_on"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.save_license_data(data)
        self.after(0, self.refresh_license_status)

    def retain_browser_driver(self, driver) -> None:
        if driver is None:
            return
        if all(existing is not driver for existing in self.active_browser_drivers):
            self.active_browser_drivers.append(driver)
        print("RBSK browser kept open for this run only. Session/profile reuse is disabled.")

    def reusable_browser_driver(self):
        self.active_browser_drivers = []
        return None

    def save_login_password(self, login_id: str, password: str) -> None:
        login_id = Students.clean(login_id)
        if not login_id or not password:
            return
        data = self.load_saved_login_data()
        accounts = data.setdefault("accounts", {})
        accounts[login_id] = protect_text_for_windows_user(password)
        data["last_login_id"] = login_id
        SAVED_LOGIN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.saved_login_data = data

    def saved_password_for(self, login_id: str) -> str | None:
        login_id = Students.clean(login_id)
        token = (self.saved_login_data.get("accounts") or {}).get(login_id)
        if not token:
            return None
        try:
            return unprotect_text_for_windows_user(token)
        except Exception:
            return None

    def password_for_run(self, args: argparse.Namespace, login_id: str) -> str:
        entered_login_id = Students.clean(getattr(args, "entered_password_login_id", ""))
        entered_password = getattr(args, "entered_password", "") or ""
        if entered_password and entered_login_id == Students.clean(login_id):
            print("Using password entered on the main screen. You only need to enter CAPTCHA in RBSK.")
            return entered_password
        saved_password = self.saved_password_for(login_id)
        if saved_password:
            print("Using saved password. You only need to enter CAPTCHA in RBSK.")
            return saved_password
        raise Students.ImportErrorWithHelp(
            "Password is not available. Enter it in the Password field beside Login ID on the main screen, then click Start again."
        )

    def has_saved_password_token(self, login_id: str) -> bool:
        login_id = Students.clean(login_id)
        return bool(login_id and (self.saved_login_data.get("accounts") or {}).get(login_id))

    def saved_login_status_text(self, login_id: str) -> str:
        login_id = Students.clean(login_id)
        if not login_id:
            return "Password status: enter Login ID"
        if self.has_saved_password_token(login_id):
            if self.saved_password_for(login_id):
                return "Password saved. Next run: CAPTCHA only."
            return "Saved password could not be read. Enter it again once."
        return "No password saved yet. First run asks once."

    def refresh_saved_login_status(self) -> None:
        if hasattr(self, "saved_login_status"):
            if hasattr(self, "login_password") and self.login_password.get():
                self.saved_login_status.set("Password entered. Browser will ask for CAPTCHA only.")
            else:
                self.saved_login_status.set(self.saved_login_status_text(self.default_login_id.get()))

    def save_password_from_form(self) -> None:
        login_id = Students.clean(self.default_login_id.get())
        password = self.login_password.get()
        if not login_id:
            self.saved_login_status.set("Enter Login ID before saving password.")
            return
        if not password:
            self.saved_login_status.set("Enter Password before clicking Save Password.")
            return
        try:
            self.save_login_password(login_id, password)
        except Exception as exc:
            self.saved_login_status.set(f"Password was not saved: {exc}")
            return
        self.remember_credentials.set(True)
        self.login_password.set("")
        self.saved_login_status.set("Password saved securely. Next run: CAPTCHA only.")

    def forget_saved_password(self) -> None:
        login_id = Students.clean(self.default_login_id.get()) or getattr(self.saved_login_data, "get", lambda k, d=None: {})("last_login_id", "") or "MHT-1271251"
        data = self.load_saved_login_data()
        accounts = data.setdefault("accounts", {})
        if not accounts:
            if messagebox.askyesno(
                "No saved password",
                "No saved password token was found.\n\nReset the browser session and open the RBSK login window now?",
            ):
                self.reset_browser_session_for_login(confirm=False, open_login_after=True)
            return
        if not messagebox.askyesno(
            "Clear saved password",
            f"Clear all saved passwords / credentials for {login_id}?\n\n"
            "The browser session will also be reset so the RBSK login window opens fresh.",
        ):
            return
        data["accounts"] = {}
        data["last_login_id"] = login_id
        SAVED_LOGIN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.saved_login_data = data
        self.refresh_saved_login_status()
        messagebox.showinfo("Saved password cleared", "Saved passwords were cleared. Opening a fresh RBSK login window next.")
        self.reset_browser_session_for_login(confirm=False, open_login_after=True)

    def browser_profile_dir_for_login(self, login_id: str) -> Path:
        login_id = Students.clean(login_id) or Students.clean(self.default_login_id.get()) or "default"
        slug = Students.slugify(login_id)
        candidates: list[Path] = []
        override = (self.saved_login_data.get("profile_overrides") or {}).get(login_id)
        if override and Path(override).exists():
            candidates.append(Path(override))
        candidates.extend(path for path in BROWSER_PROFILE_ROOT.glob(f"{slug}*") if path.is_dir() and path.name != "_archived")
        if candidates:
            return max(candidates, key=lambda path: path.stat().st_mtime)
        return BROWSER_PROFILE_ROOT / slug

    def set_browser_profile_override(self, login_id: str, profile_dir: Path) -> None:
        login_id = Students.clean(login_id)
        if not login_id:
            return
        data = self.load_saved_login_data()
        data.setdefault("accounts", {})
        data["profile_overrides"] = {}
        data["last_login_id"] = login_id
        SAVED_LOGIN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.saved_login_data = data

    def clear_browser_profile_override(self, login_id: str) -> None:
        login_id = Students.clean(login_id)
        data = self.load_saved_login_data()
        overrides = data.setdefault("profile_overrides", {})
        overrides.pop(login_id, None)
        SAVED_LOGIN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.saved_login_data = data

    def make_recovery_profile_dir(self, login_id: str) -> Path:
        slug = Students.slugify(login_id or "default")
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        return BROWSER_PROFILE_ROOT / f"{slug}_recovery_{timestamp}"

    def is_browser_profile_start_error(self, exc: Exception) -> bool:
        text = str(exc).lower()
        return any(
            needle in text
            for needle in [
                "devtoolsactiveport",
                "session not created",
                "chrome failed to start",
                "user data directory is already in use",
                "cannot create default profile directory",
                "chrome not reachable",
            ]
        )

    def open_login_window_for_login(self, login_id: str, profile_dir: Path | None = None) -> None:
        login_id = Students.clean(login_id) or Students.clean(self.default_login_id.get()) or "default"
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Please wait", "A process is already running.")
            return

        def task():
            selected_profile_dir = None
            print(f"Opening RBSK login window for {login_id}...")
            driver = Students.start_browser(self.browser.get(), selected_profile_dir)
            self.retain_browser_driver(driver)
            driver.get(Students.APP_URL)
            if login_id != "default":
                time.sleep(1.5)
                try:
                    driver.execute_script(
                        """
                        const loginId = arguments[0];
                        const visible = el => {
                          const style = getComputedStyle(el);
                          const rect = el.getBoundingClientRect();
                          return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
                        };
                        const inputs = Array.from(document.querySelectorAll('input')).filter(visible);
                        const userInput = inputs.find(el => {
                          const hay = `${el.name || ''} ${el.id || ''} ${el.placeholder || ''} ${el.autocomplete || ''} ${el.type || ''}`.toLowerCase();
                          return el.type !== 'password' && ['user', 'login', 'mht', 'id', 'name'].some(hint => hay.includes(hint));
                        }) || inputs.find(el => el.type !== 'password');
                        if (userInput) {
                          userInput.focus();
                          const setter = Object.getOwnPropertyDescriptor(userInput.__proto__, 'value')?.set;
                          if (setter) setter.call(userInput, loginId);
                          else userInput.value = loginId;
                          userInput.dispatchEvent(new Event('input', {bubbles: true}));
                          userInput.dispatchEvent(new Event('change', {bubbles: true}));
                          return true;
                        }
                        return false;
                        """,
                        login_id,
                    )
                except Exception as exc:
                    print(f"Could not auto-fill Login ID: {exc}")
            print("RBSK login window is open. Enter password and CAPTCHA manually.")

        self.start_worker("Opening RBSK login", task)

    def reset_browser_session_for_login(self, confirm: bool = True, open_login_after: bool = True) -> None:
        login_id = Students.clean(self.default_login_id.get()) or getattr(self.saved_login_data, "get", lambda k, d=None: {})("last_login_id", "") or "default"
        if confirm and not messagebox.askyesno(
            "Reset browser session",
            f"This will archive saved Chrome/Edge browser sessions and reset browser memory.\n\n"
            "After reset, the RBSK login window will open automatically. Continue?",
        ):
            return
        for driver in list(self.active_browser_drivers):
            try:
                driver.quit()
            except Exception:
                pass
        self.active_browser_drivers = []
        try:
            Students._SHARED_BROWSER_DRIVER = None
        except Exception:
            pass
        archived: list[str] = []
        errors: list[str] = []
        archive_root = BROWSER_PROFILE_ROOT / "_archived"
        archive_root.mkdir(parents=True, exist_ok=True)
        
        candidates = set(BROWSER_PROFILE_ROOT.glob("*"))
        candidates.add(BROWSER_PROFILE_DIR)
        
        for path in candidates:
            if not path.is_dir() or path.name == "_archived":
                continue
            try:
                target = archive_root / f"{path.name}_{time.strftime('%Y%m%d_%H%M%S')}"
                shutil.move(str(path), str(target))
                archived.append(str(target))
            except Exception as exc:
                # If locked by running Chrome, try removing temporary subfolders or ignoring
                errors.append(f"{path.name}: {exc}")
                
        self.clear_browser_profile_override(login_id)
        self.saved_login_data = self.load_saved_login_data()
        self.refresh_saved_login_status()
        
        if errors:
            self.show_guided_dialog(
                "Browser Session Reset",
                "Browser session reset performed.\nNote: If Chrome is currently open, close Chrome windows to fully release files.\n\n"
                + "\n".join(errors),
                wait=False,
            )
        else:
            self.show_guided_dialog(
                "Browser Session Reset",
                "Browser session archived and reset successfully.\n\nOpening the RBSK login window now.",
                wait=False,
            )
        if open_login_after:
            fresh_profile_dir = BROWSER_PROFILE_ROOT / Students.slugify(login_id)
            if fresh_profile_dir.exists():
                fresh_profile_dir = self.make_recovery_profile_dir(login_id)
            self.open_login_window_for_login(login_id, profile_dir=fresh_profile_dir)

    def remember_login_id_only(self, login_id: str) -> None:
        login_id = Students.clean(login_id)
        if not login_id:
            return
        data = self.load_saved_login_data()
        data["last_login_id"] = login_id
        data.setdefault("accounts", {})
        SAVED_LOGIN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.saved_login_data = data

    def _configure_style(self) -> None:
        s = self.ui_scale  # shorthand for scale factor
        self.configure(bg="#f4f8fb")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f8fb")
        style.configure("Panel.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabel", background="#f4f8fb", foreground="#18212f", font=("Segoe UI", scaled_font_size(10, s)))
        style.configure("Panel.TLabel", background="#ffffff", foreground="#18212f", font=("Segoe UI", scaled_font_size(10, s)))
        style.configure("Header.TLabel", background="#22c3be", foreground="#041f2a", font=("Segoe UI", scaled_font_size(18, s), "bold"))
        style.configure("SubHeader.TLabel", background="#22c3be", foreground="#041f2a", font=("Segoe UI", scaled_font_size(10, s)))
        style.configure("Title.TLabel", background="#ffffff", foreground="#18212f", font=("Segoe UI", scaled_font_size(13, s), "bold"))
        style.configure("TButton", font=("Segoe UI", scaled_font_size(10, s), "bold"), padding=(int(12 * s), int(8 * s)))
        style.configure("Primary.TButton", background="#245eea", foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#194bc2")])
        style.configure("Danger.TButton", background="#8f2b2b", foreground="#ffffff")
        style.configure("TNotebook", background="#f4f8fb", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(int(14 * s), int(8 * s)), font=("Segoe UI", scaled_font_size(10, s), "bold"))
        style.configure("Dashboard.Treeview", font=("Segoe UI", scaled_font_size(9, s)), rowheight=max(18, int(24 * s)), background="#ffffff", fieldbackground="#ffffff")
        style.configure("Dashboard.Treeview.Heading", font=("Segoe UI", scaled_font_size(9, s), "bold"))
        style.configure("Horizontal.TProgressbar", troughcolor="#e5e7eb", background="#2563eb", thickness=max(8, int(14 * s)))

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg="#22c3be", height=86)
        header.pack(fill="x")
        header.pack_propagate(False)
        ttk.Label(header, text="RBSK Screening Health Only", style="Header.TLabel").pack(anchor="w", padx=24, pady=(16, 0))
        ttk.Label(
            header,
            text="Only existing students: Screening > School > Student > Health/Doctor > Preview > Submit",
            style="SubHeader.TLabel",
        ).pack(anchor="w", padx=24, pady=(3, 0))

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=18, pady=16)

        self.notebook = ttk.Notebook(body)
        self.notebook.pack(fill="both", expand=True)

        self.batch_tab = ttk.Frame(self.notebook, padding=16)
        self.dashboard_tab = ttk.Frame(self.notebook, padding=16)
        self.logs_tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(self.batch_tab, text="1. Screening Health")
        self.notebook.add(self.dashboard_tab, text="2. Dashboard")
        self.notebook.add(self.logs_tab, text="3. Logs")

        self._build_batch_tab()
        self._build_dashboard(self.dashboard_tab)
        self._build_logs_tab()
        self.notebook.select(self.batch_tab)

        output_panel = ttk.Frame(body, style="Panel.TFrame", padding=12)
        output_panel.pack(fill="x", expand=False, pady=(14, 0))
        ttk.Label(output_panel, text="Progress Log", style="Title.TLabel").pack(anchor="w")
        self.output = tk.Text(output_panel, height=7, wrap="word", font=("Consolas", 10), bg="#0f172a", fg="#e5eefb")
        self.output.pack(fill="x", expand=False, pady=(8, 0))

    def _build_demo_tab(self) -> None:
        ttk.Label(self.demo_tab, text="Safe local demo. This never submits to RBSK.", style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        self.demo_excel = tk.StringVar(value=str(APP_DIR / "students_sample_data.xlsx"))
        self.demo_limit = tk.StringVar(value="2")
        self.demo_keep_open = tk.BooleanVar(value=True)

        self._file_row(self.demo_tab, 1, "Student Excel", self.demo_excel, [("Excel/CSV", "*.xlsx *.csv"), ("All files", "*.*")])
        ttk.Label(self.demo_tab, text="Rows to show").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(self.demo_tab, textvariable=self.demo_limit, width=10).grid(row=2, column=1, sticky="w", pady=8)
        ttk.Checkbutton(self.demo_tab, text="Keep browser open after demo", variable=self.demo_keep_open).grid(row=3, column=0, columnspan=3, sticky="w", pady=8)
        ttk.Button(self.demo_tab, text="Start Local Test", style="Primary.TButton", command=self.start_demo).grid(row=4, column=0, sticky="w", pady=(12, 0))

        self.demo_tab.columnconfigure(1, weight=1)

    def _build_batch_tab(self) -> None:
        ttk.Label(self.batch_tab, text="Screening Health entry for already-added students only.", style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        self.batch_control = tk.StringVar(value=str(DEFAULT_BATCH_CONTROL))
        self.range_start = tk.StringVar(value="")
        self.batch_limit = tk.StringVar(value="")
        self.run_all_pending = tk.BooleanVar(value=True)
        self.only_batch = tk.StringVar(value="")
        self.dry_run = tk.BooleanVar(value=False)
        self.stop_on_error = tk.BooleanVar(value=False)
        self.allow_insecure_ssl = tk.BooleanVar(value=True)
        self.default_login_id = tk.StringVar(value=Students.clean(self.saved_login_data.get("last_login_id")) or "MHT-1271251")
        self.login_password = tk.StringVar(value="")
        self.manual_login = tk.BooleanVar(value=False)
        self.remember_credentials = tk.BooleanVar(value=True)
        self.use_browser_profile = tk.BooleanVar(value=False)
        self.no_store_session = tk.BooleanVar(value=True)
        self.browser_form_fill = tk.BooleanVar(value=True)
        self.auto_select_institute = tk.BooleanVar(value=True)
        self.manual_form_open = tk.BooleanVar(value=True)
        self.unattended_after_start = tk.BooleanVar(value=True)
        self.screening_health_entry = tk.BooleanVar(value=True)
        self.screening_month = tk.StringVar(value="")
        self.browser = tk.StringVar(value="chrome")
        self.speed_mode = tk.StringVar(value="Fast Safe")
        self.saved_login_status = tk.StringVar(value="")
        self.license_key = tk.StringVar(value="")
        self.license_status = tk.StringVar(value="")
        self.mobile_source = tk.StringVar(value="portal_school")
        self.custom_school_mobile = tk.StringVar(value="")
        self.screening_institution_type = tk.StringVar(value="School")

        self._file_row(self.batch_tab, 1, "Batch Control or Student Excel", self.batch_control, [("Excel/CSV", "*.xlsx *.csv"), ("All files", "*.*")])
        ttk.Label(self.batch_tab, text="Only Batch Name").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(self.batch_tab, textvariable=self.only_batch).grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Label(self.batch_tab, text="Login ID").grid(row=3, column=0, sticky="w", pady=8)
        login_fields = ttk.Frame(self.batch_tab)
        login_fields.grid(row=3, column=1, sticky="w", pady=8)
        ttk.Entry(login_fields, textvariable=self.default_login_id, width=24).pack(side="left")
        ttk.Label(login_fields, text="Password").pack(side="left", padx=(16, 6))
        ttk.Entry(login_fields, textvariable=self.login_password, width=28, show="*").pack(side="left")
        ttk.Button(login_fields, text="Save Password", command=self.save_password_from_form).pack(side="left", padx=(8, 0))
        login_status = ttk.Frame(self.batch_tab)
        login_status.grid(row=3, column=2, sticky="w", padx=(8, 0), pady=8)
        ttk.Label(login_status, textvariable=self.saved_login_status).pack(side="left")
        ttk.Label(self.batch_tab, text="Start Sr No").grid(row=4, column=0, sticky="w", pady=8)
        range_frame = ttk.Frame(self.batch_tab)
        range_frame.grid(row=4, column=1, sticky="w", pady=8)
        ttk.Entry(range_frame, textvariable=self.range_start, width=10).pack(side="left")
        ttk.Label(range_frame, text="  Number of rows").pack(side="left", padx=(14, 6))
        ttk.Entry(range_frame, textvariable=self.batch_limit, width=10).pack(side="left")
        ttk.Checkbutton(range_frame, text="Run all pending rows", variable=self.run_all_pending).pack(side="left", padx=(18, 0))
        self.run_range_text = tk.StringVar(value="Current run: all pending rows")
        ttk.Label(self.batch_tab, textvariable=self.run_range_text).grid(row=4, column=2, sticky="w", padx=(8, 0), pady=8)
        ttk.Label(self.batch_tab, text="Browser").grid(row=5, column=0, sticky="w", pady=8)
        ttk.Combobox(self.batch_tab, textvariable=self.browser, values=("chrome", "edge"), state="readonly", width=12).grid(row=5, column=1, sticky="w", pady=8)
        speed_frame = ttk.Frame(self.batch_tab)
        speed_frame.grid(row=5, column=2, sticky="w", padx=(8, 0), pady=8)
        ttk.Label(speed_frame, text="Speed mode").pack(side="left")
        ttk.Combobox(speed_frame, textvariable=self.speed_mode, values=("Fast Safe", "Conservative"), state="readonly", width=14).pack(side="left", padx=(8, 0))
        options_frame = ttk.Frame(self.batch_tab)
        options_frame.grid(row=6, column=0, columnspan=4, sticky="ew", pady=4)
        
        ttk.Checkbutton(options_frame, text="Dry run only (safe test, no live insert)", variable=self.dry_run).grid(row=0, column=0, sticky="w", pady=4, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="Stop if one center fails", variable=self.stop_on_error).grid(row=0, column=1, sticky="w", pady=4)
        
        ttk.Checkbutton(options_frame, text="Autofill Login ID and password; I will enter CAPTCHA only", variable=self.manual_login, onvalue=False, offvalue=True).grid(row=1, column=0, sticky="w", pady=4, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="Remember password securely for this Windows user", variable=self.remember_credentials).grid(row=1, column=1, sticky="w", pady=4)
        
        ttk.Checkbutton(options_frame, text="Do not store browser session/profile", variable=self.no_store_session, state="disabled").grid(row=2, column=0, sticky="w", pady=4, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="Live mode fills the RBSK browser form", variable=self.browser_form_fill).grid(row=2, column=1, sticky="w", pady=4)
        
        ttk.Checkbutton(options_frame, text="Auto-select school when SCHOOL NAME changes", variable=self.auto_select_institute).grid(row=3, column=0, sticky="w", pady=4, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="Unattended after entry starts (no popup waits during rows)", variable=self.unattended_after_start).grid(row=3, column=1, sticky="w", pady=4)
        
        ttk.Checkbutton(options_frame, text="Allow manual popup fallback if unattended is off", variable=self.manual_form_open).grid(row=4, column=0, sticky="w", pady=4, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="RBSK certificate fix (keep checked if SSL error appears)", variable=self.allow_insecure_ssl).grid(row=4, column=1, sticky="w", pady=4)

        self.mobile_frame = ttk.LabelFrame(self.batch_tab, text=" Mobile Number Option for Confirmation ", padding=6)
        self.mobile_frame.grid(row=7, column=0, columnspan=4, sticky="ew", pady=4)
        ttk.Radiobutton(self.mobile_frame, text="Direct Portal / Enter Mobile per School (Default)", variable=self.mobile_source, value="portal_school").pack(side="left", padx=(10, 14))
        ttk.Radiobutton(self.mobile_frame, text="Add Mobile using Excel File", variable=self.mobile_source, value="excel").pack(side="left", padx=14)
        ttk.Label(self.mobile_frame, text="School Mobile No:").pack(side="left", padx=(18, 4))
        ttk.Entry(self.mobile_frame, textvariable=self.custom_school_mobile, width=15).pack(side="left")
        ttk.Label(self.mobile_frame, text="(blank = asks mobile per school)").pack(side="left", padx=(6, 0))

        screening_frame = ttk.Frame(self.batch_tab)
        screening_frame.grid(row=8, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Checkbutton(screening_frame, text="Screening Health mode only (no Add Student)", variable=self.screening_health_entry, state="disabled").pack(side="left")
        ttk.Label(screening_frame, text="Institution Type").pack(side="left", padx=(18, 6))
        ttk.Combobox(
            screening_frame,
            textvariable=self.screening_institution_type,
            values=("School", "AWC", "AWS"),
            state="readonly",
            width=10,
        ).pack(side="left")
        ttk.Label(screening_frame, text="Select Month").pack(side="left", padx=(18, 6))
        ttk.Combobox(
            screening_frame,
            textvariable=self.screening_month,
            values=screening_month_options(),
            state="readonly",
            width=18,
        ).pack(side="left")
        ttk.Label(screening_frame, text="Select here for automatic Search; blank means manual").pack(side="left", padx=(8, 0))

        def _on_institution_type_change(*_):
            if Students.is_awc_institution(self.screening_institution_type.get()):
                self.mobile_frame.grid_remove()
            else:
                self.mobile_frame.grid(row=7, column=0, columnspan=4, sticky="ew", pady=4)
        self.screening_institution_type.trace_add("write", _on_institution_type_change)

        license_frame = ttk.LabelFrame(self.batch_tab, text=" Monthly Licence ", padding=6)
        license_frame.grid(row=9, column=0, columnspan=4, sticky="ew", pady=6)
        ttk.Label(license_frame, text=f"Machine ID: {app_machine_id()}").grid(row=0, column=0, sticky="w", padx=(10, 12), pady=4)
        ttk.Button(license_frame, text="Copy Machine ID", command=self.copy_machine_id).grid(row=0, column=1, sticky="w", padx=(0, 18), pady=4)
        ttk.Label(license_frame, text="Enter Access Key Here").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(license_frame, textvariable=self.license_key, width=44, show="*").grid(row=0, column=3, sticky="ew", padx=(0, 8), pady=4)
        ttk.Button(license_frame, text="Activate Key", command=self.activate_license_from_ui).grid(row=0, column=4, sticky="w", padx=(0, 12), pady=4)
        ttk.Label(license_frame, textvariable=self.license_status, wraplength=360).grid(row=1, column=0, columnspan=5, sticky="w", padx=(10, 0), pady=(2, 0))
        license_frame.columnconfigure(3, weight=1)

        buttons = ttk.Frame(self.batch_tab)
        buttons.grid(row=10, column=0, columnspan=4, sticky="w", pady=(12, 0))
        ttk.Button(buttons, text="Start Direct Portal Screening (No Excel)", style="Primary.TButton", command=self.start_direct_portal_screening).pack(side="left", padx=(0, 10))
        ttk.Button(buttons, text="Start Screening Health (Excel File)", command=self.start_screening_health).pack(side="left", padx=(0, 10))
        ttk.Button(buttons, text="Update Other PC Folder", command=self.update_other_pc_folder).pack(side="left")
        self.batch_control.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.range_start.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.batch_limit.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.run_all_pending.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.screening_month.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.only_batch.trace_add("write", lambda *_: self.schedule_dashboard_refresh())
        self.default_login_id.trace_add("write", lambda *_: (self.remember_login_id_only(self.default_login_id.get()), self.refresh_saved_login_status()))
        self.login_password.trace_add("write", lambda *_: self.refresh_saved_login_status())
        self.refresh_saved_login_status()
        self.refresh_license_status()

        self.batch_tab.columnconfigure(1, weight=1)

    def _build_dashboard(self, parent) -> None:
        dashboard = ttk.Frame(parent, style="Panel.TFrame", padding=12)
        dashboard.pack(fill="both", expand=True)

        header = ttk.Frame(dashboard, style="Panel.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="Institute Completion Dashboard", style="Title.TLabel").pack(side="left")
        self.dashboard_range_text = tk.StringVar(value="Current run: all pending rows")
        ttk.Label(header, textvariable=self.dashboard_range_text, style="Panel.TLabel").pack(side="left", padx=(18, 0))
        ttk.Button(header, text="Refresh", command=self.refresh_dashboard).pack(side="right")

        cards = tk.Frame(dashboard, bg="#ffffff")
        cards.pack(fill="x", pady=(10, 8))
        self.card_vars = {
            "total": tk.StringVar(value="0"),
            "run": tk.StringVar(value="0"),
            "completed": tk.StringVar(value="0"),
            "pending": tk.StringVar(value="0"),
            "failed": tk.StringVar(value="0"),
            "current": tk.StringVar(value="-"),
            "mapping": tk.StringVar(value="Not checked"),
            "speed": tk.StringVar(value="-"),
        }
        card_specs = [
            ("Total Rows", "total", "#eff6ff", "#1d4ed8"),
            ("Selected Run", "run", "#f8fafc", "#334155"),
            ("Completed", "completed", "#ecfdf5", "#047857"),
            ("Pending", "pending", "#fff7ed", "#c2410c"),
            ("Failed This Run", "failed", "#fef2f2", "#b91c1c"),
            ("Current School", "current", "#f0fdfa", "#0f766e"),
            ("Screening Mode", "mapping", "#f5f3ff", "#6d28d9"),
            ("Speed / ETA", "speed", "#f0f9ff", "#0369a1"),
        ]
        for column, (title, key, bg, fg) in enumerate(card_specs):
            card = tk.Frame(cards, bg=bg, highlightbackground="#d9e2ec", highlightthickness=1)
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            cards.columnconfigure(column, weight=1)
            tk.Label(card, text=title, bg=bg, fg="#334155", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
            tk.Label(card, textvariable=self.card_vars[key], bg=bg, fg=fg, font=("Segoe UI", 14, "bold"), wraplength=190, justify="left").pack(anchor="w", padx=10, pady=(2, 8))

        bars = ttk.Frame(dashboard, style="Panel.TFrame")
        bars.pack(fill="x", pady=(0, 8))
        ttk.Label(bars, text="Overall progress", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        self.overall_progress = ttk.Progressbar(bars, mode="determinate", maximum=100)
        self.overall_progress.grid(row=0, column=1, sticky="ew", padx=(8, 14))
        self.overall_progress_text = tk.StringVar(value="0%")
        ttk.Label(bars, textvariable=self.overall_progress_text, style="Panel.TLabel", width=8).grid(row=0, column=2, sticky="w")
        ttk.Label(bars, text="Current School", style="Panel.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.institute_progress = ttk.Progressbar(bars, mode="determinate", maximum=100)
        self.institute_progress.grid(row=1, column=1, sticky="ew", padx=(8, 14), pady=(6, 0))
        self.institute_progress_text = tk.StringVar(value="0%")
        ttk.Label(bars, textvariable=self.institute_progress_text, style="Panel.TLabel", width=14).grid(row=1, column=2, sticky="w", pady=(6, 0))
        bars.columnconfigure(1, weight=1)

        table_frame = ttk.Frame(dashboard, style="Panel.TFrame")
        table_frame.pack(fill="both", expand=True)
        columns = ("awc", "total", "completed", "pending", "failed", "status")
        self.institute_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=7, style="Dashboard.Treeview")
        self.institute_tree.tag_configure("current", background="#dff7f5")
        headings = {
            "awc": "School / Institute Name",
            "total": "Total",
            "completed": "Completed",
            "pending": "Pending",
            "failed": "Failed",
            "status": "Status",
        }
        widths = {"awc": 360, "total": 70, "completed": 90, "pending": 80, "failed": 70, "status": 140}
        for col in columns:
            self.institute_tree.heading(col, text=headings[col])
            self.institute_tree.column(col, width=widths[col], anchor="w" if col in {"awc", "status"} else "center")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.institute_tree.yview)
        self.institute_tree.configure(yscrollcommand=scrollbar.set)
        self.institute_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _build_logs_tab(self) -> None:
        panel = ttk.Frame(self.logs_tab, style="Panel.TFrame", padding=12)
        panel.pack(fill="both", expand=True)

        header = ttk.Frame(panel, style="Panel.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="Failed Entry Logs", style="Title.TLabel").pack(side="left")
        ttk.Button(header, text="Refresh Logs", command=self.refresh_logs).pack(side="right")
        ttk.Button(header, text="Open Logs Folder", command=lambda: self.open_path(APP_DIR / "logs")).pack(side="right", padx=(0, 8))
        ttk.Button(header, text="Start New File / Clear Logs", command=self.start_new_file_clear_logs).pack(side="right", padx=(0, 8))
        ttk.Button(header, text="Copy Failed Sr Nos", command=self.copy_pending_failed_sr_nos).pack(side="right", padx=(0, 8))
        ttk.Button(header, text="Export Pending Failed", command=lambda: self.export_logs(pending_only=True)).pack(side="right", padx=(0, 8))
        ttk.Button(header, text="Export All Logs", command=lambda: self.export_logs(pending_only=False)).pack(side="right", padx=(0, 8))

        cards = tk.Frame(panel, bg="#ffffff")
        cards.pack(fill="x", pady=(10, 8))
        self.log_card_vars = {
            "failed": tk.StringVar(value="0"),
            "fixed": tk.StringVar(value="0"),
            "pending": tk.StringVar(value="0"),
            "last": tk.StringVar(value="-"),
        }
        specs = [
            ("Failed Log Rows", "failed", "#fef2f2", "#b91c1c"),
            ("Fixed / Completed", "fixed", "#ecfdf5", "#047857"),
            ("Pending Failed", "pending", "#fff7ed", "#c2410c"),
            ("Last Run", "last", "#eff6ff", "#1d4ed8"),
        ]
        for column, (title, key, bg, fg) in enumerate(specs):
            card = tk.Frame(cards, bg=bg, highlightbackground="#d9e2ec", highlightthickness=1)
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            cards.columnconfigure(column, weight=1)
            tk.Label(card, text=title, bg=bg, fg="#334155", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
            tk.Label(card, textvariable=self.log_card_vars[key], bg=bg, fg=fg, font=("Segoe UI", 13, "bold"), wraplength=220, justify="left").pack(anchor="w", padx=10, pady=(2, 8))

        table_frame = ttk.Frame(panel, style="Panel.TFrame")
        table_frame.pack(fill="both", expand=True)
        columns = ("sr", "name", "awc", "status", "reason", "details")
        self.logs_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14, style="Dashboard.Treeview")
        self.logs_tree.tag_configure("pending", background="#fff7ed")
        self.logs_tree.tag_configure("fixed", background="#ecfdf5")
        headings = {
            "sr": "Sr No",
            "name": "Child Name",
            "awc": "AWC NAME",
            "status": "Status",
            "reason": "Easy Reason",
            "details": "Screenshot / Details",
        }
        widths = {"sr": 70, "name": 220, "awc": 230, "status": 130, "reason": 260, "details": 220}
        for col in columns:
            self.logs_tree.heading(col, text=headings[col])
            self.logs_tree.column(col, width=widths[col], anchor="w" if col != "sr" else "center")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.logs_tree.yview)
        self.logs_tree.configure(yscrollcommand=scrollbar.set)
        self.logs_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.logs_tree.bind("<Double-1>", self.show_selected_log_details)

    def simple_error_reason(self, error: str) -> str:
        text = Students.clean(error).lower()
        if "duplicate record" in text or "already exists in rbsk" in text or "cannot create a duplicate" in text:
            return "Already exists in RBSK (duplicate)"
        if "duplicate window" in text:
            return "Skipped after duplicate window"
        if "missing mobile" in text:
            return "Missing mobile number"
        if "not mapped" in text or "unmapped" in text:
            return "AWC name does not match RBSK list"
        if "dob/gender" in text or "birth date" in text or "gender expected" in text:
            return "Portal did not accept DOB/Gender"
        if "validation" in text or "required" in text or "must" in text:
            return "RBSK form validation failed"
        if "save button" in text or "save/submit" in text or "could not click" in text:
            return "Save button could not be clicked"
        if "institute" in text or "awc" in text or "select institution" in text:
            if "mapping" in text or "alias" in text or "not found in rbsk list" in text:
                return "AWC name does not match RBSK list"
            return "Institute/AWC was not ready"
        if "form was not detected" in text or "form did not open" in text or "manual handoff" in text:
            return "Final Add Child form was not open"
        if "portal is still loading" in text or "loading" in text:
            return "RBSK page was still loading"
        if "devtoolsactiveport" in text or "session not created" in text or "chrome failed to start" in text:
            return "Browser session/profile could not start"
        if "invalid session id" in text or "browser was closed" in text:
            return "Browser was closed during entry"
        return "Check details before retry"

    def diagnostic_portal_reason(self, diagnostic: Path) -> str:
        if not diagnostic.exists():
            return ""
        try:
            text = diagnostic.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return ""
        plain = re.sub(r"<[^>]+>", " ", text)
        plain = Students.clean(plain)
        if "Duplicate Record Found" in plain or "cannot create a duplicate" in plain:
            child_id = ""
            name = ""
            dob = ""
            id_match = re.search(r"STUDENT/CHILD ID\s+([A-Za-z0-9-]+)", plain, flags=re.IGNORECASE)
            name_match = re.search(r"NAME\s+(.+?)\s+DATE OF BIRTH", plain, flags=re.IGNORECASE)
            dob_match = re.search(r"DATE OF BIRTH\s+([0-9-]+)", plain, flags=re.IGNORECASE)
            if id_match:
                child_id = id_match.group(1)
            if name_match:
                name = Students.clean(name_match.group(1))
            if dob_match:
                dob = dob_match.group(1)
            reason = "Duplicate record already exists in RBSK"
            if child_id:
                reason += f" (Student/Child ID {child_id})"
            parts = []
            if name:
                parts.append(f"Name: {name}")
            if dob:
                parts.append(f"DOB: {dob}")
            if parts:
                reason += ": " + ", ".join(parts)
            return reason
        if "Date must be between 6 weeks and 6 years old" in plain:
            return "RBSK validation after Save: Date must be between 6 weeks and 6 years old."
        return ""

    def refresh_logs(self) -> None:
        try:
            self.log_rows = self.load_log_rows()
            self.update_logs_view()
        except Exception as exc:
            if hasattr(self, "logs_tree"):
                self.logs_tree.delete(*self.logs_tree.get_children())
                self.logs_tree.insert("", "end", values=("-", "-", "-", "Log error", str(exc), "-"))

    def load_log_rows(self) -> list[dict]:
        if not self.dashboard_rows:
            try:
                rows, completed = self.load_dashboard_rows()
                self.dashboard_rows = rows
                self.dashboard_completed = completed
            except Exception:
                pass
        row_map = {row.get("_dashboard_key"): row for row in self.dashboard_rows}
        logs_root = APP_DIR / "logs"
        batch_names = sorted({Students.clean(row.get("_dashboard_batch")) for row in self.dashboard_rows if Students.clean(row.get("_dashboard_batch"))}) or ["Direct Student File"]
        result: list[dict] = []
        for batch_name in batch_names:
            batch_dir = logs_root / Students.slugify(batch_name)
            failed_path = batch_dir / "failed.csv"
            completed = {self.dashboard_key(batch_name, sr_no) for sr_no in Students.load_progress(batch_dir / "progress.json")}
            if failed_path.exists():
                with failed_path.open("r", encoding="utf-8-sig", newline="") as handle:
                    for item in csv.DictReader(handle):
                        sr_no = Students.clean(item.get("Sr No"))
                        key = self.dashboard_key(batch_name, sr_no)
                        source_row = row_map.get(key, {})
                        name = Students.clean(item.get("Student Full Name")) or Students.clean(source_row.get("Student Full Name"))
                        awc = Students.clean(source_row.get("Institute Name")) or "(not found in Excel)"
                        error = Students.clean(item.get("error"))
                        screenshot = batch_dir / f"failure_sr_{sr_no}.png"
                        diagnostic = batch_dir / f"failure_sr_{sr_no}_buttons.html"
                        diagnostic_reason = self.diagnostic_portal_reason(diagnostic)
                        if diagnostic_reason and ("form stayed open" in error.lower() or "check rbsk validation" in error.lower() or not error):
                            error = diagnostic_reason
                        details = []
                        if screenshot.exists():
                            details.append("screenshot")
                        if diagnostic.exists():
                            details.append("details")
                        result.append(
                            {
                                "batch": batch_name,
                                "key": key,
                                "sr_no": sr_no,
                                "name": name,
                                "awc": awc,
                                "status": "Fixed / completed" if key in completed else "Pending failed",
                                "reason": self.simple_error_reason(error),
                                "error": error,
                                "screenshot": str(screenshot) if screenshot.exists() else "",
                                "diagnostic": str(diagnostic) if diagnostic.exists() else "",
                                "details": ", ".join(details) if details else "no file",
                            }
                        )
            skipped_path = batch_dir / "skipped.csv"
            if skipped_path.exists():
                with skipped_path.open("r", encoding="utf-8-sig", newline="") as handle:
                    for item in csv.DictReader(handle):
                        sr_no = Students.clean(item.get("Sr No"))
                        key = self.dashboard_key(batch_name, sr_no)
                        source_row = row_map.get(key, {})
                        error = Students.clean(item.get("reason")) or Students.clean(item.get("status"))
                        result.append(
                            {
                                "batch": batch_name,
                                "key": key,
                                "sr_no": sr_no,
                                "name": Students.clean(item.get("Student Full Name")) or Students.clean(source_row.get("Student Full Name")),
                                "awc": Students.clean(item.get("AWC NAME")) or Students.clean(source_row.get("Institute Name")) or "(not found in Excel)",
                                "status": "Fixed / completed" if key in completed else "Skipped this run",
                                "reason": self.simple_error_reason(error),
                                "error": error,
                                "screenshot": "",
                                "diagnostic": "",
                                "details": Students.clean(item.get("status")) or "skipped",
                                "suggested_action": Students.clean(item.get("suggested_action")),
                            }
                        )
        return result

    def update_logs_view(self) -> None:
        if not hasattr(self, "logs_tree"):
            return
        self.logs_tree.delete(*self.logs_tree.get_children())
        unique_pending = {row["key"] for row in self.log_rows if row["status"] in {"Pending failed", "Skipped this run"}}
        unique_fixed = {row["key"] for row in self.log_rows if row["status"] == "Fixed / completed"}
        if hasattr(self, "log_card_vars"):
            self.log_card_vars["failed"].set(str(len(self.log_rows)))
            self.log_card_vars["fixed"].set(str(len(unique_fixed)))
            self.log_card_vars["pending"].set(str(len(unique_pending)))
            self.log_card_vars["last"].set(self.latest_batch_summary_status())
        for index, row in enumerate(self.log_rows):
            tag = "fixed" if row["status"] == "Fixed / completed" else "pending"
            self.logs_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(row["sr_no"], row["name"], row["awc"], row["status"], row["reason"], row["details"]),
                tags=(tag,),
            )

    def latest_batch_summary_status(self) -> str:
        path = APP_DIR / "logs" / "batch_summary.csv"
        if not path.exists():
            return "-"
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
            if not rows:
                return "-"
            last = rows[-1]
            status = Students.clean(last.get("status")) or "-"
            saved = Students.clean(last.get("saved")) or "0"
            failed = Students.clean(last.get("failed")) or "0"
            return f"{status} | saved {saved}, failed {failed}"
        except Exception:
            return "-"

    def pending_failed_sr_nos(self) -> list[str]:
        seen: set[str] = set()
        values: list[str] = []
        for row in self.log_rows:
            if row["status"] not in {"Pending failed", "Skipped this run"}:
                continue
            sr_no = Students.clean(row.get("sr_no"))
            if sr_no and sr_no not in seen:
                seen.add(sr_no)
                values.append(sr_no)
        return values

    def suggested_action_for_log(self, row: dict) -> str:
        explicit = Students.clean(row.get("suggested_action"))
        if explicit:
            return explicit
        status = row.get("status", "")
        if status == "Fixed / completed":
            return "No action needed; this Sr No is already completed."
        reason = row.get("reason", "")
        if "duplicate window" in reason:
            return "Review and rerun these Sr Nos later; they were not submitted or marked completed."
        if "Missing mobile" in reason:
            return "Add a mobile number or place a valid same-AWC mobile before this row, then rerun."
        if "duplicate" in reason or "Already exists in RBSK" in reason:
            return "Verify this child already exists in RBSK. If correct, use Mark Completed Sr Nos so it is skipped next time."
        if "DOB/Gender" in reason:
            return "Open this row again and check DOB/Gender on the final form before Save."
        if "validation" in reason:
            return "Check required fields in Excel, then retry this Sr No."
        if "Save button" in reason:
            return "Retry this Sr No after scrolling to the Save button."
        if "Institute/AWC" in reason:
            return "Select the correct institute/AWC, click Add Child if visible, or Search only if Add Child is not visible, then retry."
        if "Final Add Child form" in reason:
            return "Open Add Child and wait for the final form, then retry."
        if "loading" in reason:
            return "Refresh Screening, wait for loading to finish, then retry."
        if "Browser session" in reason:
            return "Click Reset Browser Session for this Login ID, then start again from the suggested Sr No."
        if "Browser was closed" in reason:
            return "Restart the helper from the suggested Sr No. Completed rows will be skipped."
        return "Review the full error, correct the row or portal state, then retry."

    def export_logs(self, pending_only: bool) -> None:
        self.refresh_logs()
        rows = [row for row in self.log_rows if not pending_only or row["status"] in {"Pending failed", "Skipped this run"}]
        if not rows:
            messagebox.showinfo("No rows", "No log rows to export.")
            return
        default_name = "pending_failed_entries.csv" if pending_only else "all_failed_entry_logs.csv"
        path = filedialog.asksaveasfilename(
            initialdir=APP_DIR / "logs",
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("CSV for Excel", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        headers = [
            "Sr No",
            "Child Name",
            "AWC NAME",
            "Status",
            "Easy Reason",
            "Full Error",
            "Screenshot",
            "Details",
            "Suggested Action",
        ]
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers)
            writer.writeheader()
            for row in rows:
                writer.writerow(
                    {
                        "Sr No": row.get("sr_no", ""),
                        "Child Name": row.get("name", ""),
                        "AWC NAME": row.get("awc", ""),
                        "Status": row.get("status", ""),
                        "Easy Reason": row.get("reason", ""),
                        "Full Error": row.get("error", ""),
                        "Screenshot": row.get("screenshot", ""),
                        "Details": row.get("diagnostic", ""),
                        "Suggested Action": self.suggested_action_for_log(row),
                    }
                )
        self.show_guided_dialog("Logs Exported", f"Exported {len(rows)} row(s) to:\n{path}", wait=False)

    def selected_batch_log_infos(self) -> list[dict]:
        args = self._batch_args()
        control_path = Path(args.batch_control)
        batch_rows = self.load_batch_or_direct_student_file(control_path, args)
        Students.validate_batch_control(batch_rows)
        selected_batches = Students.select_rows_by_batch(batch_rows, args.only_batch)
        logs_root = APP_DIR / args.logs_dir
        infos: list[dict] = []
        for batch_index, batch_row in enumerate(selected_batches, start=1):
            batch_name = Students.clean(batch_row.get("Batch Name")) or f"Batch {batch_index}"
            batch_log_dir = logs_root / Students.slugify(batch_name)
            excel_path = Students.resolve_batch_path(batch_row.get("Excel File", ""), control_path)
            sheet = Students.clean(batch_row.get("Sheet")) or args.sheet
            rows = Students.normalize_student_rows(Students.load_rows(excel_path, sheet))
            if getattr(args, "screening_health_entry", False):
                Students.validate_screening_health_rows(rows)
            else:
                Students.validate_columns(rows)
            infos.append(
                {
                    "batch_name": batch_name,
                    "batch_dir": batch_log_dir,
                    "excel_path": excel_path,
                    "sheet": sheet,
                    "rows": rows,
                    "metadata": self.build_run_metadata(excel_path, sheet, rows),
                }
            )
        return infos

    def direct_batch_name_for_file(self, path: Path) -> str:
        stem = Students.clean(path.stem) or "Student_File"
        return f"Direct - {stem}"

    def build_run_metadata(self, excel_path: Path, sheet: str, rows: list[dict]) -> dict:
        try:
            stat = excel_path.stat()
            size = stat.st_size
            modified = stat.st_mtime
        except Exception:
            size = 0
            modified = 0
        sr_values = [Students.clean(row.get("Sr No")) for row in rows if Students.clean(row.get("Sr No"))]
        return {
            "excel_path": str(excel_path.resolve()) if excel_path.exists() else str(excel_path),
            "sheet": sheet,
            "file_size": size,
            "modified_time": modified,
            "row_count": len(rows),
            "first_sr_no": sr_values[0] if sr_values else "",
            "last_sr_no": sr_values[-1] if sr_values else "",
        }

    def metadata_path_for_batch_dir(self, batch_dir: Path) -> Path:
        return batch_dir / "run_metadata.json"

    def read_run_metadata(self, batch_dir: Path) -> dict:
        path = self.metadata_path_for_batch_dir(batch_dir)
        if not path.exists():
            return {}
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def write_run_metadata(self, batch_dir: Path, metadata: dict) -> None:
        batch_dir.mkdir(parents=True, exist_ok=True)
        self.metadata_path_for_batch_dir(batch_dir).write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    def metadata_differs(self, old: dict, new: dict) -> bool:
        if not old:
            return False
        exact_keys = ["excel_path", "sheet", "file_size", "row_count", "first_sr_no", "last_sr_no"]
        if any(str(old.get(key, "")) != str(new.get(key, "")) for key in exact_keys):
            return True
        try:
            return abs(float(old.get("modified_time", 0) or 0) - float(new.get("modified_time", 0) or 0)) > 1.0
        except Exception:
            return str(old.get("modified_time", "")) != str(new.get("modified_time", ""))

    def log_file_counts(self, batch_dir: Path) -> dict[str, int]:
        completed = len(Students.load_progress(batch_dir / "progress.json"))

        def csv_count(path: Path) -> int:
            if not path.exists():
                return 0
            try:
                with path.open("r", encoding="utf-8-sig", newline="") as handle:
                    return sum(1 for _row in csv.DictReader(handle))
            except Exception:
                return 0

        return {
            "completed": completed,
            "success": csv_count(batch_dir / "success.csv"),
            "failed": csv_count(batch_dir / "failed.csv"),
            "skipped": csv_count(batch_dir / "skipped.csv"),
        }

    def archive_and_reset_batch_logs(self, info: dict) -> Path | None:
        batch_dir: Path = info["batch_dir"]
        if not batch_dir.exists():
            batch_dir.mkdir(parents=True, exist_ok=True)
            self.write_run_metadata(batch_dir, info["metadata"])
            Students.save_progress(batch_dir / "progress.json", set())
            return None
        archive_root = APP_DIR / "logs" / "_archived"
        archive_root.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        archive_dir = archive_root / f"{Students.slugify(info['batch_name'])}_{timestamp}"
        archive_dir.mkdir(parents=True, exist_ok=False)
        moved = False
        patterns = [
            "progress.json",
            "success.csv",
            "failed.csv",
            "skipped.csv",
            "run_metadata.json",
            "institute_mapping_missing.csv",
            "rbsk_institute_options.csv",
            "failure_sr_*.png",
            "failure_sr_*_buttons.html",
        ]
        for pattern in patterns:
            for path in batch_dir.glob(pattern):
                if not path.exists() or path.is_dir():
                    continue
                target = archive_dir / path.name
                shutil.move(str(path), str(target))
                moved = True
        self.write_run_metadata(batch_dir, info["metadata"])
        Students.save_progress(batch_dir / "progress.json", set())
        return archive_dir if moved else None

    def start_new_file_clear_logs(self) -> None:
        try:
            infos = self.selected_batch_log_infos()
        except Exception as exc:
            messagebox.showerror("Cannot read selected file", str(exc))
            return
        if not infos:
            messagebox.showinfo("No batch selected", "No selected file or center was found.")
            return
        lines = ["Archive old logs/progress and start fresh for the selected file?", ""]
        for info in infos:
            counts = self.log_file_counts(info["batch_dir"])
            lines.extend(
                [
                    f"Batch: {info['batch_name']}",
                    f"Excel: {info['metadata']['excel_path']}",
                    f"Log folder: {info['batch_dir']}",
                    f"Rows in selected Excel: {info['metadata']['row_count']}",
                    f"Completed progress: {counts['completed']}",
                    f"Success log rows: {counts['success']}",
                    f"Failed log rows: {counts['failed']}",
                    f"Skipped log rows: {counts['skipped']}",
                    "",
                ]
            )
        lines.append("Old files will be moved to logs/_archived, not permanently deleted.")
        if not messagebox.askyesno("Start New File / Clear Logs", "\n".join(lines)):
            return
        archives: list[str] = []
        for info in infos:
            archive_dir = self.archive_and_reset_batch_logs(info)
            if archive_dir:
                archives.append(str(archive_dir))
        self.refresh_dashboard(reset_run_failures=True)
        self.refresh_logs()
        message = "Fresh progress is ready for the selected file.\n\n"
        if archives:
            message += "Archived old logs here:\n" + "\n".join(archives[:8])
            if len(archives) > 8:
                message += f"\n...and {len(archives) - 8} more"
        else:
            message += "No old log files needed archiving."
        message += (
            "\n\nImportant: this only clears the helper's local logs/progress. "
            "Records already saved inside RBSK will still be blocked by RBSK as duplicate records. "
            "If a duplicate is correct, use Mark Completed Sr Nos after verifying it in RBSK."
        )
        self.show_guided_dialog("New File Started", message, wait=False)

    def warn_if_selected_file_changed(self) -> bool:
        try:
            infos = self.selected_batch_log_infos()
        except Exception as exc:
            messagebox.showerror("Cannot check selected file", str(exc))
            return False
        changed_infos = []
        for info in infos:
            old_metadata = self.read_run_metadata(info["batch_dir"])
            counts = self.log_file_counts(info["batch_dir"])
            if counts["completed"] and self.metadata_differs(old_metadata, info["metadata"]):
                changed_infos.append((info, old_metadata, counts))
        if not changed_infos:
            for info in infos:
                if not self.read_run_metadata(info["batch_dir"]):
                    self.write_run_metadata(info["batch_dir"], info["metadata"])
            return True
        lines = [
            "This looks like a new Excel file, but old completed Sr Nos exist.",
            "",
            "If this is a new file with repeated Sr Nos, archive old logs and start fresh.",
            "",
        ]
        for info, _old, counts in changed_infos:
            lines.extend(
                [
                    f"Batch: {info['batch_name']}",
                    f"Excel: {info['metadata']['excel_path']}",
                    f"Completed old Sr Nos: {counts['completed']}",
                    "",
                ]
            )
        lines.append("Archive old logs and start fresh now?")
        archive_now = messagebox.askyesno("New file detected", "\n".join(lines))
        if archive_now:
            for info, _old, _counts in changed_infos:
                self.archive_and_reset_batch_logs(info)
            self.refresh_dashboard(reset_run_failures=True)
            self.refresh_logs()
        return True

    def copy_pending_failed_sr_nos(self) -> None:
        self.refresh_logs()
        values = self.pending_failed_sr_nos()
        self.clipboard_clear()
        self.clipboard_append(",".join(values))
        messagebox.showinfo("Copied", f"Copied {len(values)} pending failed Sr No(s).")

    def latest_missing_report_path(self) -> Path | None:
        logs_root = APP_DIR / "logs"
        if not logs_root.exists():
            return None
        reports = sorted(
            logs_root.glob("*/institute_mapping_missing.csv"),
            key=lambda path: path.stat().st_mtime if path.exists() else 0,
            reverse=True,
        )
        return reports[0] if reports else None

    def missing_institute_names_from_report(self, path: Path | None) -> list[str]:
        if not path or not path.exists():
            return []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                return [Students.clean(row.get("Excel AWC NAME")) for row in csv.DictReader(handle) if Students.clean(row.get("Excel AWC NAME"))]
        except Exception:
            return []

    def read_alias_rows(self) -> list[dict[str, str]]:
        path = APP_DIR / "institute_aliases.csv"
        if not path.exists():
            return []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                return list(csv.DictReader(handle))
        except Exception:
            return []

    def write_alias_rows(self, rows: list[dict[str, str]]) -> None:
        path = APP_DIR / "institute_aliases.csv"
        headers = ["Excel AWC NAME", "RBSK Institute Name"]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers)
            writer.writeheader()
            for row in rows:
                excel_name = Students.clean(row.get("Excel AWC NAME"))
                portal_name = Students.clean(row.get("RBSK Institute Name"))
                if excel_name and portal_name:
                    writer.writerow({"Excel AWC NAME": excel_name, "RBSK Institute Name": portal_name})

    def review_missing_institute_names(self) -> None:
        report_path = self.latest_missing_report_path()
        if not report_path or not report_path.exists():
            messagebox.showinfo("No report", "No institute_mapping_missing.csv report was found. Run Check Institute Names first.")
            return
        with report_path.open("r", encoding="utf-8-sig", newline="") as handle:
            missing_rows = list(csv.DictReader(handle))
        if not missing_rows:
            messagebox.showinfo("No missing names", "The latest missing institute report has no rows.")
            return

        aliases = self.read_alias_rows()
        alias_by_key = {Students.normalize_match_text(row.get("Excel AWC NAME")): dict(row) for row in aliases}
        updated = 0
        skipped = 0
        for row in missing_rows:
            excel_name = Students.clean(row.get("Excel AWC NAME"))
            if not excel_name:
                continue
            initial = Students.clean(row.get("Matched RBSK Institute Name"))
            closest = Students.clean(row.get("Closest RBSK Options"))
            options_file = Students.clean(row.get("All RBSK Options File"))
            prompt = (
                f"Excel AWC NAME:\n{excel_name}\n\n"
                "Type the exact RBSK Institute Name from the dropdown.\n"
            )
            if closest:
                prompt += f"\nClosest options from RBSK:\n{closest}\n"
            if options_file:
                prompt += f"\nFull option list file:\n{options_file}\n"
            portal_name = simpledialog.askstring("Review Missing Institute", prompt, initialvalue=initial)
            portal_name = Students.clean(portal_name)
            if not portal_name:
                skipped += 1
                continue
            alias_by_key[Students.normalize_match_text(excel_name)] = {
                "Excel AWC NAME": excel_name,
                "RBSK Institute Name": portal_name,
            }
            updated += 1

        self.write_alias_rows(list(alias_by_key.values()))
        self.show_guided_dialog(
            "Institute Aliases Updated",
            f"Updated {updated} institute alias row(s).\n"
            f"Skipped {skipped} row(s).\n\n"
            "Next: click Check Institute Names again. If all names match, start the selected row block.",
            wait=False,
        )

    def mark_completed_sr_nos(self) -> None:
        self.refresh_dashboard()
        text = simpledialog.askstring(
            "Mark Completed Sr Nos",
            "Enter Sr Nos already verified as saved in RBSK, for example:\n2004,2005\n\n"
            "Use this only when the row exists in RBSK but the script stopped before recording progress.",
        )
        if text is None:
            return
        try:
            wanted = Students.parse_sr_numbers(text)
        except Exception as exc:
            messagebox.showerror("Invalid Sr Nos", str(exc))
            return
        if not wanted:
            messagebox.showinfo("No Sr Nos", "No Sr Nos were entered.")
            return
        rows = self.dashboard_rows or self.load_dashboard_rows()[0]
        matched: dict[str, set[str]] = {}
        preview: list[str] = []
        for row in rows:
            sr_no = Students.clean(row.get("Sr No"))
            if sr_no not in wanted:
                continue
            batch_name = Students.clean(row.get("_dashboard_batch")) or "Direct Student File"
            matched.setdefault(batch_name, set()).add(sr_no)
            if len(preview) < 12:
                preview.append(f"{sr_no} - {Students.clean(row.get('Student Full Name'))} ({Students.clean(row.get('Institute Name'))})")
        missing = sorted(
            wanted - {sr for values in matched.values() for sr in values},
            key=lambda x: (0, int(x)) if x.isdigit() else (1, x),
        )
        if not matched:
            messagebox.showerror("Not found", "None of those Sr Nos were found in the selected Excel/dashboard.")
            return
        message = "Mark these Sr Nos as completed in progress.json?\n\n" + "\n".join(preview)
        if len(preview) < sum(len(values) for values in matched.values()):
            message += "\n..."
        if missing:
            message += "\n\nNot found in selected file: " + ", ".join(missing[:20])
        message += "\n\nOnly click Yes after you have checked the records exist in RBSK."
        if not messagebox.askyesno("Confirm completed rows", message):
            return
        logs_root = APP_DIR / "logs"
        total_marked = 0
        for batch_name, sr_values in matched.items():
            progress_path = logs_root / Students.slugify(batch_name) / "progress.json"
            completed = Students.load_progress(progress_path)
            before = len(completed)
            completed.update(sr_values)
            Students.save_progress(progress_path, completed)
            total_marked += len(completed) - before
        self.refresh_dashboard(reset_run_failures=False)
        self.refresh_logs()
        self.show_guided_dialog(
            "Completed Rows Marked",
            f"Marked {total_marked} new Sr No(s) as completed.\n\n"
            "They will be skipped in the next run. Existing completed Sr Nos were left unchanged.",
            wait=False,
        )

    def show_selected_log_details(self, _event=None) -> None:
        selected = self.logs_tree.selection()
        if not selected:
            return
        row = self.log_rows[int(selected[0])]
        message = (
            f"Sr No: {row['sr_no']}\n"
            f"Child: {row['name']}\n"
            f"AWC NAME: {row['awc']}\n"
            f"Status: {row['status']}\n"
            f"Easy reason: {row['reason']}\n\n"
            f"Full error:\n{row['error'] or '(no error text)'}\n\n"
            f"Screenshot:\n{row['screenshot'] or '(not saved)'}\n\n"
            f"Details:\n{row['diagnostic'] or '(not saved)'}"
        )
        self.show_guided_dialog("Failure Details", message, wait=False)

    def schedule_dashboard_refresh(self) -> None:
        if self.dashboard_refresh_job:
            try:
                self.after_cancel(self.dashboard_refresh_job)
            except Exception:
                pass
        self.dashboard_refresh_job = self.after(350, self.refresh_dashboard)

    def dashboard_key(self, batch_name: str, sr_no: str) -> str:
        return f"{Students.slugify(batch_name)}::{Students.clean(sr_no)}"

    def open_dashboard_tab(self) -> None:
        self.notebook.select(self.dashboard_tab)
        self.refresh_dashboard()

    def apply_row_preset(self, count: int) -> None:
        self.run_all_pending.set(False)
        self.range_start.set("")
        self.batch_limit.set(str(count))
        self.current_run_only_sr, self.current_run_label = self.compute_run_range()
        self.update_run_range_label()
        self.refresh_dashboard()

    def apply_all_pending_preset(self) -> None:
        self.run_all_pending.set(True)
        self.range_start.set("")
        self.current_run_only_sr, self.current_run_label = self.compute_run_range()
        self.update_run_range_label()
        self.refresh_dashboard()

    def start_row_preset(self, count: int) -> None:
        self.run_all_pending.set(False)
        self.apply_row_preset(count)
        self.start_batch()

    def build_run_snapshot(self) -> dict:
        rows = self.current_run_rows()
        groups, skipped = self.pending_awc_blocks_for_preflight(rows, self.dashboard_completed)
        pending = sum(len(group["rows"]) for group in groups)
        next_sr = None
        for row in rows:
            sr_no = Students.clean(row.get("Sr No"))
            if sr_no.isdigit() and not self.row_is_completed(row):
                next_sr = int(sr_no)
                break
        keys = [row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), row.get("Sr No", "")) for row in rows]
        sr_values = [Students.clean(row.get("Sr No")) for row in rows if Students.clean(row.get("Sr No"))]
        return {
            "only_sr": self.current_run_only_sr,
            "label": self.current_run_label,
            "rows": rows,
            "keys": keys,
            "sr_values": sr_values,
            "row_count": len(rows),
            "skipped": skipped,
            "pending": pending,
            "groups": groups,
            "next_sr": next_sr,
            "all_pending": bool(hasattr(self, "run_all_pending") and self.run_all_pending.get()),
            "total_rows": len(self.dashboard_rows),
        }

    def live_start_summary(self, snapshot: dict | None = None) -> str:
        snapshot = snapshot or self.current_run_snapshot or self.build_run_snapshot()
        groups = snapshot.get("groups", [])
        awc_preview = ", ".join(f"{group['name']} ({len(group['rows'])})" for group in groups[:4])
        if len(groups) > 4:
            awc_preview += f", +{len(groups) - 4} more"
        if not awc_preview:
            awc_preview = "No pending rows"
        return (
            f"Run block: {snapshot.get('label')}\n"
            f"Selected rows: {snapshot.get('row_count', 0)}\n"
            f"Already completed/skipped in this block: {snapshot.get('skipped', 0)}\n"
            f"Pending to enter now: {snapshot.get('pending', 0)}\n"
            f"Next pending Sr No: {snapshot.get('next_sr') if snapshot.get('next_sr') is not None else 'none'}\n"
            f"AWC blocks: {len(groups)} ({awc_preview})"
        )

    def pending_awc_blocks_for_preflight(self, rows: list[dict], completed: set[str]) -> tuple[list[dict], int]:
        groups: list[dict] = []
        skipped = 0
        for row in rows:
            sr_no = Students.clean(row.get("Sr No"))
            key = row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), sr_no)
            if (sr_no and sr_no in completed) or (key and key in completed):
                skipped += 1
                continue
            awc = Students.clean(row.get("Institute Name")) or "(no institute)"
            if not groups or groups[-1]["name"] != awc:
                groups.append({"name": awc, "rows": [], "start_sr": sr_no})
            groups[-1]["rows"].append(row)
        return groups, skipped

    def preflight_check(self) -> None:
        try:
            self.refresh_dashboard(reset_run_failures=True)
            self.current_run_only_sr, self.current_run_label = self.compute_run_range()
            self.update_run_range_label()
            args = self._batch_args()
            if not self.warn_if_selected_file_changed():
                return
            control_path = Path(args.batch_control)
            batch_rows = self.load_batch_or_direct_student_file(control_path, args)
            Students.validate_batch_control(batch_rows)
            selected_batches = Students.select_rows_by_batch(batch_rows, args.only_batch)
            if not selected_batches:
                names = ", ".join(Students.available_batch_names(batch_rows)) or "none"
                raise Students.ImportErrorWithHelp(f"No batch rows selected. Leave Only Batch Name blank or use one of: {names}")

            lines = [
                "Preflight passed. No RBSK records were saved.",
                "",
                f"Run block: {self.current_run_label}",
                f"Selected center/login batches: {len(selected_batches)}",
            ]
            output_lines = ["Preflight Check", "==============="]

            for batch_index, batch_row in enumerate(selected_batches, start=1):
                batch_name = Students.clean(batch_row.get("Batch Name")) or f"Batch {batch_index}"
                login_id = Students.clean(batch_row.get("Login ID")) or Students.clean(getattr(args, "default_login_id", ""))
                batch_log_dir = APP_DIR / args.logs_dir / Students.slugify(batch_name)
                batch_args = Students.batch_arg_namespace(args, batch_row, control_path, batch_log_dir)
                profile_text = "Not reused"
                if getattr(args, "use_browser_profile", False):
                    profile_text = str(self.browser_profile_dir_for_login(login_id))

                rows = Students.normalize_student_rows(Students.load_rows(Path(batch_args.excel), batch_args.sheet))
                if getattr(batch_args, "screening_health_entry", False):
                    Students.validate_screening_health_rows(rows)
                else:
                    Students.validate_columns(rows)
                selected_rows = Students.select_rows(rows, batch_args.only_sr, batch_args.limit)
                if not selected_rows:
                    raise Students.ImportErrorWithHelp(f"No student rows selected for {batch_name}.")

                completed = Students.load_progress(Path(batch_args.progress)) if batch_args.resume else set()
                groups, skipped = self.pending_awc_blocks_for_preflight(selected_rows, completed)
                pending_count = sum(len(group["rows"]) for group in groups)
                first_sr = Students.clean(selected_rows[0].get("Sr No"))
                last_sr = Students.clean(selected_rows[-1].get("Sr No"))
                password_status = self.saved_login_status_text(login_id)
                awc_preview = ", ".join(
                    f"{group['name']} ({len(group['rows'])}, from {group['start_sr']})" for group in groups[:5]
                )
                if len(groups) > 5:
                    awc_preview += f", +{len(groups) - 5} more"
                if not awc_preview:
                    awc_preview = "No pending rows in this block"

                block = [
                    "",
                    f"{batch_index}. {batch_name}",
                    f"Login ID: {login_id or '(blank)'}",
                    f"Password: {password_status}",
                    f"Browser profile: {profile_text}",
                    f"Excel: {batch_args.excel}",
                    f"Selected Sr No: {first_sr}-{last_sr}",
                    f"Selected rows: {len(selected_rows)}",
                    f"Already completed in this block: {skipped}",
                    f"Pending rows to enter: {pending_count}",
                    f"AWC change blocks: {len(groups)}",
                    f"School order: {awc_preview}",
                ]
                lines.extend(block)
                output_lines.extend(block)

            lines.extend(
                [
                    "",
                    "Next:",
                    "1. If the row block and school order look correct, click Start Screening Health.",
                    "2. Enter CAPTCHA only when the browser opens.",
                    "3. On Screening, confirm State/District/Block/MHT, Financial Year, and Month.",
                ]
            )
            self.clear_output()
            self.append_output("\n".join(output_lines) + "\n")
            self.show_guided_dialog("Preflight Check", "\n".join(lines), wait=False)
        except Exception as exc:
            self.show_guided_dialog("Preflight Failed", str(exc), wait=False)

    def compute_run_range(self) -> tuple[str | None, str]:
        if hasattr(self, "run_all_pending") and self.run_all_pending.get():
            return None, "Current run: all pending rows"
        count = self._int_or_none(self.batch_limit.get())
        start_text = Students.clean(self.range_start.get())
        if count is not None and count <= 0:
            raise Students.ImportErrorWithHelp("Number of rows must be blank or greater than 0.")
        if start_text and not start_text.isdigit():
            raise Students.ImportErrorWithHelp("Start Sr No must be a number, for example 101.")
        start_sr = int(start_text) if start_text else self.next_pending_sr()
        if start_text and start_sr <= 0:
            raise Students.ImportErrorWithHelp("Start Sr No must be greater than 0.")
        if count is None:
            if start_text:
                return str(start_sr), f"Current run: Sr No {start_sr}"
            return None, "Current run: all rows"
        end_sr = start_sr + count - 1
        return f"{start_sr}-{end_sr}", f"Current run: Sr No {start_sr}-{end_sr} ({count} row(s))"

    def next_pending_sr(self) -> int:
        rows = self.dashboard_rows
        completed = self.dashboard_completed
        if not rows:
            try:
                rows, completed = self.load_dashboard_rows()
            except Exception:
                rows, completed = [], set()
        first_numeric = None
        for row in rows:
            sr_no = Students.clean(row.get("Sr No"))
            if not sr_no.isdigit():
                continue
            if first_numeric is None:
                first_numeric = int(sr_no)
            key = row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), sr_no)
            if key not in completed:
                return int(sr_no)
        return first_numeric or 1

    def selected_run_row_count(self) -> int:
        if hasattr(self, "run_all_pending") and self.run_all_pending.get():
            return sum(
                1
                for row in self.dashboard_rows
                if (row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), row.get("Sr No", "")))
                not in self.dashboard_completed
            )
        try:
            only_sr, _label = self.compute_run_range()
        except Exception:
            return 0
        if not only_sr:
            return len(self.dashboard_rows)
        try:
            selected = Students.parse_sr_numbers(only_sr)
        except Exception:
            return 0
        return sum(1 for row in self.dashboard_rows if Students.clean(row.get("Sr No")) in selected)

    def current_run_rows(self) -> list[dict]:
        rows = self.dashboard_rows
        if not rows:
            try:
                rows, completed = self.load_dashboard_rows()
                self.dashboard_rows = rows
                self.dashboard_completed = completed
            except Exception:
                return []
        only_sr = self.current_run_only_sr
        if not only_sr:
            try:
                only_sr, _label = self.compute_run_range()
            except Exception:
                only_sr = None
        if not only_sr:
            return list(rows)
        try:
            selected = Students.parse_sr_numbers(only_sr)
        except Exception:
            return []
        return [row for row in rows if Students.clean(row.get("Sr No")) in selected]

    def row_is_completed(self, row: dict) -> bool:
        key = row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), row.get("Sr No", ""))
        return key in self.dashboard_completed

    def next_pending_sr_in_current_run(self) -> int | None:
        for row in self.current_run_rows():
            sr_no = Students.clean(row.get("Sr No"))
            if sr_no.isdigit() and not self.row_is_completed(row):
                return int(sr_no)
        return None

    def update_run_range_label(self) -> None:
        try:
            _only_sr, label = self.compute_run_range()
        except Exception as exc:
            label = f"Current run: {exc}"
        self.current_run_label = label
        if hasattr(self, "run_range_text"):
            self.run_range_text.set(label)
        if hasattr(self, "dashboard_range_text"):
            self.dashboard_range_text.set(label)

    def refresh_dashboard(self, reset_run_failures: bool = False) -> None:
        self.dashboard_refresh_job = None
        if reset_run_failures:
            self.dashboard_failed_this_run.clear()
            self.dashboard_current_awc = "-"
            self.dashboard_mapping_status = "Not checked"
            self.current_awc_done = 0
            self.current_awc_total = 0
        try:
            rows, completed = self.load_dashboard_rows()
            self.dashboard_rows = rows
            self.dashboard_completed = completed
            self.dashboard_error = ""
        except Exception as exc:
            self.dashboard_rows = []
            self.dashboard_completed = set()
            self.dashboard_error = str(exc)
        self.update_run_range_label()
        self.update_dashboard_view()

    def load_dashboard_rows(self) -> tuple[list[dict], set[str]]:
        selected_path = Path(self.batch_control.get())
        logs_root = APP_DIR / "logs"
        all_rows: list[dict] = []
        completed_keys: set[str] = set()
        args = self._batch_args()
        batch_rows = self.load_batch_or_direct_student_file(selected_path, args)
        Students.validate_batch_control(batch_rows)
        selected_batches = Students.select_rows_by_batch(batch_rows, self.only_batch.get().strip() or None)
        if not selected_batches:
            return [], set()
        for batch_index, batch_row in enumerate(selected_batches, start=1):
            batch_name = Students.clean(batch_row.get("Batch Name")) or f"Batch {batch_index}"
            excel_path = Students.resolve_batch_path(batch_row.get("Excel File", ""), selected_path)
            sheet = Students.clean(batch_row.get("Sheet")) or "Students"
            rows = Students.normalize_student_rows(Students.load_rows(excel_path, sheet))
            Students.validate_screening_health_rows(rows)
            self._add_dashboard_metadata(rows, batch_name)
            all_rows.extend(rows)
            completed_keys.update(self._completed_keys_for_batch(logs_root, batch_name))
        return all_rows, completed_keys

    def _add_dashboard_metadata(self, rows: list[dict], batch_name: str) -> None:
        for row in rows:
            sr_no = Students.clean(row.get("Sr No"))
            row["_dashboard_batch"] = batch_name
            row["_dashboard_key"] = self.dashboard_key(batch_name, sr_no)

    def _completed_keys_for_batch(self, logs_root: Path, batch_name: str) -> set[str]:
        progress_path = logs_root / Students.slugify(batch_name) / "progress.json"
        completed = Students.load_progress(progress_path)
        return {self.dashboard_key(batch_name, sr_no) for sr_no in completed}

    def update_dashboard_view(self) -> None:
        if not hasattr(self, "card_vars"):
            return
        if self.dashboard_error:
            self.card_vars["total"].set("0")
            self.card_vars["run"].set("0")
            self.card_vars["completed"].set("0")
            self.card_vars["pending"].set("0")
            self.card_vars["failed"].set("0")
            self.card_vars["current"].set("Check file")
            self.card_vars["mapping"].set(self.dashboard_mapping_status)
            self.card_vars["speed"].set("-")
            self.overall_progress["value"] = 0
            self.institute_progress["value"] = 0
            self.overall_progress_text.set("0%")
            self.institute_progress_text.set("0%")
            self.institute_tree.delete(*self.institute_tree.get_children())
            self.institute_tree.insert("", "end", values=(self.dashboard_error, "-", "-", "-", "-", "File error"))
            return

        groups: dict[str, dict] = {}
        order: list[str] = []
        for row in self.dashboard_rows:
            awc = Students.clean(row.get("Institute Name")) or "(no institute)"
            if awc not in groups:
                groups[awc] = {"total": 0, "completed": 0, "failed": 0}
                order.append(awc)
            key = row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), row.get("Sr No", ""))
            groups[awc]["total"] += 1
            if key in self.dashboard_completed:
                groups[awc]["completed"] += 1
            elif key in self.dashboard_failed_this_run:
                groups[awc]["failed"] += 1

        total = len(self.dashboard_rows)
        completed = sum(group["completed"] for group in groups.values())
        failed = sum(group["failed"] for group in groups.values())
        pending = max(total - completed, 0)
        run_count = self.selected_run_row_count()
        self.card_vars["total"].set(str(total))
        self.card_vars["run"].set(str(run_count))
        self.card_vars["completed"].set(str(completed))
        self.card_vars["pending"].set(str(pending))
        self.card_vars["failed"].set(str(failed))
        self.card_vars["current"].set(self.dashboard_current_awc or "-")
        self.card_vars["mapping"].set(self.dashboard_mapping_status)
        self.card_vars["speed"].set(self.run_timing_text)

        overall_pct = int((completed / total) * 100) if total else 0
        self.overall_progress["value"] = overall_pct
        self.overall_progress_text.set(f"{overall_pct}%")

        if self.current_awc_total:
            current_pct = (self.current_awc_done / self.current_awc_total) * 100
        else:
            current_pct = 0
        self.institute_progress["value"] = current_pct
        if self.current_awc_total:
            self.institute_progress_text.set(f"{current_pct:.2f}% ({self.current_awc_done}/{self.current_awc_total})")
        else:
            self.institute_progress_text.set("0%")

        self.institute_tree.delete(*self.institute_tree.get_children())
        current_awc = self.dashboard_current_awc
        if current_awc.startswith("Complete: "):
            current_awc = current_awc.replace("Complete: ", "", 1)
        for awc in order:
            group = groups[awc]
            group_pending = max(group["total"] - group["completed"], 0)
            if group["failed"]:
                status = "Needs review"
            elif group_pending == 0:
                status = "Complete"
            elif group["completed"]:
                status = "In progress"
            else:
                status = "Pending"
            tags = ("current",) if awc == current_awc else ()
            self.institute_tree.insert(
                "",
                "end",
                values=(awc, group["total"], group["completed"], group_pending, group["failed"], status),
                tags=tags,
            )

    def progress_event(self, event: dict) -> None:
        self.after(0, lambda event=dict(event): self.handle_progress_event(event))

    def handle_progress_event(self, event: dict) -> None:
        event_type = event.get("type", "")
        batch_name = Students.clean(event.get("label")) or "Direct Student File"
        sr_no = Students.clean(event.get("sr_no"))
        key = self.dashboard_key(batch_name, sr_no) if sr_no else ""
        if event_type == "batch_start":
            self.dashboard_failed_this_run.clear()
            self.current_awc_done = 0
            self.current_awc_total = 0
            self.run_timing_started_at = time.time()
            self.run_timing_done = 0
            self.run_timing_total = int(event.get("pending") or event.get("selected") or 0)
            self.run_timing_row_seconds.clear()
            self.run_timing_text = "Starting"
        elif event_type == "institute_mapping":
            matched = Students.clean(event.get("matched")) or "0"
            total = Students.clean(event.get("total")) or "0"
            missing = Students.clean(event.get("missing")) or "0"
            if missing != "0":
                self.dashboard_mapping_status = f"{matched}/{total} matched, {missing} missing"
            else:
                self.dashboard_mapping_status = f"{matched}/{total} matched"
        elif event_type in {"awc_start", "row_start", "awc_change_required"}:
            self.dashboard_current_awc = Students.clean(event.get("institute")) or self.dashboard_current_awc
            if event_type in {"awc_start", "awc_change_required"}:
                self.current_awc_done = 0
                self.current_awc_total = int(event.get("pending") or event.get("institute_total") or 0)
            elif event_type == "row_start" and event.get("institute_total"):
                self.current_awc_total = int(event.get("institute_total") or self.current_awc_total or 0)
        elif event_type == "row_saved":
            self.dashboard_current_awc = Students.clean(event.get("institute")) or self.dashboard_current_awc
            if event.get("institute_total"):
                self.current_awc_total = int(event.get("institute_total") or 0)
            if event.get("institute_index"):
                self.current_awc_done = int(event.get("institute_index") or 0)
            if key:
                self.dashboard_completed.add(key)
                self.dashboard_failed_this_run.discard(key)
            self.update_run_timing_from_event(event)
        elif event_type == "row_failed":
            self.dashboard_current_awc = Students.clean(event.get("institute")) or self.dashboard_current_awc
            if event.get("institute_total"):
                self.current_awc_total = int(event.get("institute_total") or 0)
            if key and key not in self.dashboard_completed:
                self.dashboard_failed_this_run.add(key)
            self.update_run_timing_from_event(event)
        elif event_type == "row_skipped":
            self.dashboard_current_awc = Students.clean(event.get("institute")) or self.dashboard_current_awc
            if event.get("institute_total"):
                self.current_awc_total = int(event.get("institute_total") or 0)
            if event.get("institute_index"):
                self.current_awc_done = int(event.get("institute_index") or self.current_awc_done or 0)
            self.update_run_timing_from_event(event)
        elif event_type == "awc_complete":
            self.dashboard_current_awc = f"Complete: {Students.clean(event.get('institute'))}"
            if self.current_awc_total:
                self.current_awc_done = self.current_awc_total
        elif event_type == "batch_complete":
            self.dashboard_current_awc = "Run complete"
        self.update_dashboard_view()
        if event_type in {"row_saved", "row_failed", "batch_complete"}:
            self.refresh_logs()

    def update_run_timing_from_event(self, event: dict) -> None:
        self.run_timing_done += 1
        try:
            row_seconds = float(event.get("row_seconds") or 0)
        except Exception:
            row_seconds = 0
        if row_seconds > 0:
            self.run_timing_row_seconds.append(row_seconds)
        if self.run_timing_row_seconds:
            avg = sum(self.run_timing_row_seconds) / len(self.run_timing_row_seconds)
        elif self.run_timing_started_at and self.run_timing_done:
            avg = (time.time() - self.run_timing_started_at) / self.run_timing_done
        else:
            avg = 0
        remaining = max((self.run_timing_total or 0) - self.run_timing_done, 0)
        if avg and remaining:
            eta_seconds = int(avg * remaining)
            mins, secs = divmod(eta_seconds, 60)
            eta = f"{mins}m {secs}s" if mins else f"{secs}s"
        else:
            eta = "done" if self.run_timing_done else "-"
        self.run_timing_text = f"{avg:.1f}s/row, ETA {eta}" if avg else "-"

    def _build_guide_tab(self) -> None:
        guide = (
            "Local test:\n"
            "1. Choose the sample Excel.\n"
            "2. Click Start Local Test.\n"
            "3. Watch rows appear on the local test form.\n\n"
            "Actual RBSK Screening Health:\n"
            "1. Keep the real screening Excel selected.\n"
            "2. Confirm Login ID. For direct Excel files, use the Login ID field in this tab.\n"
            "3. Keep Autofill saved Login ID/password checked. First time, enter password once and keep Remember password checked.\n"
            "4. Enter Start Sr No and Number of rows for a block, or check Run all pending rows for the full file.\n"
            "5. Keep Live mode, Auto-select school, browser profile reuse, and manual fallback checked.\n"
            "6. The script never adds/registers students.\n"
            "7. Browser opens RBSK. The tool fills Login ID/password if saved.\n"
            "8. Enter CAPTCHA only, click Login, then click OK in the popup.\n"
            "9. On Screening confirm State, District, Block, MHT, Financial Year, and Month.\n"
            "10. The tool selects each school, opens already-added students, clicks the health/doctor icon, fills health details, previews, submits, enters mobile, and confirms.\n"
            "11. If RBSK blocks an automatic step, follow the popup and click OK to continue.\n\n"
            "Important: never use sample data on the actual RBSK website."
        )
        text = tk.Text(self.guide_tab, wrap="word", height=16, font=("Segoe UI", 11), bg="#ffffff", fg="#18212f")
        text.insert("1.0", guide)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Button(self.guide_tab, text="Open Detailed RBSK Steps", command=lambda: self.open_path(APP_DIR / "RBSK_BATCH_STEPS.md")).pack(anchor="w", pady=10)

    def _file_row(self, parent, row: int, label: str, variable: tk.StringVar, filetypes) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=8)
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=row, column=2, sticky="w", pady=8)
        ttk.Button(btn_frame, text="Browse File", command=lambda: self.browse_file(variable, filetypes)).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="Sync Folder", command=lambda: self.browse_folder(variable)).pack(side="left")

    def browse_file(self, variable: tk.StringVar, filetypes) -> None:
        path = filedialog.askopenfilename(initialdir=APP_DIR, filetypes=filetypes)
        if path:
            variable.set(path)
            if hasattr(self, "refresh_dashboard"):
                self.refresh_dashboard()

    def browse_folder(self, variable: tk.StringVar) -> None:
        folder_path = filedialog.askdirectory(initialdir=APP_DIR, title="Select Folder containing Excel Student Lists")
        if folder_path:
            variable.set(folder_path)
            if hasattr(self, "refresh_dashboard"):
                self.refresh_dashboard()

    def update_other_pc_folder(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Please wait", "Stop or finish the current screening process before updating another PC.")
            return
        folder_path = filedialog.askdirectory(
            initialdir=APP_DIR.parent,
            title="Select Existing RBSK Installation Folder on Other PC",
        )
        if not folder_path:
            return
        target = Path(folder_path)
        try:
            summary = destination_license_summary(target)
        except Exception as exc:
            messagebox.showerror("Safe Update Not Started", str(exc))
            return

        confirmation = (
            "Update application files in this existing RBSK folder?\n\n"
            f"Folder: {target}\n"
            f"Existing PC Machine ID: {summary['machine_id']}\n"
            f"Existing licence expiry: {summary['expires_on']}\n"
            f"Existing remaining entries: {summary['remaining_entries']} / {summary['entry_limit']}\n\n"
            "The destination licence key, used/remaining entries, saved password, browser profiles, logs, progress, and Excel files will remain unchanged."
        )
        if not messagebox.askyesno("Confirm Safe Update", confirmation):
            return
        try:
            result = safe_update_installation(APP_DIR, target)
        except Exception as exc:
            messagebox.showerror("Safe Update Failed", str(exc))
            return

        self.show_guided_dialog(
            "Other PC Updated",
            "Application update completed safely.\n\n"
            f"Existing PC Machine ID: {result['machine_id']}\n"
            f"Existing licence expiry: {result['expires_on']}\n"
            f"Existing remaining entries: {result['remaining_entries']} / {result['entry_limit']}\n"
            f"Program files updated: {len(result['copied_files'])}\n\n"
            "The existing licence key and token count were verified unchanged.",
            wait=False,
        )

    def append_output(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.insert("end", text)
        self.output.see("end")
        self.output.configure(state="disabled")

    def clear_output(self) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.configure(state="disabled")

    def _drain_output(self) -> None:
        try:
            while True:
                self.append_output(self.output_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(150, self._drain_output)

    def start_worker(self, title: str, target) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Please wait", "A process is already running.")
            return
        self.clear_output()
        self.append_output(f"{title}\n" + "=" * len(title) + "\n")
        self.worker = threading.Thread(target=self._run_with_output, args=(target,), daemon=True)
        self.worker.start()

    def _run_with_output(self, target) -> None:
        redirect = TextRedirector(self.output_queue)
        try:
            with contextlib.redirect_stdout(redirect), contextlib.redirect_stderr(redirect):
                target()
        except Exception as exc:
            self.output_queue.put(f"\nERROR: {exc}\n")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))

    def start_demo(self) -> None:
        def task():
            args = argparse.Namespace(
                excel=self.demo_excel.get(),
                sheet="Students",
                only_sr=None,
                limit=self._int_or_none(self.demo_limit.get()),
                browser=self.browser.get(),
                demo_log_dir="demo_logs",
                demo_delay=0.7,
                keep_browser_open=self.demo_keep_open.get(),
                allow_insecure_ssl=False,
            )
            Students.run_demo_form(args)

        self.start_worker("Starting client demo", task)

    def start_batch(self) -> None:
        self.screening_health_entry.set(True)
        self.browser_form_fill.set(True)
        try:
            self.refresh_dashboard(reset_run_failures=True)
            self.current_run_only_sr, self.current_run_label = self.compute_run_range()
            self.update_run_range_label()
            self.current_run_snapshot = self.build_run_snapshot()
            if not self.warn_if_selected_file_changed():
                return
        except Exception as exc:
            messagebox.showerror("Row range error", str(exc))
            return

        if not self.dry_run.get():
            try:
                self.ensure_license_available()
            except Exception as exc:
                messagebox.showerror("Licence Required", str(exc))
                self.refresh_license_status()
                return
            if not self.browser_form_fill.get():
                messagebox.showwarning("Browser form fill required", "Live RBSK entry should use browser form filling. The option has been re-enabled.")
                self.browser_form_fill.set(True)
            if self.current_run_snapshot and self.current_run_snapshot.get("all_pending"):
                ok_all = messagebox.askyesno(
                    "Run all pending rows?",
                    "Run all pending rows is checked.\n\n"
                    "This will validate and enter every pending row in the selected file, not just 50/200/1000 rows.\n\n"
                    "Continue with ALL pending rows?",
                )
                if not ok_all:
                    return
            if (
                self.current_run_snapshot
                and not self.current_run_snapshot.get("all_pending")
                and self.current_run_snapshot.get("row_count")
                and self.current_run_snapshot.get("row_count") == self.current_run_snapshot.get("total_rows")
                and self.current_run_snapshot.get("total_rows", 0) > 200
            ):
                ok_full = messagebox.askyesno(
                    "Whole file selected",
                    "The current settings select the whole Excel file.\n\n"
                    "For a controlled block, enter Start Sr No and Number of rows, or click Start Next 50/200/1000.\n\n"
                    "Continue with the whole file?",
                )
                if not ok_full:
                    return
            preflight = self.live_start_summary(self.current_run_snapshot)
            ok = messagebox.askyesno(
                "Live RBSK screening entry",
                "Dry run is OFF. This will update Screening Health only for already-added students.\n\n"
                f"{preflight}\n\n"
                "It will not add/register new students.",
            )
            if not ok:
                return

        def task():
            args = self._batch_args()
            self.run_batch_with_gui(args)

        self.start_worker("Starting RBSK batch", task)

    def start_screening_health(self) -> None:
        self.screening_health_entry.set(True)
        self.browser_form_fill.set(True)
        self.start_batch()

    def start_direct_portal_screening(self) -> None:
        self.screening_health_entry.set(True)
        self.browser_form_fill.set(True)
        try:
            self.ensure_license_available()
        except Exception as exc:
            messagebox.showerror("Licence Required", str(exc))
            self.refresh_license_status()
            return
        inst_type = self.screening_institution_type.get()
        if Students.is_awc_institution(inst_type):
            msg = (
                "This mode screens AWC children (0–5 years) directly from the RBSK web portal without needing an Excel file.\n\n"
                "• Institution Type: AWC (Anganwadi Centre)\n"
                "• Child age range: 0–5 years\n"
                "• Reads Age directly from the portal page.\n"
                "• Calculates Weight & Height automatically from Age (WHO Child Growth Standards).\n"
                "• Sets Blood Pressure with Diastolic in the 65–80 range.\n"
                "• Sets Eye vision to 6/6.\n"
                "• Mobile Number: NOT asked and NOT entered.\n\n"
                "Do you want to start Direct Portal Screening for AWC?"
            )
        else:
            msg = (
                "This mode screens students directly from the active RBSK web portal without needing an Excel file.\n\n"
                "• Reads Age directly from the portal page (e.g. Age: 10 years 0 months).\n"
                "• Calculates Weight & Height automatically from Age.\n"
                "• Sets Blood Pressure with Diastolic in the 65–80 range.\n"
                "• Sets Eye vision to 6/6.\n"
                "• Prompts for fixed Mobile Number per school/college.\n\n"
                "Do you want to start Direct Portal Screening?"
            )
        ok = messagebox.askyesno("Direct Portal Screening (No Excel)", msg)
        if not ok:
            return

        def task():
            args = self._batch_args()
            self.run_direct_portal_with_gui(args)

        self.start_worker("Direct Portal Screening", task)

    def run_direct_portal_with_gui(self, args: argparse.Namespace) -> None:
        login_id = getattr(args, "default_login_id", "") or "MHT-1271251"
        args.browser_profile_dir = None
        args.use_browser_profile = False
        password = None
        if not getattr(args, "manual_login", False):
            if not login_id:
                raise Students.ImportErrorWithHelp("Login ID is blank. Fill Login ID in the launcher.")
            password = self.password_for_run(args, login_id)
            if getattr(args, "remember_credentials", False):
                self.save_login_password(login_id, password)
                self.after(0, self.refresh_saved_login_status)

        driver, _session, _storage = self.capture_session_gui(args, login_id, password)
        self.popup_browser_driver = driver
        try:
            try:
                Students.run_direct_portal_screening(driver, args)
            except Exception:
                if Students.browser_is_open(driver):
                    raise
                print("RBSK browser was closed. Direct screening stopped without showing any popups.")
        finally:
            self.popup_browser_driver = None
            try:
                driver.quit()
            except Exception:
                pass

    def check_institute_names(self) -> None:
        try:
            self.refresh_dashboard(reset_run_failures=True)
            self.current_run_only_sr, self.current_run_label = self.compute_run_range()
            self.update_run_range_label()
            self.current_run_snapshot = self.build_run_snapshot()
            if not self.warn_if_selected_file_changed():
                return
        except Exception as exc:
            messagebox.showerror("Row range error", str(exc))
            return

        ok = messagebox.askyesno(
            "Check Institute Names",
            "This will open RBSK and compare Excel AWC NAME values with the portal institute list.\n\n"
            "No student records will be saved.",
        )
        if not ok:
            return

        def task():
            args = self._batch_args()
            args.dry_run = False
            args.browser_form_fill = True
            args.auto_select_institute = True
            args.mapping_check_only = True
            self.run_batch_with_gui(args)

        self.start_worker("Checking institute names", task)

    def prompt_school_mobile(self, school_name: str) -> str:
        # When AWC is selected, skip mobile number prompt entirely
        if Students.is_awc_institution(self.screening_institution_type.get()):
            return ""
        driver = getattr(self, "popup_browser_driver", None)
        if driver is not None and not Students.browser_is_open(driver):
            return ""

        result_q: queue.Queue[str] = queue.Queue()

        def _dialog():
            active_driver = getattr(self, "popup_browser_driver", None)
            if active_driver is not None and not Students.browser_is_open(active_driver):
                result_q.put("")
                return

            monitor_job = None

            def close_if_browser_closed():
                nonlocal monitor_job
                current_driver = getattr(self, "popup_browser_driver", None)
                if current_driver is not None and not Students.browser_is_open(current_driver):
                    for child in self.winfo_children():
                        try:
                            if isinstance(child, tk.Toplevel) and child.title() == "School Mobile Number":
                                child.destroy()
                        except Exception:
                            pass
                    return
                monitor_job = self.after(250, close_if_browser_closed)

            while True:
                monitor_job = self.after(250, close_if_browser_closed)
                ans = simpledialog.askstring(
                    "School Mobile Number",
                    f"Direct Portal Mobile Mode is active.\n\nPlease enter 10-digit Mobile Number for school:\n\n{school_name}",
                    parent=self,
                )
                if monitor_job is not None:
                    try:
                        self.after_cancel(monitor_job)
                    except Exception:
                        pass
                current_driver = getattr(self, "popup_browser_driver", None)
                if current_driver is not None and not Students.browser_is_open(current_driver):
                    result_q.put("")
                    return
                if ans is None:
                    result_q.put("")
                    return
                digits = "".join(ch for ch in Students.clean(ans) if ch.isdigit())
                if len(digits) >= 10:
                    result_q.put(digits[-10:])
                    return
                messagebox.showwarning("Invalid mobile number", "Enter a valid 10-digit mobile number.", parent=self)

        self.after(0, _dialog)
        return result_q.get()

    def _batch_args(self) -> argparse.Namespace:
        snapshot_only_sr = self.current_run_snapshot.get("only_sr") if self.current_run_snapshot else self.current_run_only_sr
        return argparse.Namespace(
            batch_control=self.batch_control.get(),
            batch_sheet="Batches",
            batch_limit=None,
            only_batch=self.only_batch.get().strip() or None,
            logs_dir="logs",
            stop_on_batch_error=self.stop_on_error.get(),
            excel="",
            sheet="Students",
            browser=self.browser.get(),
            dry_run=self.dry_run.get(),
            no_browser=False,
            only_sr=snapshot_only_sr,
            limit=None,
            resume=True,
            delay=0.25,
            timeout=120,
            keep_browser_open=False,
            success_log="success.csv",
            failed_log="failed.csv",
            skipped_log="skipped.csv",
            progress="progress.json",
            state_code=None,
            district_code=None,
            health_block_code=None,
            institution_type="1",
            institution_code=None,
            created_by=None,
            allow_insecure_ssl=self.allow_insecure_ssl.get(),
            manual_login=self.manual_login.get(),
            default_login_id=self.default_login_id.get().strip(),
            entered_password=self.login_password.get(),
            entered_password_login_id=self.default_login_id.get().strip(),
            remember_credentials=self.remember_credentials.get(),
            browser_profile_dir=None,
            use_browser_profile=False,
            browser_form_fill=self.browser_form_fill.get(),
            screening_health_entry=True,
            screening_month=self.screening_month.get().strip(),
            screening_institution_type=self.screening_institution_type.get(),
            manual_form_open=self.manual_form_open.get(),
            unattended_after_start=self.unattended_after_start.get(),
            auto_select_institute=self.auto_select_institute.get(),
            mapping_check_only=False,
            institute_aliases=str(APP_DIR / "institute_aliases.csv"),
            speed_mode=self.speed_mode.get(),
            mobile_source=self.mobile_source.get(),
            custom_school_mobile=self.custom_school_mobile.get().strip(),
            prompt_mobile_callback=self.prompt_school_mobile,
            run_snapshot=self.current_run_snapshot.copy() if isinstance(self.current_run_snapshot, dict) else self.current_run_snapshot,
            manual_form_open_callback=self.info_blocking,
            progress_callback=self.progress_event,
            license_check_callback=self.ensure_license_available,
            license_success_callback=self.record_license_success,
        )

    def run_batch_with_gui(self, args: argparse.Namespace) -> int:
        control_path = Path(args.batch_control)
        batch_rows = self.load_batch_or_direct_student_file(control_path, args)
        Students.validate_batch_control(batch_rows)
        selected_batches = Students.select_rows_by_batch(batch_rows, args.only_batch)
        if args.batch_limit:
            selected_batches = selected_batches[: args.batch_limit]
        if not selected_batches:
            names = ", ".join(Students.available_batch_names(batch_rows)) or "none"
            raise Students.ImportErrorWithHelp(f"No batch rows selected. Leave Only Batch Name blank or use one of: {names}")

        logs_root = APP_DIR / args.logs_dir
        summary_path = logs_root / "batch_summary.csv"
        overall_exit = 0
        run_saved = 0
        run_duplicates = 0
        run_failed = 0
        run_setup_failed = 0
        setup_error_message = ""
        run_skipped = 0
        run_skipped_this_run = 0

        print(f"Batch control: {control_path}")
        print(f"Selected centers: {len(selected_batches)}")

        for batch_index, batch_row in enumerate(selected_batches, start=1):
            batch_name = Students.clean(batch_row.get("Batch Name")) or f"Batch {batch_index}"
            login_id = Students.clean(batch_row.get("Login ID")) or Students.clean(getattr(args, "default_login_id", ""))
            batch_log_dir = logs_root / Students.slugify(batch_name)
            batch_args = Students.batch_arg_namespace(args, batch_row, control_path, batch_log_dir)
            selected_sr_values: set[str] = set()
            completed_before: set[str] = set()
            batch_args.browser_profile_dir = None
            batch_args.use_browser_profile = False

            print("\n" + "=" * 72)
            print(f"Center {batch_index}/{len(selected_batches)}: {batch_name}")
            print(f"Login ID: {login_id}")
            print(f"Excel: {batch_args.excel}")
            print("Browser session/profile storage: disabled")

            try:
                rows = Students.normalize_student_rows(Students.load_rows(Path(batch_args.excel), batch_args.sheet))
                Students.validate_screening_health_rows(rows)
                self.write_run_metadata(batch_log_dir, self.build_run_metadata(Path(batch_args.excel), batch_args.sheet, rows))
                rows = Students.select_rows(rows, batch_args.only_sr, batch_args.limit)
                if not rows:
                    raise Students.ImportErrorWithHelp("No student rows selected for this center.")
                selected_sr_values = {Students.clean(row.get("Sr No")) for row in rows if Students.clean(row.get("Sr No"))}
                completed_before = Students.load_progress(Path(batch_args.progress)) & selected_sr_values

                session = None
                storage = {}
                driver = None
                if not batch_args.dry_run:
                    password = None
                    if not getattr(batch_args, "manual_login", False):
                        if not login_id:
                            raise Students.ImportErrorWithHelp("Login ID is blank. Fill Login ID in the launcher or batch control file.")
                        password = self.password_for_run(batch_args, login_id)
                        if getattr(batch_args, "remember_credentials", False):
                            self.save_login_password(login_id, password)
                            self.after(0, self.refresh_saved_login_status)
                    driver, session, storage = self.capture_session_gui(batch_args, login_id, password)

                context = Students.infer_context(Students.parse_storage(storage), batch_args)
                uses_row_institute = any(
                    Students.clean(row.get("Institute Name"))
                    or Students.clean(row.get("Institution Code") or row.get("Institute Code"))
                    for row in rows
                )
                live_browser_form = bool(batch_args.browser_form_fill and not batch_args.dry_run)
                Students.ensure_context(
                    context,
                    require_created_by=not batch_args.dry_run and not live_browser_form,
                    require_institution_code=not uses_row_institute and not live_browser_form,
                )
                batch_args.portal_session = session
                batch_args.portal_context = context
                if batch_args.dry_run:
                    print(f"Dry run Screening-only validation passed for {len(rows)} selected row(s). No students added.")
                    result = {"saved": 0, "duplicates": 0, "failed": 0, "skipped": 0, "skipped_this_run": 0, "exit_code": 0}
                else:
                    result = Students.process_screening_health_rows_browser(driver, rows, batch_args, label=batch_name)
                    if batch_args.keep_browser_open:
                        self.retain_browser_driver(driver)
                    else:
                        driver.quit()
                run_saved += int(result.get("saved", 0))
                run_duplicates += int(result.get("duplicates", 0))
                run_failed += int(result.get("failed", 0))
                run_skipped += int(result.get("skipped", 0))
                run_skipped_this_run += int(result.get("skipped_this_run", 0))
                status = "success" if result["exit_code"] == 0 and result["failed"] == 0 else "partial_or_failed"
                Students.append_batch_summary(
                    summary_path,
                    {
                        "Batch Name": batch_name,
                        "Login ID": login_id,
                        "Excel File": batch_args.excel,
                        "status": status,
                        "saved": result["saved"],
                        "duplicates": result.get("duplicates", 0),
                        "failed": result["failed"],
                        "skipped": result["skipped"],
                        "skipped_this_run": result.get("skipped_this_run", 0),
                        "message": "",
                    },
                )
                if result["exit_code"] != 0 or result["failed"]:
                    overall_exit = int(result["exit_code"]) or 1
                    if args.stop_on_batch_error:
                        return overall_exit
            except Exception as exc:
                try:
                    if 'driver' in locals() and driver:
                        if getattr(batch_args, "keep_browser_open", False) and not Students.browser_session_error(exc):
                            self.retain_browser_driver(driver)
                        else:
                            driver.quit()
                except Exception:
                    pass
                exc_text = str(exc)
                overall_exit = 1
                is_setup_failure = (
                    "Institute mapping check failed" in exc_text
                    or "Could not read institute list from RBSK" in exc_text
                    or "institute_mapping_missing.csv" in exc_text
                    or self.is_browser_profile_start_error(exc)
                    or Students.browser_session_error(exc)
                    or "browser was closed or stopped" in exc_text.lower()
                )
                if is_setup_failure:
                    run_setup_failed += 1
                    setup_error_message = exc_text
                else:
                    run_failed += 1
                if selected_sr_values:
                    completed_after = Students.load_progress(Path(batch_args.progress)) & selected_sr_values
                    recovered_saved = len(completed_after - completed_before)
                    recovered_skipped = len(completed_before)
                    if recovered_saved:
                        run_saved += recovered_saved
                        print(f"Recovered saved count from progress.json after interruption: {recovered_saved}")
                    if recovered_skipped:
                        run_skipped += recovered_skipped
                print(f"Center failed: {exc}")
                Students.append_batch_summary(
                    summary_path,
                    {
                        "Batch Name": batch_name,
                        "Login ID": login_id,
                        "Excel File": Students.clean(batch_row.get("Excel File")),
                        "status": "failed",
                        "saved": 0,
                        "failed": 0,
                        "skipped": 0,
                        "message": str(exc),
                    },
                )
                if args.stop_on_batch_error:
                    return overall_exit

        print("\nBatch run complete.")
        print(f"Summary: {summary_path}")
        self.progress_event({"type": "batch_complete", "label": "All", "saved": run_saved, "duplicates": run_duplicates, "failed": run_failed, "skipped": run_skipped, "skipped_this_run": run_skipped_this_run})
        finished_snapshot = getattr(args, "run_snapshot", None)
        self.after(0, lambda: self.show_run_complete(run_saved, run_failed, run_skipped, summary_path, run_setup_failed, setup_error_message, finished_snapshot, run_duplicates, run_skipped_this_run))
        return overall_exit

    def load_batch_or_direct_student_file(self, selected_path: Path, args: argparse.Namespace):
        if selected_path.is_dir():
            excel_files = sorted(
                [
                    f for f in selected_path.glob("*")
                    if f.is_file() and f.suffix.lower() in (".xlsx", ".xls", ".csv") and not f.name.startswith("~$")
                ],
                key=lambda x: x.name.lower()
            )
            if not excel_files:
                raise Students.ImportErrorWithHelp(f"No Excel (.xlsx/.xls) or CSV files found in folder: {selected_path}")
            print(f"Syncing folder: Found {len(excel_files)} Excel file(s) in {selected_path.name}:")
            all_batch_rows = []
            for ef in excel_files:
                print(f"  • {ef.name}")
                try:
                    sub_rows = self.load_batch_or_direct_student_file(ef, args)
                    all_batch_rows.extend(sub_rows)
                except Exception as exc:
                    print(f"Warning: Skipping file {ef.name}: {exc}")
            if not all_batch_rows:
                raise Students.ImportErrorWithHelp(f"No valid student Excel files found in folder: {selected_path}")
            return all_batch_rows

        try:
            rows = Students.load_rows(selected_path, args.batch_sheet)
            Students.validate_batch_control(rows)
            return rows
        except Exception as batch_error:
            direct_sheet = ""
            for candidate_sheet in ("Students", "Sheet1"):
                try:
                    student_rows = Students.normalize_student_rows(Students.load_rows(selected_path, candidate_sheet))
                    Students.validate_screening_health_rows(student_rows)
                    direct_sheet = candidate_sheet
                    break
                except Exception:
                    continue
            if not direct_sheet:
                raise batch_error
            batch_name = self.direct_batch_name_for_file(selected_path)
            print("Selected file looks like a student data Excel, not a batch control file.")
            print(f"Using it as one direct batch named: {batch_name}")
            return [
                {
                    "Batch Name": batch_name,
                    "Login ID": "",
                    "Excel File": str(selected_path),
                    "Sheet": direct_sheet,
                    "State Code": "",
                    "District Code": "",
                    "Health Block Code": "",
                    "Institution Type": "1",
                    "Institution Code": "",
                }
            ]

    def capture_session_gui(self, args: argparse.Namespace, login_id: str, password: str | None):
        profile_dir = getattr(args, "browser_profile_dir", None)
        recovery_notice = ""
        reused_driver = False
        driver = self.reusable_browser_driver()
        if driver is not None:
            reused_driver = True
            print("Reusing the open RBSK browser window for this run.")
        else:
            driver = None
        try:
            if driver is None:
                driver = Students.start_browser(args.browser, profile_dir)
        except Exception as exc:
            if not profile_dir or not self.is_browser_profile_start_error(exc):
                raise
            recovery_dir = self.make_recovery_profile_dir(login_id)
            print("Saved browser profile could not start after the interrupted run.")
            print(f"Original browser profile: {profile_dir}")
            print(f"Recovery browser profile: {recovery_dir}")
            print("Opening RBSK with a fresh browser session for this Login ID.")
            try:
                driver = Students.start_browser(args.browser, recovery_dir)
            except Exception:
                raise exc
            args.browser_profile_dir = str(recovery_dir)
            self.set_browser_profile_override(login_id, recovery_dir)
            recovery_notice = (
                "The saved Chrome/Edge session was locked or crashed from the stopped script.\n\n"
                "The helper opened a fresh browser session for this Login ID.\n\n"
                "Do this once:\n"
                "1. Login again if RBSK asks.\n"
                "2. Enter CAPTCHA manually.\n"
                "3. Continue normally.\n\n"
                "Your saved password, completed progress, success logs, and failed logs were not deleted."
            )
        try:
            if reused_driver:
                try:
                    current_url = str(driver.current_url or "").lower()
                except Exception:
                    current_url = ""
                if "rbsk.mohfw.gov.in" not in current_url:
                    driver.get(Students.APP_URL)
            else:
                driver.get(Students.APP_URL)
            if recovery_notice:
                self.info_blocking("Browser Session Recovered", recovery_notice)
            auto_filled = False
            if not reused_driver and not getattr(args, "manual_login", False) and password:
                Students.fill_login_fields(driver, login_id, password)
                auto_filled = True

            reused_logged_in = False
            if reused_driver:
                try:
                    time.sleep(0.5)
                    page_text = Students.clean(driver.execute_script("return document.body.innerText || document.body.textContent || '';")).lower()
                    reused_logged_in = (
                        ("welcome" in page_text or "screening" in page_text or "dashboard" in page_text)
                        and "captcha" not in page_text
                    )
                except Exception:
                    reused_logged_in = False

            if reused_logged_in:
                print("Open RBSK browser session is still logged in; continuing in the same screen.")
            elif auto_filled:
                login_message = (
                    "Browser is open and Login ID/password were filled automatically.\n\n"
                    "Do only this now:\n"
                    "1. Enter CAPTCHA manually.\n"
                    "2. Click Login in RBSK.\n"
                    "3. Wait until the portal dashboard/page is visible.\n"
                    "4. Then click OK here.\n\n"
                    "If RBSK already opens logged in, just confirm the dashboard is visible and click OK."
                )
            else:
                login_message = (
                    "Browser is open.\n\n"
                    "1. Type Login ID manually.\n"
                    "2. Type password manually.\n"
                    "3. Solve CAPTCHA manually.\n"
                    "4. Click login in RBSK.\n"
                    "5. Wait until the portal dashboard/page is visible.\n"
                    "6. Then click OK here."
                )

            if not reused_logged_in:
                self.info_blocking("Login / CAPTCHA", login_message)

            Students.browser_open_screening_page(driver, speed_mode=getattr(args, "speed_mode", ""))
            if Students.clean(getattr(args, "screening_month", "")):
                print(f"Screening month selected in launcher: {args.screening_month}")
                print("The tool will select this month and open the school list automatically.")
            else:
                self.info_blocking(
                    "Confirm Screening Page",
                    "On the Screening page:\n\n1. Confirm State, District, Block, MHT, and Financial Year.\n2. Select the Month manually.\n3. Click OK here.\n\nThe tool will select School, click Search, open each school row, select each already-added student, click the health/doctor icon, fill health details from Excel, set both eye vision values to 6/6 unless Excel has values, preview, submit, enter mobile number, and confirm.\n\nIt will not add/register new students.",
                )
            args.screening_page_prepared = True
            time.sleep(1)
            session, storage = Students.collect_browser_session(driver)
            if getattr(args, "browser_profile_dir", None):
                self.set_browser_profile_override(login_id, Path(args.browser_profile_dir))
            return driver, session, storage
        except Exception:
            if reused_driver:
                self.retain_browser_driver(driver)
            else:
                driver.quit()
            raise

    def dialog_step(self, title: str, message: str) -> tuple[str, str]:
        text = f"{title}\n{message}".lower()
        if "failed" in text or "error" in text or "could not" in text:
            return "Needs attention", "#dc2626"
        if "run complete" in text or "saved" in text and "failed this run" in text:
            return "Run complete", "#059669"
        if "manual login" in text or "captcha" in text or "login / captcha" in text:
            return "CAPTCHA only login", "#2563eb"
        if "institute names" in text or "mapping" in text:
            return "Screening Check", "#6d28d9"
        if "confirm student master" in text:
            return "Confirm Screening Page", "#0f766e"
        if "click proceed" in text or "register child" in text:
            return "Click Proceed", "#7c3aed"
        if "same institute" in text or "add student" in text or "add child" in text:
            return "Open Health Form", "#2563eb"
        if "select" in text and "institute" in text:
            return "Select institute once", "#0f766e"
        return title, "#2563eb"

    def latest_pending_failure_lines(self, limit: int = 5) -> list[str]:
        try:
            self.refresh_logs()
            rows = [row for row in self.log_rows if row.get("status") in {"Pending failed", "Skipped this run"}]
        except Exception:
            rows = []
        if not rows:
            return []
        lines: list[str] = []
        for row in rows[-limit:]:
            sr_no = Students.clean(row.get("sr_no"))
            name = Students.clean(row.get("name"))
            awc = Students.clean(row.get("awc"))
            reason = Students.clean(row.get("reason"))
            error = Students.clean(row.get("error"))
            if len(error) > 180:
                error = error[:177] + "..."
            lines.append(f"Sr No {sr_no}: {name} | {awc} | {reason} | {error}")
        return lines

    def show_guided_dialog(self, title: str, message: str, wait: bool = True) -> None:
        step, color = self.dialog_step(title, message)
        s = self.ui_scale
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(bg="#f8fafc")
        dialog_w = min(int(620 * s), screen_w - 40)
        dialog_h = min(int(520 * s), screen_h - 80)
        dialog.geometry(f"{dialog_w}x{dialog_h}+20+20")
        dialog.minsize(min(int(560 * s), screen_w - 40), min(int(420 * s), screen_h - 80))
        dialog.resizable(True, True)
        dialog.attributes("-topmost", True)

        header = tk.Frame(dialog, bg=color)
        header.pack(fill="x")
        tk.Label(header, text=step, bg=color, fg="#ffffff", font=("Segoe UI", scaled_font_size(18, s), "bold")).pack(anchor="w", padx=int(22 * s), pady=(int(18 * s), int(4 * s)))
        tk.Label(header, text=title, bg=color, fg="#e0f2fe", font=("Segoe UI", scaled_font_size(10, s), "bold")).pack(anchor="w", padx=int(22 * s), pady=(0, int(16 * s)))

        footer = tk.Frame(dialog, bg="#f8fafc")
        footer.pack(side="bottom", fill="x", padx=22, pady=(0, 18))
        ok_button = tk.Button(
            footer,
            text="OK",
            command=dialog.destroy,
            bg=color,
            fg="#ffffff",
            activebackground=color,
            activeforeground="#ffffff",
            font=("Segoe UI", 12, "bold"),
            relief="flat",
            padx=22,
            pady=10,
            width=12,
            takefocus=True,
            cursor="hand2",
        )
        ok_button.pack(side="right")
        tk.Button(
            footer,
            text="Minimize",
            command=dialog.iconify,
            bg="#e2e8f0",
            fg="#0f172a",
            activebackground="#cbd5e1",
            activeforeground="#0f172a",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            padx=18,
            pady=10,
            width=12,
            takefocus=True,
            cursor="hand2",
        ).pack(side="left")

        body = tk.Frame(dialog, bg="#f8fafc")
        body.pack(side="top", fill="both", expand=True, padx=22, pady=18)
        message_box = tk.Text(body, wrap="word", height=10, font=("Segoe UI", 12), bg="#ffffff", fg="#0f172a", relief="solid", borderwidth=1, padx=14, pady=12)
        message_box.insert("1.0", message)
        message_box.configure(state="disabled")
        message_box.pack(fill="both", expand=True)

        dialog.bind("<Return>", lambda _event: dialog.destroy())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        ok_button.focus_set()
        dialog.focus_force()
        if wait:
            self.wait_window(dialog)

    def show_run_complete(
        self,
        saved: int,
        failed: int,
        skipped: int,
        summary_path: Path,
        setup_failed: int = 0,
        setup_error: str = "",
        snapshot: dict | None = None,
        duplicates: int = 0,
        skipped_this_run: int = 0,
    ) -> None:
        self.play_completion_sound()
        completed_now = saved + duplicates + skipped
        self.refresh_dashboard(reset_run_failures=False)
        snapshot = snapshot or self.current_run_snapshot or self.build_run_snapshot()
        run_total = int(snapshot.get("row_count") or 0)
        keys = snapshot.get("keys") or []
        run_completed = sum(1 for key in keys if key in self.dashboard_completed)
        next_sr = None
        for row in snapshot.get("rows") or []:
            key = row.get("_dashboard_key") or self.dashboard_key(row.get("_dashboard_batch", ""), row.get("Sr No", ""))
            sr_no = Students.clean(row.get("Sr No"))
            if sr_no.isdigit() and key not in self.dashboard_completed:
                next_sr = int(sr_no)
                break
        next_line = ""
        if next_sr is not None:
            next_line = f"\nSuggested next Start Sr No for this block: {next_sr}\n"
        elif run_total:
            next_line = "\nSelected block is complete. Use Start Next 50/200/1000 for the next block.\n"
        message = (
            f"Completed / already done: {completed_now}\n"
            f"Saved in this run: {saved}\n"
            f"Already in RBSK this run: {duplicates}\n"
            f"Skipped already completed: {skipped}\n"
            f"Skipped this run for review: {skipped_this_run}\n"
            f"Failed this run: {failed}\n\n"
            f"Setup / mapping problems: {setup_failed}\n"
            f"Selected block completed: {run_completed}/{run_total}\n"
            f"Summary file:\n{summary_path}\n\n"
            f"{next_line}"
        )
        if setup_failed:
            setup_lower = setup_error.lower()
            if "institute" in setup_lower or "mapping" in setup_lower or "awc name" in setup_lower:
                missing_path = self.latest_missing_report_path() or APP_DIR / "logs" / "Direct_Student_File" / "institute_mapping_missing.csv"
                missing_lines = self.missing_institute_names_from_report(missing_path)
                message += "Institute name mapping failed before entry. No student rows were attempted for the unresolved block.\n"
                if missing_lines:
                    message += "\nMissing AWC name(s):\n" + "\n".join(f"- {name}" for name in missing_lines[:12]) + "\n"
                message += f"\nReport:\n{missing_path}\n\nNext: click Review Missing Institute Names, fix aliases, then run again."
            else:
                message += (
                    "Browser/session stopped before the next row was confirmed. No extra student failure was recorded.\n\n"
                    f"Reason:\n{setup_error}\n\n"
                    "Next: close old Chrome/Edge windows, click Reset Browser Session for this Login ID if needed, "
                    "then run again from the suggested Sr No. Completed rows will be skipped."
                )
        elif failed or skipped_this_run:
            failure_lines = self.latest_pending_failure_lines()
            if failure_lines:
                message += "Latest review row detail(s):\n" + "\n".join(f"- {line}" for line in failure_lines) + "\n\n"
            message += "Next: check the failed/skipped rows in the dashboard/logs, fix the issue or rerun that range. Completed rows will be skipped."
        else:
            message += "Next: keep Number of rows or increase it when you are ready for the next larger batch."
        title = "Run Complete"
        if setup_failed:
            title = "Run Stopped Before Entry"
        elif failed:
            title = "Run Finished With Failed Rows"
        elif skipped_this_run:
            title = "Run Complete With Skipped Rows"
        self.show_guided_dialog(title, message, wait=False)

    def play_completion_sound(self) -> None:
        try:
            import winsound

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            try:
                self.bell()
            except Exception:
                pass

    def info_blocking(self, title: str, message: str) -> None:
        event = threading.Event()

        def show():
            self.show_guided_dialog(title, message, wait=True)
            event.set()

        self.after(0, show)
        event.wait()

    def _int_or_none(self, value: str) -> int | None:
        value = value.strip()
        if not value:
            return None
        return int(value)

    def open_path(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True) if path.suffix == "" else None
        try:
            import os

            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))


if __name__ == "__main__":
    check_cryptography_at_startup()
    EasyLauncher().mainloop()

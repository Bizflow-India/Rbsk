import base64
import tkinter as tk
from datetime import date, timedelta
from tkinter import messagebox, ttk

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.padding import PKCS7


LICENSE_AES_KEY = b"RBSKScreen2026SecureKey@#1234567"
LICENSE_AES_IV = b"InitVector123456"
DEFAULT_ENTRY_LIMIT = 10000


def add_months(value: date, months: int) -> date:
    month = value.month - 1 + months
    year = value.year + month // 12
    month = month % 12 + 1
    days = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    return date(year, month, min(value.day, days[month - 1]))


def generate_activation_key(machine_id: str, expiry_date: date, entry_limit: int = DEFAULT_ENTRY_LIMIT) -> str:
    payload = f"{machine_id.strip().upper()}|{expiry_date:%Y-%m-%d}|{int(entry_limit)}"
    padder = PKCS7(128).padder()
    padded = padder.update(payload.encode("utf-8")) + padder.finalize()
    cipher = Cipher(algorithms.AES(LICENSE_AES_KEY), modes.CBC(LICENSE_AES_IV))
    encryptor = cipher.encryptor()
    encrypted = encryptor.update(padded) + encryptor.finalize()
    return base64.b64encode(encrypted).decode("ascii")


class KeyGenerator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RBSK Monthly Key Generator")
        self.geometry("560x360")
        self.resizable(False, False)
        self.machine_id = tk.StringVar()
        self.months = tk.StringVar(value="1")
        self.entry_limit = tk.StringVar(value=str(DEFAULT_ENTRY_LIMIT))
        self.output_key = tk.StringVar()
        self._build()

    def _build(self):
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Generate Activation Key", font=("Segoe UI", 15, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))
        ttk.Label(frame, text="Client Machine ID").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.machine_id, width=48).grid(row=1, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(frame, text="Validity").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Combobox(frame, textvariable=self.months, values=("1", "3", "6", "12"), state="readonly", width=10).grid(row=2, column=1, sticky="w", pady=6)
        ttk.Label(frame, text="month(s)").grid(row=2, column=1, sticky="w", padx=(92, 0), pady=6)

        ttk.Label(frame, text="Successful Entry Limit").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.entry_limit, width=12).grid(row=3, column=1, sticky="w", pady=6)

        ttk.Button(frame, text="Generate Key", command=self.generate).grid(row=4, column=0, sticky="w", pady=(16, 10))
        ttk.Button(frame, text="Copy Key", command=self.copy_key).grid(row=4, column=1, sticky="w", pady=(16, 10))

        ttk.Label(frame, text="Generated Key").grid(row=5, column=0, sticky="nw", pady=6)
        output = tk.Text(frame, height=5, width=56, wrap="word")
        output.grid(row=5, column=1, columnspan=2, sticky="ew", pady=6)
        self.output = output
        frame.columnconfigure(1, weight=1)

    def generate(self):
        machine_id = self.machine_id.get().strip().upper()
        if not machine_id:
            messagebox.showerror("Missing Machine ID", "Enter the client's Machine ID.")
            return
        try:
            months = int(self.months.get())
            limit = int(self.entry_limit.get())
            if limit <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid value", "Validity and entry limit must be positive numbers.")
            return
        expiry = add_months(date.today(), months)
        key = generate_activation_key(machine_id, expiry, limit)
        self.output.delete("1.0", "end")
        self.output.insert("1.0", key)

    def copy_key(self):
        key = self.output.get("1.0", "end").strip()
        if not key:
            return
        self.clipboard_clear()
        self.clipboard_append(key)
        messagebox.showinfo("Copied", "Activation key copied.")


if __name__ == "__main__":
    KeyGenerator().mainloop()

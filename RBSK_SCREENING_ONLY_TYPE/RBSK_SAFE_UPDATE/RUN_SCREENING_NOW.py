"""
RUN_SCREENING_NOW.py
====================
Simple standalone script to run Direct Portal Screening on all students
in the currently open school list.

Usage:
  1. Double-click this script
  2. Browser opens → Enter CAPTCHA manually → Click Login
  3. Click OK in the popup
  4. The tool navigates to Screening, asks you to select Month/School
  5. Enter mobile number when asked
  6. It fills Weight, Height, BP, Vision for ALL unscreened students automatically

Requirements:
  - selenium must be installed: pip install selenium
  - Chrome or Edge browser
"""
import argparse
import sys
import time
from pathlib import Path

# Add current directory to path
APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import Students


def main():
    print("=" * 60)
    print("   RBSK Direct Portal Screening - Quick Start Script")
    print("=" * 60)
    print()

    # --- Ask for browser ---
    browser = "chrome"
    try:
        choice = input("Browser? [1] Chrome  [2] Edge  (default=1): ").strip()
        if choice == "2":
            browser = "edge"
    except Exception:
        pass

    # --- Ask for login ID ---
    login_id = "MHT-1271251"
    try:
        inp = input(f"Login ID (default={login_id}): ").strip()
        if inp:
            login_id = inp
    except Exception:
        pass

    # --- Try to load saved password ---
    password = None
    saved_login_file = APP_DIR / "rbsk_saved_login.json"
    if saved_login_file.exists():
        try:
            import json
            import base64
            import ctypes
            import ctypes.wintypes

            class DATA_BLOB(ctypes.Structure):
                _fields_ = [
                    ("cbData", ctypes.wintypes.DWORD),
                    ("pbData", ctypes.POINTER(ctypes.c_char)),
                ]

            with open(saved_login_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            token = saved.get("accounts", {}).get(login_id, "")
            if token:
                encrypted = base64.b64decode(token.encode("ascii"))
                in_buffer = ctypes.create_string_buffer(encrypted)
                in_blob = DATA_BLOB(len(encrypted), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
                out_blob = DATA_BLOB()
                ok = ctypes.windll.crypt32.CryptUnprotectData(
                    ctypes.byref(in_blob), None, None, None, None, 0x01, ctypes.byref(out_blob)
                )
                if ok:
                    try:
                        password = ctypes.string_at(out_blob.pbData, out_blob.cbData).decode("utf-8")
                    finally:
                        ctypes.windll.kernel32.LocalFree(out_blob.pbData)
                    print(f"✓ Using saved password for {login_id}")
        except Exception as e:
            print(f"Could not load saved password: {e}")

    if not password:
        try:
            import getpass
            password = getpass.getpass(f"Enter password for {login_id}: ")
        except Exception:
            password = input(f"Enter password for {login_id}: ").strip()

    # --- Ask for screening month ---
    screening_month = ""
    try:
        inp = input("Screening Month (e.g. 'July 2026' or press Enter to select manually): ").strip()
        if inp:
            screening_month = inp
    except Exception:
        pass

    # --- Ask for institution type ---
    institution_type = "School"
    try:
        inp = input("Institution Type? [1] School  [2] AWC/AWS  (default=1): ").strip()
        if inp == "2" or inp.lower() in {"awc", "aws"}:
            institution_type = "AWC" if inp.lower() != "aws" else "AWS"
    except Exception:
        pass

    # --- Ask for mobile number ---
    custom_mobile = ""
    if Students.is_awc_institution(institution_type):
        print("AWC/AWS mode selected: mobile number will not be asked.")
    else:
        try:
            inp = input("Mobile number for ALL schools (10 digits, or Enter to ask per school): ").strip()
            if inp:
                custom_mobile = inp
        except Exception:
            pass

    # --- Speed mode ---
    speed_mode = "Fast Safe"

    print()
    print("Starting browser...")

    # --- Start browser ---
    driver = Students.start_browser(browser)

    try:
        # Navigate to RBSK portal
        driver.get(Students.APP_URL)
        time.sleep(2)

        # Auto-fill login
        auto_filled = False
        if password:
            try:
                Students.fill_login_fields(driver, login_id, password)
                auto_filled = True
                print(f"✓ Login ID and password auto-filled for {login_id}")
            except Exception:
                print("Could not auto-fill login fields. Please enter manually.")

        # Wait for user to enter CAPTCHA and login
        print()
        print("=" * 60)
        if auto_filled:
            print("  Login fields are auto-filled!")
            print("  1. Enter CAPTCHA manually in the browser")
            print("  2. Click Login in the browser")
            print("  3. Wait for the dashboard to appear")
        else:
            print("  1. Enter Login ID, Password, CAPTCHA in the browser")
            print("  2. Click Login")
            print("  3. Wait for the dashboard to appear")
        print("=" * 60)
        input("  Press ENTER here after you are logged in... ")
        print()

        # Navigate to screening page
        print("Navigating to Screening page...")
        Students.browser_open_screening_page(driver, speed_mode=speed_mode)

        if not screening_month:
            print()
            print("=" * 60)
            print("  On the Screening page:")
            print("  1. Confirm State, District, Block, MHT, Financial Year")
            print("  2. Select the MONTH manually in the browser")
            print("=" * 60)
            input("  Press ENTER here after selecting the month... ")

        # Build args namespace
        args = argparse.Namespace(
            speed_mode=speed_mode,
            screening_month=screening_month,
            screening_page_prepared=True,
            screening_institution_type=institution_type,
            custom_school_mobile=custom_mobile,
            prompt_mobile_callback=None,
            browser=browser,
            delay=0.25,
            resume=False,
            progress="screening_progress.json",
            success_log="screening_success.csv",
            failed_log="screening_failed.csv",
        )

        # Run direct portal screening
        print()
        print("Starting Direct Portal Screening for ALL students...")
        print("=" * 60)
        result = Students.run_direct_portal_screening(driver, args)

        print()
        print("=" * 60)
        print(f"  SCREENING COMPLETE")
        print(f"  Saved: {result.get('saved', 0)}")
        print(f"  Failed: {result.get('failed', 0)}")
        print("=" * 60)

    except KeyboardInterrupt:
        print("\n\nStopped by user.")
    except Exception as exc:
        print(f"\nError: {exc}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            input("\nPress ENTER to close the browser... ")
        except Exception:
            pass
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()

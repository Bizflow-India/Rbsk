"""
RUN_EXCEL_SCREENING.py
======================
Reads student rows from the Excel batch file (sample_screening_health_students.xlsx)
and automatically performs RBSK Screening Health Entry for each pending row.

For each row the tool will:
  1. Navigate to Screening page (you select Month / confirm filters manually once)
  2. Open the school/institute from the Excel "School Name" column
  3. Find the student by name in the portal student list
  4. Fill health parameters:
       - Weight (kg)        — from Excel or auto-calculated by age
       - Height/Length (cm) — from Excel or auto-calculated by age
       - Blood Pressure     — from Excel (default: 100/70, range 60–100 diastolic)
       - Vision Left Eye    — from Excel (default: 6/6)
       - Vision Right Eye   — from Excel (default: 6/6)
  5. Click Preview & Submit
  6. Fill the mobile number from the Excel row in the Preview modal
  7. Click Confirm & Submit
  8. Proceed to the next pending row automatically

Excel file expected: sample_screening_health_students.xlsx
  Required columns:  School Name, Student Full Name, Mobile Number
  Optional columns:  Weight, Height, Blood Pressure, Vision Left, Vision Right, HB Count

Usage:
  Double-click this script  (or: python RUN_EXCEL_SCREENING.py)
  1. Browser opens -> Enter CAPTCHA manually -> Click Login
  2. Press ENTER after you are logged in
  3. The tool navigates to the Screening page -- select Month & click Search
  4. Press ENTER to start the automated entry loop
"""

import argparse
import sys
import time
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import Students

EXCEL_FILE = APP_DIR / "sample_screening_health_students.xlsx"
SHEET_NAME = "Students"          # first sheet if not found

PROGRESS_FILE   = APP_DIR / "excel_screening_progress.json"
SUCCESS_LOG     = APP_DIR / "excel_screening_success.csv"
FAILED_LOG      = APP_DIR / "excel_screening_failed.csv"


# ---------------------------------------------------------------------------
# Column mapping helpers
# ---------------------------------------------------------------------------

def normalise_row(row: dict) -> dict:
    """
    The Excel uses 'School Name' for the institute.
    Students.py looks for 'Institute Name' in screening functions,
    so we add an alias so the existing functions work without modification.
    """
    if "School Name" in row and "Institute Name" not in row:
        row["Institute Name"] = row["School Name"]
    # Also map 'Vision Left' -> 'Vision - Left Eye' alias used by screening_health_payload
    if "Vision Left" in row and "Vision - Left Eye" not in row:
        row["Vision - Left Eye"] = row["Vision Left"]
    if "Vision Right" in row and "Vision - Right Eye" not in row:
        row["Vision - Right Eye"] = row["Vision Right"]
    return row


def load_excel_rows() -> list:
    """Load and normalise rows from the Excel batch file."""
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}\n"
            "Place the file in the same folder as this script."
        )
    try:
        rows = Students.load_xlsx_rows(EXCEL_FILE, SHEET_NAME)
    except Exception:
        # Try first sheet if named sheet not found
        rows = Students.load_xlsx_rows(EXCEL_FILE, None)

    rows = Students.normalize_student_rows(rows)
    rows = [normalise_row(r) for r in rows]
    return rows


# ---------------------------------------------------------------------------
# Login helpers
# ---------------------------------------------------------------------------

def load_saved_password(login_id: str):
    import json
    import base64
    import ctypes
    import ctypes.wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [
            ("cbData", ctypes.wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_char)),
        ]

    saved_login_file = APP_DIR / "rbsk_saved_login.json"
    if not saved_login_file.exists():
        return None
    try:
        with open(saved_login_file, "r", encoding="utf-8") as f:
            saved = json.load(f)
        token = saved.get("accounts", {}).get(login_id, "")
        if not token:
            return None
        encrypted = base64.b64decode(token.encode("ascii"))
        in_buffer = ctypes.create_string_buffer(encrypted)
        in_blob = DATA_BLOB(len(encrypted), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
        out_blob = DATA_BLOB()
        ok = ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(in_blob), None, None, None, None, 0x01, ctypes.byref(out_blob)
        )
        if ok:
            try:
                return ctypes.string_at(out_blob.pbData, out_blob.cbData).decode("utf-8")
            finally:
                ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    print("=" * 65)
    print("   RBSK Excel Screening -- Automated Health Entry")
    print("=" * 65)
    print(f"   Excel file: {EXCEL_FILE.name}")
    print()

    # --- Load Excel rows ---
    try:
        rows = load_excel_rows()
    except Exception as exc:
        print(f"ERROR loading Excel: {exc}")
        input("Press ENTER to exit...")
        return

    if not rows:
        print("No student rows found in the Excel file.")
        input("Press ENTER to exit...")
        return

    # Validate rows have the minimum columns needed
    try:
        Students.validate_screening_health_rows(rows)
    except Students.ImportErrorWithHelp as exc:
        print(f"Validation ERROR: {exc}")
        input("Press ENTER to exit...")
        return

    print(f"Loaded {len(rows)} student row(s) from Excel.\n")
    for i, row in enumerate(rows[:5], 1):
        name    = Students.screening_student_name(row)
        school  = Students.screening_institute_name(row)
        mobile  = Students.screening_mobile(row)
        health  = Students.screening_health_payload(row)
        print(
            f"  {i}. {name} | {school} | "
            f"Wt={health.get('weight')} Ht={health.get('height')} "
            f"BP={health.get('bloodPressure')} "
            f"VL={health.get('visionLeft')} VR={health.get('visionRight')} "
            f"Mobile={mobile}"
        )
    if len(rows) > 5:
        print(f"  ... and {len(rows) - 5} more row(s).")
    print()

    # --- Browser ---
    browser = "chrome"
    try:
        choice = input("Browser? [1] Chrome  [2] Edge  (default=1): ").strip()
        if choice == "2":
            browser = "edge"
    except Exception:
        pass

    # --- Login ID ---
    login_id = "MHT-1271251"
    try:
        inp = input(f"Login ID (default={login_id}): ").strip()
        if inp:
            login_id = inp
    except Exception:
        pass

    # --- Password ---
    password = load_saved_password(login_id)
    if password:
        print(f"Saved password found for {login_id}.")
    else:
        try:
            import getpass
            password = getpass.getpass(f"Enter password for {login_id}: ")
        except Exception:
            password = input(f"Enter password for {login_id}: ").strip()

    # --- Screening month ---
    screening_month = ""
    try:
        inp = input(
            "Screening Month (e.g. 'July 2026', or press Enter to select manually in browser): "
        ).strip()
        if inp:
            screening_month = inp
    except Exception:
        pass

    # --- Custom mobile (global override) ---
    custom_mobile = ""
    try:
        inp = input(
            "Global Mobile Number for ALL schools (10 digits, or Enter to use per-row mobile from Excel): "
        ).strip()
        if inp:
            custom_mobile = inp
    except Exception:
        pass

    # --- Speed ---
    speed_mode = "Fast Safe"

    # --- Resume ---
    resume = False
    try:
        inp = input("Resume from last saved progress? [y/N]: ").strip().lower()
        resume = inp in ("y", "yes")
    except Exception:
        pass

    print()
    print("Starting browser...")
    driver = Students.start_browser(browser)

    try:
        driver.get(Students.APP_URL)
        time.sleep(2)

        # Auto-fill login
        if password:
            try:
                Students.fill_login_fields(driver, login_id, password)
                print(f"Login fields auto-filled for {login_id}.")
            except Exception:
                print("Could not auto-fill login fields. Please fill them manually.")

        print()
        print("=" * 65)
        print("  LOGIN STEP:")
        print("  1. Enter CAPTCHA manually in the browser")
        print("  2. Click Login")
        print("  3. Wait until the RBSK dashboard is visible")
        print("=" * 65)
        input("  >> Press ENTER here after you are logged in... ")
        print()

        # Navigate to Screening page
        print("Navigating to Screening page...")
        Students.browser_open_screening_page(driver, speed_mode=speed_mode)

        if not screening_month:
            print()
            print("=" * 65)
            print("  SCREENING SETUP:")
            print("  1. Confirm State / District / Block / MHT")
            print("  2. Select the MONTH for this screening cycle")
            print("  3. (Optional) click Search to see the school list")
            print("=" * 65)
            input("  >> Press ENTER here after selecting the Month... ")

        # Build args for process_screening_health_rows_browser()
        # mobile_source controls whether the mobile is taken from each Excel row
        # ('excel') or asked once per school ('portal_school').
        mobile_source = "portal_school" if custom_mobile else "excel"

        args = argparse.Namespace(
            speed_mode=speed_mode,
            screening_month=screening_month,
            screening_page_prepared=True,
            mobile_source=mobile_source,
            custom_school_mobile=custom_mobile,
            prompt_mobile_callback=None,
            browser=browser,
            delay=0.25,
            resume=resume,
            progress=str(PROGRESS_FILE),
            success_log=str(SUCCESS_LOG),
            failed_log=str(FAILED_LOG),
        )

        print()
        print("=" * 65)
        print(f"  Starting automated Screening Health Entry for {len(rows)} row(s)...")
        print("=" * 65)
        print()

        result = Students.process_screening_health_rows_browser(
            driver,
            rows,
            args,
            label="Excel Screening Health Entry",
        )

        print()
        print("=" * 65)
        print("  SCREENING COMPLETE")
        print(f"  Saved   : {result.get('saved', 0)}")
        print(f"  Failed  : {result.get('failed', 0)}")
        print(f"  Skipped : {result.get('skipped', 0)}")
        print(f"  Success log : {SUCCESS_LOG.name}")
        print(f"  Failed  log : {FAILED_LOG.name}")
        print("=" * 65)

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

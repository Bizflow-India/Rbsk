from __future__ import annotations

import argparse
import csv
import difflib
import getpass
import json
import re
import sys
import time
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import requests
except ModuleNotFoundError:
    requests = None

try:
    import urllib3
except ModuleNotFoundError:
    urllib3 = None


BASE_URL = "https://rbsk.mohfw.gov.in"
LOGIN_URL = f"{BASE_URL}/RBSK/login"
APP_URL = LOGIN_URL
STUDENTMASTER_URL = f"{BASE_URL}/RBSK/studentmaster"
API_BASE = f"{BASE_URL}/rbsk_webapi/api"
CREATE_STUDENT_ENDPOINT = f"{API_BASE}/Student/Create_Update_Students"
HIERARCHY_ENDPOINT = f"{API_BASE}/Hierarchy_Master/Get_HierarchyMaster"
INSTITUTE_MAPPING_CACHE: dict[str, dict[str, str]] = {}

REQUIRED_COLUMNS = [
    "Sr No",
    "Student Full Name",
    "Birth Date",
    "Gender",
    "Mother Name",
    "Father Name",
    "Mobile Number",
    "Mobile Owner",
    "Full Address",
    "Pin Code",
    "Blood Group",
    "Class",
    "Section",
]

BASE_REQUIRED_COLUMNS = [
    "Sr No",
    "Student Full Name",
    "Birth Date",
    "Gender",
    "Mother Name",
    "Father Name",
    "Mobile Number",
    "Mobile Owner",
    "Full Address",
]

COLUMN_ALIASES = {
    "Child Name": "Student Full Name",
    "Student Name": "Student Full Name",
    "Name": "Student Full Name",
    "Birth Date (DD-MM-YYYY)": "Birth Date",
    "DOB": "Birth Date",
    "Mother's Name": "Mother Name",
    "Fathers Name": "Father Name",
    "Father's Name": "Father Name",
    "Whose Mobile": "Mobile Owner",
    "Address": "Full Address",
    "Institute": "Institute Name",
    "Institutes": "Institute Name",
    "Institute Name (AWC NAME)": "Institute Name",
    "AWC NAME": "Institute Name",
    "Center": "Institute Name",
    "Center Name": "Institute Name",
    "AWC Name": "Institute Name",
    "Screening Location": "Institute Name",
}

OPTIONAL_DEFAULT_COLUMNS = {
    "Pin Code": "",
    "Blood Group": "",
    "Class": "0",
    "Section": "0",
    "Birth Weight (grams)": "",
}

BATCH_REQUIRED_COLUMNS = [
    "Batch Name",
    "Login ID",
    "Excel File",
]

BATCH_OPTIONAL_COLUMNS = [
    "Sheet",
    "State Code",
    "District Code",
    "Health Block Code",
    "Institution Type",
    "Institution Code",
]

DEMO_FORM_FILE = "test_form.html"
DEMO_FIELDS = {
    "student_full_name": "Student Full Name",
    "birth_date": "Birth Date",
    "gender": "Gender",
    "mother_name": "Mother Name",
    "father_name": "Father Name",
    "mobile_number": "Mobile Number",
    "mobile_owner": "Mobile Owner",
    "full_address": "Full Address",
    "pin_code": "Pin Code",
    "blood_group": "Blood Group",
    "class_name": "Class",
    "section": "Section",
}

GENDER_MAP = {
    "m": "M",
    "male": "M",
    "f": "F",
    "female": "F",
    "t": "T",
    "transgender": "T",
    "other": "T",
}

MOBILE_OWNER_MAP = {
    "1": 1,
    "parent": 1,
    "parents": 1,
    "2": 2,
    "guardian": 2,
    "3": 3,
    "other": 3,
    "others": 3,
}

BLOOD_GROUP_MAP = {
    "": "",
    "0": "",
    "a+": "1",
    "a-": "2",
    "b+": "3",
    "b-": "4",
    "ab+": "5",
    "ab-": "6",
    "o+": "7",
    "o-": "8",
    "1": "1",
    "2": "2",
    "3": "3",
    "4": "4",
    "5": "5",
    "6": "6",
    "7": "7",
    "8": "8",
}


class ImportErrorWithHelp(RuntimeError):
    pass


class DuplicateRecordFound(RuntimeError):
    pass


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"none", "nan", "nat"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def is_awc_institution(value: Any) -> bool:
    text = clean(value).upper()
    return text in {"AWC", "AWS", "1"}


def safe_input(prompt: str = "") -> str:
    if prompt:
        print(prompt)
    if sys.stdin is None or not hasattr(sys.stdin, "readline"):
        time.sleep(1)
        return ""
    try:
        return input().strip()
    except (EOFError, RuntimeError, OSError, ValueError):
        time.sleep(1)
        return ""



def emit_progress(args: argparse.Namespace, event_type: str, **data: Any) -> None:
    callback = getattr(args, "progress_callback", None)
    if not callback:
        return
    try:
        callback({"type": event_type, **data})
    except Exception:
        pass


def load_xlsx_rows(path: Path, sheet_name: str | None = None) -> list[dict[str, Any]]:
    """Small .xlsx reader for plain workbooks, avoiding pandas/openpyxl dependency."""
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

    with zipfile.ZipFile(path) as zf:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", ns):
                shared_strings.append("".join(t.text or "" for t in item.findall(".//a:t", ns)))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("r:Relationship", rel_ns)}

        selected_sheet = None
        for sheet in workbook.findall("a:sheets/a:sheet", ns):
            name = sheet.attrib["name"]
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            if sheet_name is None or name == sheet_name:
                target = rel_targets[rid].lstrip("/")
                selected_sheet = target if target.startswith("xl/") else f"xl/{target}"
                break

        if not selected_sheet:
            raise ImportErrorWithHelp(f"Sheet not found: {sheet_name}")

        root = ET.fromstring(zf.read(selected_sheet))
        rows: list[list[str]] = []
        for row in root.findall("a:sheetData/a:row", ns):
            values_by_col: dict[int, str] = {}
            for cell in row.findall("a:c", ns):
                ref = cell.attrib.get("r", "")
                col_letters = "".join(ch for ch in ref if ch.isalpha())
                col_index = letters_to_index(col_letters)
                cell_type = cell.attrib.get("t")
                value_el = cell.find("a:v", ns)
                inline_el = cell.find("a:is/a:t", ns)
                value = ""
                if inline_el is not None:
                    value = inline_el.text or ""
                elif value_el is not None:
                    value = value_el.text or ""
                if cell_type == "s" and value:
                    value = shared_strings[int(value)]
                values_by_col[col_index] = value
            if values_by_col:
                max_col = max(values_by_col)
                rows.append([values_by_col.get(i, "") for i in range(1, max_col + 1)])

    if not rows:
        return []
    headers = [clean(v) for v in rows[0]]
    result: list[dict[str, Any]] = []
    for row in rows[1:]:
        item = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers)) if headers[i]}
        if any(clean(v) for v in item.values()):
            result.append(item)
    return result


def letters_to_index(letters: str) -> int:
    value = 0
    for ch in letters.upper():
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value


def load_csv_rows(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def load_rows(path: Path, sheet_name: str | None) -> list[dict[str, Any]]:
    if not path.exists():
        raise ImportErrorWithHelp(f"Excel/CSV file not found: {path}")
    if path.suffix.lower() == ".csv":
        return load_csv_rows(path)
    if path.suffix.lower() == ".xlsx":
        return load_xlsx_rows(path, sheet_name)
    raise ImportErrorWithHelp("Use a .xlsx or .csv file.")


def normalize_student_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized_rows: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        normalized = dict(row)
        for source, target in COLUMN_ALIASES.items():
            if target not in normalized and source in normalized:
                normalized[target] = normalized.get(source)
        if clean(normalized.get("Institute Name")):
            normalized["Institute Name"] = re.sub(r"\s*\(\d+\)\s*$", "", clean(normalized.get("Institute Name")))
        if "Sr No" not in normalized or not clean(normalized.get("Sr No")):
            normalized["Sr No"] = str(index)
        for column, default in OPTIONAL_DEFAULT_COLUMNS.items():
            normalized.setdefault(column, default)
        normalized_rows.append(normalized)
    return normalized_rows


def excel_date_to_iso(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    if text.replace(".", "", 1).isdigit():
        # Excel's Windows date system uses 1899-12-30 as the practical epoch.
        return (datetime(1899, 12, 30) + timedelta(days=float(text))).strftime("%Y-%m-%d")
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Unsupported date format: {text}")


def excel_date_to_ddmmyyyy(value: Any) -> str:
    iso = excel_date_to_iso(value)
    if not iso:
        return ""
    return datetime.strptime(iso, "%Y-%m-%d").strftime("%d-%m-%Y")


def age_years_months_from_dob(value: Any) -> str:
    iso = excel_date_to_iso(value)
    if not iso:
        return ""
    dob = datetime.strptime(iso, "%Y-%m-%d")
    today = datetime.today()
    years = today.year - dob.year
    months = today.month - dob.month
    if today.day < dob.day:
        months -= 1
    if months < 0:
        years -= 1
        months += 12
    return f"{max(years, 0)}.{max(months, 0)}" if months else str(max(years, 0))


def select_rows(rows: list[dict[str, Any]], only_sr: str | None, limit: int | None) -> list[dict[str, Any]]:
    selected = rows
    if only_sr:
        wanted = parse_sr_numbers(only_sr)
        selected = [row for row in rows if clean(row.get("Sr No")) in wanted]
    if limit:
        selected = selected[:limit]
    return selected


def parse_sr_numbers(text: str) -> set[str]:
    selected: set[str] = set()
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = [int(x.strip()) for x in part.split("-", 1)]
            selected.update(str(i) for i in range(start, end + 1))
        else:
            selected.add(clean(part))
    return selected


def validate_columns(rows: list[dict[str, Any]]) -> None:
    if not rows:
        raise ImportErrorWithHelp("No rows found in the selected sheet/file.")
    missing = [col for col in BASE_REQUIRED_COLUMNS if col not in rows[0]]
    if missing:
        raise ImportErrorWithHelp(f"Missing required column(s): {', '.join(missing)}")


def require_selenium():
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.edge.options import Options as EdgeOptions
    except ModuleNotFoundError as exc:
        raise ImportErrorWithHelp(
            "Selenium is required for manual login/session capture. Install it with:\n"
            "  python -m pip install selenium"
        ) from exc
    return webdriver, ChromeOptions, EdgeOptions


def start_browser(browser_name: str, profile_dir: str | Path | None = None):
    webdriver, ChromeOptions, EdgeOptions = require_selenium()
    if browser_name == "edge":
        options = EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-dev-shm-usage")
        if profile_dir:
            Path(profile_dir).mkdir(parents=True, exist_ok=True)
            options.add_argument(f"--user-data-dir={profile_dir}")
        return webdriver.Edge(options=options)
    options = ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--remote-debugging-port=0")
    if profile_dir:
        Path(profile_dir).mkdir(parents=True, exist_ok=True)
        options.add_argument(f"--user-data-dir={profile_dir}")
    return webdriver.Chrome(options=options)


_SHARED_BROWSER_DRIVER = None


def browser_ensure_session_or_create(browser_type: str = "chrome", profile_dir: str | Path | None = None, allow_insecure_ssl: bool = False):
    global _SHARED_BROWSER_DRIVER
    if _SHARED_BROWSER_DRIVER is not None:
        try:
            _SHARED_BROWSER_DRIVER.current_url
            return _SHARED_BROWSER_DRIVER
        except Exception:
            _SHARED_BROWSER_DRIVER = None

    driver = start_browser(browser_type, profile_dir)
    _SHARED_BROWSER_DRIVER = driver
    return driver


def ask_school_mobile_console(school_name: str) -> str:
    print(f"\nEnter 10-digit Mobile Number for school '{school_name}': ")
    try:
        val = safe_input().strip()
        return val
    except Exception:
        return "9999999999"



def fill_login_fields(driver, login_id, password):
    script = """
    const username = arguments[0];
    const password = arguments[1];

    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    };
    const setValue = (el, value) => {
      if (!el) return false;
      el.focus();
      const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
      if (setter) setter.call(el, value);
      else el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    };
    const inputs = Array.from(document.querySelectorAll('input')).filter(visible);
    const byHint = (hints, type) => inputs.find((el) => {
      const hay = `${el.name || ''} ${el.id || ''} ${el.placeholder || ''} ${el.autocomplete || ''} ${el.type || ''}`.toLowerCase();
      return (!type || el.type === type) && hints.some((hint) => hay.includes(hint));
    });

    const userEl = byHint(['user', 'login', 'mht', 'id', 'name'], null) || inputs.find((el) => el.type !== 'password');
    const passEl = byHint(['pass'], 'password') || inputs.find((el) => el.type === 'password');
    return { usernameFilled: setValue(userEl, username), passwordFilled: setValue(passEl, password) };
    """
    deadline = time.time() + 10.0
    result: dict[str, Any] = {}
    while time.time() < deadline:
        try:
            result = driver.execute_script(script, login_id, password) or {}
            if result.get("usernameFilled") and result.get("passwordFilled"):
                print("RBSK login page opened and Login ID/password were filled automatically.")
                return result
        except Exception:
            pass
        browser_sleep(0.2)
    raise ImportErrorWithHelp(
        f"RBSK login fields did not become ready at {LOGIN_URL}. Refresh the login page and try again."
    )


def capture_logged_in_driver(args: argparse.Namespace, login_id: str | None = None, password: str | None = None):
    driver = start_browser(args.browser, getattr(args, "browser_profile_dir", None))
    driver.get(APP_URL)
    if login_id and password and not getattr(args, "manual_login", False):
        fill_login_fields(driver, login_id, password)
    return driver


def capture_logged_in_session(args: argparse.Namespace, login_id: str | None = None, password: str | None = None) -> tuple[requests.Session, dict[str, Any]]:
    driver = start_browser(args.browser, getattr(args, "browser_profile_dir", None))
    try:
        driver.get(APP_URL)
        if login_id and password and not getattr(args, "manual_login", False):
            fill_login_fields(driver, login_id, password)

        print("\nLogin step:")
        if getattr(args, "manual_login", False):
            print("1. Type Login ID, password, and CAPTCHA manually in the browser.")
        elif login_id:
            print(f"1. Login ID has been filled for {login_id} if the portal fields were detectable.")
        else:
            print("1. Complete login in the opened browser.")
        print("2. Enter CAPTCHA manually and submit the portal login.")
        safe_input("Press Enter here only after you can see the logged-in RBSK portal...")

        driver.get(STUDENTMASTER_URL)
        print("\nStudent Master context step:")
        print("1. If the page shows filters, select the correct state/district/block/institution.")
        print("2. If the page already shows the correct institution context, continue.")
        safe_input("Press Enter here only after the correct Student Master context is visible...")
        time.sleep(1)
        return collect_browser_session(driver)
    finally:
        if not args.keep_browser_open:
            driver.quit()


def collect_browser_session(driver) -> tuple[Any, dict[str, Any]]:
    storage = driver.execute_script(
        """
        const dump = (store) => {
          const out = {};
          for (let i = 0; i < store.length; i++) {
            const key = store.key(i);
            out[key] = store.getItem(key);
          }
          return out;
        };
        return {localStorage: dump(localStorage), sessionStorage: dump(sessionStorage), url: location.href};
        """
    )

    if requests is None:
        return None, storage

    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"], domain=cookie.get("domain"), path=cookie.get("path", "/"))

    xsrf = session.cookies.get("XSRF-TOKEN")
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "Origin": BASE_URL,
            "Referer": STUDENTMASTER_URL,
            "User-Agent": driver.execute_script("return navigator.userAgent;"),
            "Cache-Control": "no-store",
            "Pragma": "no-cache",
        }
    )
    if xsrf:
        session.headers["X-XSRF-TOKEN"] = xsrf

    return session, storage


def apply_ssl_setting(session: Any, args: argparse.Namespace) -> Any:
    if session is None:
        return session
    if getattr(args, "allow_insecure_ssl", False) and urllib3 is not None:
        session.verify = False
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    return session


def new_requests_session_or_raise() -> requests.Session:
    if requests is None:
        raise ImportErrorWithHelp(
            "Python package dependency is incomplete: requests/urllib3 is not available.\n"
            "Live browser form mode can still run, but dry-run/API checks need urllib3 installed."
        )
    return requests.Session()


def parse_storage(storage: dict[str, Any]) -> dict[str, Any]:
    parsed: dict[str, Any] = {}
    for area in ("localStorage", "sessionStorage"):
        for key, value in (storage.get(area) or {}).items():
            parsed[f"{area}.{key}"] = parse_json_maybe(value)
    return parsed


def parse_json_maybe(value: str) -> Any:
    try:
        return json.loads(value)
    except Exception:
        return value


def walk_values(obj: Any):
    if isinstance(obj, dict):
        yield obj
        for value in obj.values():
            yield from walk_values(value)
    elif isinstance(obj, list):
        for item in obj:
            yield from walk_values(item)


def first_value(parsed_storage: dict[str, Any], keys: list[str]) -> Any:
    normalized = {key.lower(): key for key in keys}
    for root in parsed_storage.values():
        for item in walk_values(root):
            for key, value in item.items():
                if key.lower() in normalized and clean(value):
                    return value
    return None


def first_numeric_value(parsed_storage: dict[str, Any], keys: list[str]) -> Any:
    normalized = {key.lower(): key for key in keys}
    for root in parsed_storage.values():
        for item in walk_values(root):
            for key, value in item.items():
                if key.lower() in normalized and int_or_none(value) is not None:
                    return value
    return None


def first_digits_value(parsed_storage: dict[str, Any], keys: list[str]) -> Any:
    normalized = {key.lower(): key for key in keys}
    for root in parsed_storage.values():
        for item in walk_values(root):
            for key, value in item.items():
                if key.lower() in normalized:
                    digits = only_digits(value, 20)
                    if digits:
                        return digits
    return None


def int_or_none(value: Any) -> int | None:
    text = clean(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def infer_context(parsed_storage: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    context = {
        "state_code": args.state_code or first_value(parsed_storage, ["state_ID", "stateId", "State_Code", "stateCode", "statename"]),
        "district_code": args.district_code
        or first_value(parsed_storage, ["district_ID", "districtId", "District_Code", "districtCode", "districtname"]),
        "health_block_code": args.health_block_code
        or first_value(parsed_storage, ["healthBlock_ID", "healthBlockId", "HealthBlock_Code", "blockCode", "blockname"]),
        "mht_code": first_value(parsed_storage, ["mhT_Code", "MHT_Code", "mhtCode", "mhtid", "mht_ID"]),
        "institution_type": args.institution_type or first_value(parsed_storage, ["Institution_Type", "institutionType", "institutiontype"]),
        "institution_code": args.institution_code or first_value(parsed_storage, ["Institution_Code", "institutionCode"]),
        "created_By": args.created_by
        or first_numeric_value(parsed_storage, ["mhT_Code", "MHT_Code", "mhtCode", "mhtid", "mht_ID"])
        or first_digits_value(parsed_storage, ["loginID", "loginId"])
        or first_numeric_value(parsed_storage, ["userID", "userId", "User_ID", "user_ID", "created_By", "createdBy", "id", "ID"]),
    }

    if not context["institution_type"]:
        context["institution_type"] = 1
    return {key: clean(value) for key, value in context.items()}


def ensure_context(context: dict[str, Any], require_created_by: bool = True, require_institution_code: bool = True) -> None:
    required = ["state_code", "district_code", "health_block_code"]
    if require_institution_code:
        required.append("institution_code")
    if require_created_by:
        required.append("created_By")
    missing = [key for key in required if not clean(context.get(key))]
    if missing:
        raise ImportErrorWithHelp(
            "Could not infer required portal context: "
            + ", ".join(missing)
            + "\nOpen the student master page after login and/or pass overrides such as:\n"
            + "  --state-code 27 --district-code 0 --health-block-code 0 --institution-code 123 --created-by 123"
        )


def fetch_option_map(session: requests.Session, master_type: int, state_code: str) -> dict[str, str]:
    payload = {
        "masterType": master_type,
        "stateID": int(state_code or 0),
        "districtID": 0,
        "healthBlockID": 0,
        "mhtid": 0,
        "deicCode": 0,
        "finYear": "",
        "screenCycleID": 0,
        "finMonthID": 0,
        "institutionType": 0,
        "facilityType": 0,
    }
    try:
        data = session.post(HIERARCHY_ENDPOINT, json=payload, timeout=60).json()
    except Exception:
        return {}
    items = data.get("data") if isinstance(data, dict) else []
    if not isinstance(items, list):
        return {}
    option_map: dict[str, str] = {}
    for item in items:
        code = item.get("Class_Code") or item.get("class_code") or item.get("Section_Code") or item.get("section_code") or item.get("code")
        label = item.get("Class_Name") or item.get("class_name") or item.get("Section_Name") or item.get("section_name") or item.get("name")
        if code is not None:
            option_map[clean(code).lower()] = clean(code)
        if label is not None and code is not None:
            option_map[clean(label).lower()] = clean(code)
    return option_map


def fetch_institution_map(session: requests.Session, context: dict[str, Any]) -> dict[str, str]:
    option_map: dict[str, str] = {}
    for item in fetch_institution_options(session, context):
        code = clean(item.get("code"))
        label = clean(item.get("label"))
        if code:
            option_map[code.lower()] = code
        if label and code:
            option_map[normalize_match_text(label)] = code
    return option_map


def fetch_institution_options(session: requests.Session, context: dict[str, Any]) -> list[dict[str, str]]:
    payload = {
        "masterType": 9,
        "stateID": int(float(context.get("state_code") or 0)),
        "districtID": int(float(context.get("district_code") or 0)),
        "healthBlockID": int(float(context.get("health_block_code") or 0)),
        "mhtid": int_or_zero(context.get("mht_code")),
        "deicCode": 0,
        "finYear": "",
        "screenCycleID": 0,
        "finMonthID": 0,
        "institutionType": int(float(context.get("institution_type") or 1)),
        "facilityType": 0,
    }
    try:
        data = session.post(HIERARCHY_ENDPOINT, json=payload, timeout=60).json()
    except Exception:
        return {}
    items = data.get("data") if isinstance(data, dict) else []
    if not isinstance(items, list):
        return []
    options: list[dict[str, str]] = []
    for item in items:
        code = (
            item.get("Institution_Code")
            or item.get("institution_code")
            or item.get("Institute_Code")
            or item.get("code")
            or item.get("value")
        )
        label = (
            item.get("Institution_Name")
            or item.get("institution_name")
            or item.get("Institute_Name")
            or item.get("name")
            or item.get("label")
        )
        label_text = clean(label)
        if label_text:
            options.append({"label": label_text, "code": clean(code)})
    return options


def normalize_match_text(value: Any) -> str:
    return " ".join(re.sub(r"[()\[\]{}.,_-]+", " ", clean(value).lower().replace("...", "")).split())


def resolve_row_institution_code(row: dict[str, Any], context: dict[str, Any], institution_map: dict[str, str]) -> int:
    explicit_code = clean(row.get("Institution Code") or row.get("Institute Code"))
    if explicit_code:
        return int(float(explicit_code))
    institute_name = clean(row.get("Institute Name"))
    if institute_name:
        exact = institution_map.get(normalize_match_text(institute_name))
        if exact:
            return int(float(exact))
        matches = [
            code
            for label, code in institution_map.items()
            if normalize_match_text(institute_name) in label or label in normalize_match_text(institute_name)
        ]
        if len(matches) == 1:
            return int(float(matches[0]))
        if not matches:
            raise ValueError(f"Institute Name not found in RBSK list: {institute_name}")
        raise ValueError(f"Institute Name matched multiple RBSK entries: {institute_name}")
    return int(float(context["institution_code"]))


def map_class_section(value: Any, options: dict[str, str], fallback_letters: bool = False) -> int:
    text = clean(value)
    if not text:
        return 0
    mapped = options.get(text.lower())
    if mapped:
        return int(float(mapped))
    if text.isdigit():
        return int(text)
    if fallback_letters and len(text) == 1 and text.isalpha():
        return ord(text.upper()) - ord("A") + 1
    raise ValueError(f"Could not map value '{text}'")


def build_payload(
    row: dict[str, Any],
    context: dict[str, Any],
    class_map: dict[str, str],
    section_map: dict[str, str],
    institution_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    gender = GENDER_MAP.get(clean(row.get("Gender")).lower(), clean(row.get("Gender")).upper())
    mobile_owner = MOBILE_OWNER_MAP.get(clean(row.get("Mobile Owner")).lower())
    if mobile_owner is None:
        raise ValueError(f"Unknown Mobile Owner: {row.get('Mobile Owner')}")

    blood_group = BLOOD_GROUP_MAP.get(clean(row.get("Blood Group")).lower())
    if blood_group is None:
        raise ValueError(f"Unknown Blood Group: {row.get('Blood Group')}")

    institution_type = int(float(context["institution_type"] or 1))
    institution_code = resolve_row_institution_code(row, context, institution_map or {})
    student_class = 0 if institution_type == 1 else map_class_section(row.get("Class"), class_map)
    student_section = 0 if institution_type == 1 else map_class_section(row.get("Section"), section_map, fallback_letters=True)

    return {
        "action": "Insert",
        "student_code": 0,
        "institution_type": institution_type,
        "institution_code": institution_code,
        "state_code": int(float(context["state_code"])),
        "district_code": int(float(context["district_code"])),
        "health_block_code": int(float(context["health_block_code"])),
        "student_name": clean(row.get("Student Full Name")),
        "student_dob": excel_date_to_iso(row.get("Birth Date")),
        "student_gender": gender,
        "student_father": clean(row.get("Father Name")),
        "student_mother": clean(row.get("Mother Name")),
        "mobile": only_digits(row.get("Mobile Number"), 10),
        "whose_mobile": mobile_owner,
        "student_address": clean(row.get("Full Address")),
        "pincode": only_digits(row.get("Pin Code"), 6),
        "student_blood_group": blood_group,
        "student_class": student_class,
        "student_section": student_section,
        "birth_place": 0,
        "birth_weight": int_or_zero(row.get("Birth Weight (grams)") or row.get("Birth Weight")),
        "apaar_id": clean(row.get("APAAR ID") or row.get("PEN") or row.get("PEN AWC Number")),
        "abha_detail": None,
        "student_roll_number": clean(row.get("Roll Number")),
        "student_remarks": clean(row.get("Remarks")),
        "birth_place_name": "",
        "ip_address": "NA",
        "created_By": int_or_zero(context.get("created_By")),
    }


def only_digits(value: Any, max_len: int) -> str:
    return "".join(ch for ch in clean(value) if ch.isdigit())[:max_len]


def int_or_zero(value: Any) -> int:
    text = clean(value)
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def validate_payload(payload: dict[str, Any]) -> list[str]:
    errors = []
    for key in ["student_name", "student_dob", "student_gender", "student_father", "student_mother", "mobile", "student_address"]:
        if not clean(payload.get(key)):
            errors.append(f"{key} is required")
    if payload["student_gender"] not in {"M", "F", "T"}:
        errors.append(f"student_gender must be M/F/T, got {payload['student_gender']}")
    if payload["mobile"] and len(payload["mobile"]) != 10:
        errors.append("mobile must be 10 digits")
    if payload["pincode"] and len(payload["pincode"]) != 6:
        errors.append("pincode must be 6 digits")
    if payload.get("institution_type") != 1 and not payload["student_class"]:
        errors.append("student_class could not be mapped")
    return errors


def append_csv(path: Path, row: dict[str, Any], headers: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    exists = path.exists()
    with path.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
        if not exists:
            writer.writeheader()
        writer.writerow(row)


def load_progress(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return set(str(x) for x in data.get("completed_sr_numbers", []))
    except Exception:
        return set()


def save_progress(path: Path, completed: set[str]) -> None:
    path.write_text(json.dumps({"completed_sr_numbers": sorted(completed, key=lambda x: int(x) if x.isdigit() else x)}, indent=2), encoding="utf-8")


def enforce_entry_license(args: argparse.Namespace) -> None:
    callback = getattr(args, "license_check_callback", None)
    if callback:
        callback()


def record_licensed_success(args: argparse.Namespace, count: int = 1) -> None:
    callback = getattr(args, "license_success_callback", None)
    if callback:
        callback(count)


def submit_student(session: requests.Session, payload: dict[str, Any], timeout: int) -> tuple[bool, str, dict[str, Any]]:
    response = session.post(CREATE_STUDENT_ENDPOINT, json=payload, timeout=timeout)
    try:
        data = response.json()
    except Exception:
        data = {"raw": response.text[:1000]}
    if response.status_code in {401, 419}:
        return False, "Session expired or unauthorized. Log in again and resume.", data
    if response.status_code >= 400:
        return False, f"HTTP {response.status_code}", data
    message = clean(data.get("message") if isinstance(data, dict) else "")
    if any(term in message.lower() for term in ["invalid", "failed", "error", "not found"]):
        return False, message or "Portal returned a failure message", data
    if isinstance(data, dict) and data.get("succecced") and data.get("status_Code"):
        return True, message or "Saved", data
    return False, message or "Portal did not return success", data


def browser_sleep(seconds: float = 0.4) -> None:
    time.sleep(seconds)


def speed_is_fast_safe(speed_mode: Any = "") -> bool:
    return clean(speed_mode).lower().replace("-", " ").startswith("fast")


def speed_wait(speed_mode: Any, fast: float, conservative: float) -> float:
    return fast if speed_is_fast_safe(speed_mode) else conservative


def speed_timeout(speed_mode: Any, fast: float, conservative: float) -> float:
    return fast if speed_is_fast_safe(speed_mode) else conservative


def js_click_text(driver, texts: list[str]) -> bool:
    script = """
    const texts = arguments[0].map(t => String(t).toLowerCase());
    const nodes = Array.from(document.querySelectorAll('button, [role="button"], a, div, span'));
    const visible = el => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
    };
    const el = nodes.find(n => visible(n) && texts.some(t => (n.innerText || n.textContent || '').trim().toLowerCase().includes(t)));
    if (!el) return false;
    el.click();
    return true;
    """
    return bool(driver.execute_script(script, texts))


def js_find_visible_button(driver, text: str, exact: bool = True):
    script = """
    const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
    const exact = Boolean(arguments[1]);
    const visible = el => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
    };
    const disabled = el => {
      const tokens = String(el.className || '').toLowerCase().split(/\\s+/);
      return el.disabled || el.getAttribute('aria-disabled') === 'true' || tokens.some(token => token === 'disabled' || token.includes('is-disabled'));
    };
    const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
    const nodes = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'));
    const candidates = nodes
      .map(el => {
        const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
        const r = clickable.getBoundingClientRect();
        const label = norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'));
        const nativeButton = Boolean(clickable.matches('button, [role="button"], input[type="button"], input[type="submit"], a'));
        return { el: clickable, label, r, nativeButton };
      })
      .filter(item => visible(item.el) && !disabled(item.el))
      .filter(item => item.nativeButton || (item.r.width <= 360 && item.r.height <= 90))
      .filter(item => exact ? item.label === target : item.label.includes(target));
    candidates.sort((a, b) => (b.r.top - a.r.top) || (b.r.left - a.r.left));
    return candidates.length ? candidates[0].el : null;
    """
    return driver.execute_script(script, text, exact)


def browser_click_element(driver, element) -> bool:
    browser_wait_for_portal_idle(driver)
    return browser_click_element_no_idle(driver, element)


def browser_click_element_no_idle(driver, element) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element)
        browser_sleep(0.2)
    except Exception:
        pass
    try:
        from selenium.webdriver.common.action_chains import ActionChains

        ActionChains(driver).move_to_element(element).pause(0.1).click(element).perform()
        return True
    except Exception:
        pass
    try:
        element.click()
        return True
    except Exception:
        pass
    try:
        return bool(
            driver.execute_script(
                """
                const el = arguments[0];
                const r = el.getBoundingClientRect();
                const opts = {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                  clientX: r.left + r.width / 2,
                  clientY: r.top + r.height / 2
                };
                for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                  el.dispatchEvent(new MouseEvent(type, opts));
                }
                return true;
                """,
                element,
            )
        )
    except Exception:
        return False


def browser_click_button(driver, text: str, exact: bool = True) -> bool:
    browser_wait_for_portal_idle(driver)
    element = js_find_visible_button(driver, text, exact=exact)
    if not element:
        return False
    return browser_click_element(driver, element)


def browser_click_first_button(driver, labels: list[str], exact: bool = True) -> bool:
    for label in labels:
        if browser_click_button(driver, label, exact=exact):
            return True
    return False


def browser_child_modal_is_open(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent || '');
                if (text.includes('register child') || text.includes('without abha') || text.includes('choose an option to continue')) return true;
                if (text.includes('add new child details') && (text.includes('birth date') || text.includes('child name') || text.includes('save'))) return true;
                return Array.from(document.querySelectorAll('[role="dialog"], [aria-modal="true"], form#modal-form'))
                  .some(el => visible(el));
                """
            )
        )
    except Exception:
        return False


def browser_click_add_child_button(driver) -> bool:
    browser_wait_for_portal_idle(driver)
    if browser_register_child_screen_is_open(driver) or browser_find_student_form_root(driver) or browser_child_modal_is_open(driver):
        return False
    element = driver.execute_script(
        """
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const disabled = el => el.disabled || el.getAttribute('aria-disabled') === 'true' || String(el.className || '').toLowerCase().includes('is-disabled');
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const addLabel = text => {
          if (!text || text.length > 60) return false;
          return text === 'add child'
            || text === 'add student'
            || text.endsWith(' add child')
            || text.endsWith(' add student')
            || text.includes('add child')
            || text.includes('add student');
        };
        const candidates = Array.from(document.querySelectorAll('button, [role="button"], a, input[type="button"], input[type="submit"], div, span'))
          .map(el => {
            const clickable = el.closest('button, [role="button"], a, input[type="button"], input[type="submit"]') || el;
            return {
              el: clickable,
              label: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')),
              rect: clickable.getBoundingClientRect(),
              nativeButton: Boolean(clickable.matches('button, [role="button"], a, input[type="button"], input[type="submit"]'))
            };
          })
          .filter(item => addLabel(item.label) && visible(item.el) && !disabled(item.el))
          .filter(item => item.nativeButton || (item.rect.width <= 420 && item.rect.height <= 110));
        candidates.sort((a, b) => (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
        return candidates.length ? candidates[0].el : null;
        """
    )
    if element:
        return browser_click_element(driver, element)
    return browser_click_first_button(driver, ["Add Child", "Add Student"], exact=True)


def browser_session_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return any(
        needle in text
        for needle in [
            "invalid session id",
            "no such window",
            "target window already closed",
            "disconnected",
            "chrome not reachable",
            "web view not found",
        ]
    )


def browser_has_visible_button(driver, labels: list[str], exact: bool = True) -> bool:
    for label in labels:
        try:
            if js_find_visible_button(driver, label, exact=exact):
                return True
        except Exception:
            continue
    return False


def browser_wait_for_script(driver, script: str, timeout: float = 8.0, interval: float = 0.35) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if driver.execute_script(script):
                return True
        except Exception:
            pass
        browser_sleep(interval)
    return False


def browser_portal_busy_state(driver) -> dict[str, Any]:
    try:
        return driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const all = Array.from(document.querySelectorAll('body *')).filter(visible);
            const bodyText = norm(document.body.innerText || document.body.textContent);
            const buttonLabel = el => norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'));
            const addChildReady = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'))
              .some(el => {
                if (!visible(el)) return false;
                const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
                const r = clickable.getBoundingClientRect();
                const smallClickable = r.width <= 360 && r.height <= 90;
                const nativeButton = clickable.matches('button, [role="button"], input[type="button"], input[type="submit"], a');
                return (nativeButton || smallClickable) && ['add child', 'add student'].includes(buttonLabel(el));
              });
            const formOpen = Array.from(document.querySelectorAll('[role="dialog"], .modal, [class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"], form'))
              .filter(visible)
              .some(el => {
                const text = norm(el.innerText || el.textContent);
                const fields = el.querySelectorAll('input:not([type="hidden"]), textarea, select, [role="combobox"]').length;
                return fields >= 4 && /(child|student|mother|father|mobile|birth)/.test(text);
              });
            const loadingText = all.find(el => {
              const text = norm(el.innerText || el.textContent);
              return text.length <= 80 && /loading|please wait|processing|saving/.test(text);
            });
            const spinner = all.find(el => {
              const cls = String(el.className || '').toLowerCase();
              const style = getComputedStyle(el);
              return cls.includes('animate-spin') || cls.includes('spinner') || style.animationName.toLowerCase().includes('spin');
            });
            const overlay = all.find(el => {
              const cls = String(el.className || '').toLowerCase();
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              const large = r.width >= window.innerWidth * 0.75 && r.height >= window.innerHeight * 0.55;
              const blackOverlay = cls.includes('bg-black') || s.backgroundColor.includes('0, 0, 0');
              const absoluteOverlay = cls.includes('absolute') || cls.includes('fixed') || ['absolute', 'fixed'].includes(s.position);
              return large && blackOverlay && absoluteOverlay;
            });
            const visibleLoading = Boolean(loadingText);
            const overlayLoading = Boolean(overlay && loadingText);
            const busy = Boolean(
              formOpen
                ? (visibleLoading || overlayLoading)
                : (!addChildReady && (visibleLoading || overlayLoading))
            );
            return {
              busy,
              formOpen,
              addChildReady,
              loadingText: loadingText ? String(loadingText.innerText || loadingText.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 160) : '',
              hasSpinner: Boolean(spinner),
              hasOverlay: Boolean(overlay),
              bodyHint: bodyText.slice(0, 240)
            };
            """
        )
    except Exception as exc:
        return {"busy": False, "error": str(exc)}


def browser_wait_for_portal_idle(driver, timeout: float = 25.0, interval: float = 0.5) -> None:
    deadline = time.time() + timeout
    last_state: dict[str, Any] = {}
    while time.time() < deadline:
        last_state = browser_portal_busy_state(driver)
        if not last_state.get("busy"):
            return
        browser_sleep(interval)
    hint = last_state.get("loadingText") or last_state.get("bodyHint") or "loading overlay"
    raise RuntimeError(f"Portal is still loading. Wait or refresh the Screening page, then retry. Last state: {hint}")


def browser_find_student_form_root(driver):
    return driver.execute_script(
        """
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const enabled = el => visible(el) && !el.disabled && !el.readOnly && el.getAttribute('aria-disabled') !== 'true';
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const buttonLabel = el => norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'));
        const saveLabels = ['save child', 'save student', 'save', 'submit', 'create', 'add', 'add child'];
        const fieldWords = ['child name', 'student name', 'birth date', 'date of birth', 'mother', 'father', 'mobile', 'gender'];
        const fieldNames = ['studentname', 'mothername', 'fathername', 'mobile', 'gender', 'birthweight'];
        const badListingWords = ['rows per page', 'copy excel pdf', 'no data available', 'page 1 of 0'];
        const roots = Array.from(document.querySelectorAll([
          '[role="dialog"]',
          '.modal',
          '[class*="modal"]',
          '[class*="Modal"]',
          '[class*="dialog"]',
          '[class*="Dialog"]',
          '[class*="drawer"]',
          '[class*="Drawer"]',
          'form',
          'section',
          'main > div',
          'body > div > div'
        ].join(','))).filter(visible);

        const scored = [];
        for (const root of roots) {
          const text = norm(root.innerText || root.textContent);
          const fields = Array.from(root.querySelectorAll('input:not([type="hidden"]), textarea, select, [role="combobox"]')).filter(enabled);
          const buttons = Array.from(root.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a')).filter(visible);
          const buttonTexts = buttons.map(buttonLabel).filter(Boolean);
          const hasSave = buttonTexts.some(t => saveLabels.includes(t) || saveLabels.some(s => t === s || t.startsWith(s + ' ')));
          const hasFieldWords = fieldWords.filter(w => text.includes(w)).length;
          const names = fields.map(el => norm(`${el.name || ''} ${el.id || ''} ${el.placeholder || ''}`)).join(' ');
          const hasFieldNames = fieldNames.filter(w => names.includes(w)).length;
          const looksLikeListing = badListingWords.some(w => text.includes(w)) || (text.includes('copy') && text.includes('excel') && text.includes('pdf'));
          const r = root.getBoundingClientRect();
          const modalish = root.matches('[role="dialog"], .modal, [class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"], [class*="drawer"], [class*="Drawer"]');
          if (fields.length >= 4 && hasSave && (hasFieldWords >= 2 || hasFieldNames >= 3) && !looksLikeListing) {
            scored.push({ root, score: (modalish ? 1000 : 0) + (20 - Math.min(fields.length, 20)) + hasFieldWords + hasFieldNames, area: r.width * r.height });
          }
        }
        scored.sort((a, b) => (b.score - a.score) || (a.area - b.area));
        return scored.length ? scored[0].root : null;
        """
    )


def browser_student_form_is_open(driver) -> bool:
    return bool(browser_find_student_form_root(driver))


def browser_wait_for_student_form(driver, timeout: float = 8.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        root = browser_find_student_form_root(driver)
        if root:
            return root
        browser_sleep(0.35)
    return None


def browser_click_button_in_root(driver, root, labels: list[str], exact: bool = True) -> bool:
    element = driver.execute_script(
        """
        const root = arguments[0];
        const labels = arguments[1].map(x => String(x || '').toLowerCase().replace(/\\s+/g, ' ').trim());
        const exact = Boolean(arguments[2]);
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const disabled = el => {
          const tokens = String(el.className || '').toLowerCase().split(/\\s+/);
          return el.disabled || el.getAttribute('aria-disabled') === 'true' || tokens.some(token => token === 'disabled' || token.includes('is-disabled'));
        };
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const buttons = Array.from(root.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a'))
          .filter(el => visible(el) && !disabled(el))
          .map(el => ({ el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')), r: el.getBoundingClientRect() }))
          .filter(item => labels.some(label => exact ? item.text === label : item.text.includes(label)));
        buttons.sort((a, b) => (b.r.top - a.r.top) || (b.r.left - a.r.left));
        return buttons.length ? buttons[0].el : null;
        """,
        root,
        labels,
        exact,
    )
    if not element:
        return False
    return browser_click_element(driver, element)


def browser_find_visible_save_button(driver):
    return driver.execute_script(
        """
        const labels = ['save child', 'save student', 'save', 'submit'];
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const disabled = el => {
          const tokens = String(el.className || '').toLowerCase().split(/\\s+/);
          return el.disabled || el.getAttribute('aria-disabled') === 'true' || tokens.some(token => token === 'disabled' || token.includes('is-disabled'));
        };
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const candidates = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"]'))
          .filter(el => visible(el) && !disabled(el))
          .map(el => {
            const text = norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'));
            const r = el.getBoundingClientRect();
            return { el, text, tag: el.tagName, rect: r };
          })
          .filter(item => labels.includes(item.text))
          .sort((a, b) => {
            const rankA = labels.indexOf(a.text);
            const rankB = labels.indexOf(b.text);
            if (rankA !== rankB) return rankA - rankB;
            return (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left);
          });
        return candidates.length ? candidates[0].el : null;
        """
    )


def browser_click_visible_save_button(driver) -> bool:
    browser_wait_for_portal_idle(driver)
    element = browser_find_visible_save_button(driver)
    if not element:
        return False
    if browser_click_element(driver, element):
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const el = arguments[0];
                el.scrollIntoView({ block: 'center', inline: 'center' });
                el.focus();
                el.click();
                return true;
                """,
                element,
            )
        )
    except Exception:
        return False


def browser_submit_modal_form(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const form = document.querySelector('form#modal-form') || Array.from(document.querySelectorAll('form')).find(visible);
                if (!form) return false;
                if (typeof form.requestSubmit === 'function') {
                  form.requestSubmit();
                  return true;
                }
                const event = new Event('submit', { bubbles: true, cancelable: true });
                return form.dispatchEvent(event);
                """
            )
        )
    except Exception:
        return False


def browser_collect_form_validation_errors(driver, root=None) -> list[str]:
    try:
        errors = driver.execute_script(
            """
            const passedRoot = arguments[0] || null;
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
            const snippets = [];
            const addSnippet = text => {
              text = norm(text);
              if (!text) return;
              if (!snippets.includes(text)) snippets.push(text);
            };
            const duplicateSummary = text => {
              text = norm(text);
              if (!/duplicate record found|cannot create a duplicate|already exists/i.test(text)) return '';
              const id = text.match(/STUDENT\\/?CHILD ID\\s+([^\\s]+)/i)?.[1] || '';
              const name = text.match(/NAME\\s+(.+?)\\s+DATE OF BIRTH/i)?.[1] || '';
              const dob = text.match(/DATE OF BIRTH\\s+([^\\s]+)/i)?.[1] || '';
              const inst = text.match(/INSTITUTION\\s+(.+?)\\s+ADDRESS/i)?.[1] || '';
              let out = 'Duplicate record already exists in RBSK';
              if (id) out += ` (Student/Child ID ${id})`;
              const parts = [];
              if (name) parts.push(`Name: ${name}`);
              if (dob) parts.push(`DOB: ${dob}`);
              if (inst) parts.push(`Institute: ${inst}`);
              if (parts.length) out += ': ' + parts.join(', ');
              return out;
            };
            const selectors = [
              '[role="alert"]',
              '[aria-live]',
              '[role="dialog"]',
              '[class*="toast"]',
              '[class*="Toast"]',
              '[id*="toast"]',
              '[class*="modal"]',
              '[class*="Modal"]',
              '[class*="dialog"]',
              '[class*="Dialog"]',
              '.error',
              '.invalid',
              '[class*="error"]',
              '[class*="Error"]',
              '[class*="invalid"]',
              '[class*="Invalid"]',
              '[class*="text-red"]',
              '[class*="text-danger"]'
            ];
            const roots = [];
            if (passedRoot) roots.push(passedRoot);
            roots.push(document.body);
            for (const currentRoot of roots) {
              for (const el of Array.from(currentRoot.querySelectorAll(selectors.join(','))).filter(visible)) {
                const text = norm(el.innerText || el.textContent);
                const duplicate = duplicateSummary(text);
                if (duplicate) addSnippet(duplicate);
                else if (text && /(required|invalid|must|error|failed|duplicate|already exists|cannot create|please select|date must)/i.test(text)) addSnippet(text);
              }
              for (const el of Array.from(currentRoot.querySelectorAll('input, textarea, select')).filter(visible)) {
                if (el.validationMessage) addSnippet(el.validationMessage);
              }
            }
            const bodyDuplicate = duplicateSummary(document.body.innerText || document.body.textContent || '');
            if (bodyDuplicate) addSnippet(bodyDuplicate);
            return snippets.slice(0, 6);
            """,
            root,
        )
        return [clean(item) for item in errors if clean(item)]
    except Exception:
        return []


def browser_close_blocking_message_dialog(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
                const dialogs = Array.from(document.querySelectorAll('[role="dialog"], [class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"], .Toastify, [id*="toast"]'))
                  .filter(visible)
                  .filter(el => /(duplicate record found|cannot create a duplicate|already exists|validation|error|failed)/i.test(norm(el.innerText || el.textContent)));
                for (const dialog of dialogs) {
                  const buttons = Array.from(dialog.querySelectorAll('button, [role="button"], a')).filter(visible);
                  const close = buttons.find(btn => /^(close|ok|okay)$/i.test(norm(btn.innerText || btn.textContent || btn.value || btn.getAttribute('aria-label'))));
                  if (close) {
                    close.click();
                    return true;
                  }
                  const iconClose = buttons.find(btn => /close/i.test(norm(btn.getAttribute('aria-label') || btn.title || btn.innerText || btn.textContent)));
                  if (iconClose) {
                    iconClose.click();
                    return true;
                  }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def duplicate_error_from_messages(messages: list[str]) -> str:
    for message in messages:
        text = clean(message)
        lower = text.lower()
        if "duplicate record" in lower or "already exists in rbsk" in lower or "cannot create a duplicate" in lower:
            return text
    return ""


def browser_wait_for_save_result(driver, previous_root=None, timeout: float = 8.0, speed_mode: Any = "") -> None:
    deadline = time.time() + timeout
    last_errors: list[str] = []
    idle_timeout = speed_timeout(speed_mode, 2.0, 4.0)
    poll = speed_wait(speed_mode, 0.18, 0.4)
    while time.time() < deadline:
        try:
            browser_wait_for_portal_idle(driver, timeout=idle_timeout, interval=0.2 if speed_is_fast_safe(speed_mode) else 0.25)
        except Exception:
            browser_sleep(poll)
            continue
        root = browser_find_student_form_root(driver)
        if not root:
            return
        last_errors = browser_collect_form_validation_errors(driver, root)
        if last_errors:
            duplicate_error = duplicate_error_from_messages(last_errors)
            if duplicate_error:
                browser_close_blocking_message_dialog(driver)
                raise DuplicateRecordFound(duplicate_error)
            raise RuntimeError("RBSK validation after Save: " + "; ".join(last_errors))
        browser_sleep(poll)
    if last_errors:
        duplicate_error = duplicate_error_from_messages(last_errors)
        if duplicate_error:
            browser_close_blocking_message_dialog(driver)
            raise DuplicateRecordFound(duplicate_error)
        raise RuntimeError("RBSK validation after Save: " + "; ".join(last_errors))
    raise RuntimeError("Save button was clicked, but the Add Child form stayed open. Check RBSK validation message on the form.")


def browser_rows_grouped_by_institute(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    labels: dict[str, str] = {}
    order: list[str] = []
    for row in rows:
        label = clean(row.get("Institute Name")) or "(no institute)"
        key = normalize_match_text(label) or "__blank__"
        if key not in grouped:
            grouped[key] = []
            labels[key] = label
            order.append(key)
        grouped[key].append(row)
    if len(order) > 1:
        print(f"Manual entry grouping: {len(order)} AWC groups. The tool will finish one AWC before moving to the next.")
    return [row for key in order for row in grouped[key]]


def browser_pending_institute_groups(rows: list[dict[str, Any]], completed: set[str], resume: bool) -> tuple[list[dict[str, Any]], int]:
    groups: list[dict[str, Any]] = []
    current_key = ""
    skipped_count = 0
    for row in rows:
        sr_no = clean(row.get("Sr No"))
        name = clean(row.get("Student Full Name"))
        if resume and sr_no in completed:
            print(f"Skipping already completed Sr No {sr_no}: {name}")
            skipped_count += 1
            continue
        label = clean(row.get("Institute Name")) or "(no institute)"
        key = normalize_match_text(label) or "__blank__"
        if not groups or key != current_key:
            groups.append({"key": key, "name": label, "rows": [], "start_sr": sr_no})
            current_key = key
        groups[-1]["rows"].append(row)
    if groups:
        print(f"Manual institute workflow: {len(groups)} AWC block(s), {sum(len(group['rows']) for group in groups)} pending student(s).")
        for group_index, group in enumerate(groups, start=1):
            print(f"  {group_index}. {group['name']}: {len(group['rows'])} pending from Sr No {group.get('start_sr')}")
    return groups, skipped_count


def browser_dismiss_success_notice(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent);
                if (!/(success|saved|created|inserted|submitted)/.test(pageText)) return false;
                const buttons = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'))
                  .filter(visible)
                  .map(el => {
                    const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
                    return { el: clickable, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')), r: clickable.getBoundingClientRect() };
                  })
                  .filter(item => ['ok', 'okay', 'close', 'done'].includes(item.text))
                  .sort((a, b) => (a.text === 'ok' ? 0 : 1) - (b.text === 'ok' ? 0 : 1) || (a.r.width * a.r.height) - (b.r.width * b.r.height));
                if (!buttons.length) return false;
                buttons[0].el.click();
                return true;
                """
            )
        )
    except Exception:
        return False


def browser_try_open_add_child_current_institute(driver, timeout: float = 25.0, speed_mode: Any = ""):
    if speed_is_fast_safe(speed_mode) and timeout > 10.0:
        timeout = 8.0
    poll = speed_wait(speed_mode, 0.18, 0.5)
    short_poll = speed_wait(speed_mode, 0.06, 0.2)
    idle_timeout = speed_timeout(speed_mode, 3.0, 6.0)
    register_timeout = speed_timeout(speed_mode, 5.0, 10.0)
    form_timeout = speed_timeout(speed_mode, 1.0, 2.0)
    deadline = time.time() + timeout
    while time.time() < deadline:
        root = browser_find_student_form_root(driver)
        if root:
            return root
        try:
            if browser_register_child_screen_is_open(driver):
                root = browser_complete_register_child_step(driver, timeout=register_timeout, speed_mode=speed_mode)
                if root:
                    return root
                browser_sleep(poll)
                continue
            if browser_child_modal_is_open(driver):
                root = browser_wait_for_student_form(driver, timeout=form_timeout)
                if root:
                    return root
                browser_sleep(poll)
                continue
            try:
                browser_wait_for_portal_idle(driver, timeout=idle_timeout, interval=0.2)
            except Exception:
                browser_sleep(poll)
                continue
            browser_dismiss_success_notice(driver)
            browser_sleep(short_poll)
            if browser_click_add_child_button(driver):
                root = browser_complete_register_child_step(driver, timeout=register_timeout, speed_mode=speed_mode)
                if root:
                    return root
        except Exception:
            browser_sleep(poll)
            continue
        browser_sleep(poll)
    return browser_find_student_form_root(driver)


def browser_register_child_screen_is_open(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent || '');
                if (text.includes('register child') && text.includes('without abha')) return true;
                if (text.includes('choose an option to continue') && text.includes('proceed')) return true;
                const nodes = Array.from(document.querySelectorAll('button, [role="button"], div, span, section')).filter(visible);
                const hasWithout = nodes.some(el => norm(el.innerText || el.textContent || el.getAttribute('aria-label')).includes('without abha'));
                const hasProceed = nodes.some(el => norm(el.innerText || el.textContent || el.getAttribute('aria-label')) === 'proceed');
                return hasWithout && hasProceed;
                """
            )
        )
    except Exception:
        return False


def browser_click_register_child_choice(driver, text: str) -> bool:
    return bool(
        driver.execute_script(
            """
            const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const candidates = Array.from(document.querySelectorAll('button, [role="button"], label, div, section'))
              .filter(visible)
              .map(el => ({ el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')), r: el.getBoundingClientRect() }))
              .filter(item => item.text && item.text.includes(target))
              .map(item => ({
                el: item.el.closest('button, [role="button"], label') || item.el,
                text: item.text,
                r: item.r,
                score: (item.text === target ? 0 : 10) + (item.r.width > 120 && item.r.height > 80 ? 0 : 3) + item.text.length / 1000
              }))
              .sort((a, b) => a.score - b.score);
            if (!candidates.length) return false;
            candidates[0].el.scrollIntoView({block:'center', inline:'center'});
            const el = candidates[0].el;
            const r = el.getBoundingClientRect();
            const opts = {bubbles:true, cancelable:true, view:window, clientX:r.left+r.width/2, clientY:r.top+r.height/2};
            for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
              el.dispatchEvent(new MouseEvent(type, opts));
            }
            return true;
            """,
            text,
        )
    )


def browser_find_register_child_button(driver, text: str):
    return driver.execute_script(
        """
        const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const disabled = el => el.disabled || el.getAttribute('aria-disabled') === 'true';
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const pageText = norm(document.body.innerText || document.body.textContent);
        if (!pageText.includes('register child') && !pageText.includes('without abha') && !pageText.includes('choose an option to continue')) return null;
        const buttons = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a'))
          .filter(el => visible(el) && !disabled(el))
          .map(el => ({
            el,
            text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')),
            rect: el.getBoundingClientRect()
          }))
          .filter(item => item.text === target || item.text.endsWith(' ' + target) || item.text.includes(target))
          .sort((a, b) => (a.text === target ? 0 : 1) - (b.text === target ? 0 : 1) || (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
        return buttons.length ? buttons[0].el : null;
        """,
        text,
    )


def browser_click_register_child_proceed(driver) -> bool:
    element = browser_find_register_child_button(driver, "Proceed")
    if element and browser_click_element_no_idle(driver, element):
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent);
                if (!pageText.includes('register child') || !pageText.includes('without abha')) return false;
                const candidates = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'))
                  .filter(visible)
                  .map(el => {
                    const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
                    return {
                      el: clickable,
                      text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')),
                      rect: clickable.getBoundingClientRect()
                    };
                  })
                  .filter(item => item.text === 'proceed')
                  .sort((a, b) => (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
                if (!candidates.length) return false;
                const el = candidates[0].el;
                el.scrollIntoView({block:'center', inline:'center'});
                const r = el.getBoundingClientRect();
                const opts = {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                  clientX: r.left + r.width / 2,
                  clientY: r.top + r.height / 2
                };
                for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                  el.dispatchEvent(new MouseEvent(type, opts));
                }
                return true;
                """
            )
        )
    except Exception:
        return False


def browser_complete_register_child_step(driver, timeout: float = 12.0, speed_mode: Any = ""):
    pause_after_choice = speed_wait(speed_mode, 0.12, 0.35)
    pause_after_proceed = speed_wait(speed_mode, 0.25, 0.8)
    idle_timeout = speed_timeout(speed_mode, 4.0, 8.0)
    form_timeout = speed_timeout(speed_mode, 1.2, 3.0)
    final_form_timeout = speed_timeout(speed_mode, 0.6, 1.0)
    poll = speed_wait(speed_mode, 0.18, 0.4)
    deadline = time.time() + timeout
    while time.time() < deadline:
        root = browser_find_student_form_root(driver)
        if root:
            return root
        if browser_register_child_screen_is_open(driver):
            for _attempt in range(3):
                root = browser_find_student_form_root(driver)
                if root:
                    return root
                browser_click_register_child_choice(driver, "without abha")
                browser_sleep(pause_after_choice)
                if not browser_click_register_child_proceed(driver):
                    browser_click_register_child_choice(driver, "proceed")
                browser_sleep(pause_after_proceed)
                try:
                    browser_wait_for_portal_idle(driver, timeout=idle_timeout)
                except Exception:
                    pass
                root = browser_wait_for_student_form(driver, timeout=form_timeout)
                if root:
                    return root
                if not browser_register_child_screen_is_open(driver):
                    break
            root = browser_wait_for_student_form(driver, timeout=speed_timeout(speed_mode, 1.0, 2.0))
            if root:
                return root
        else:
            try:
                browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 3.0, 6.0))
            except Exception:
                if not browser_register_child_screen_is_open(driver):
                    raise
        browser_sleep(poll)
    return browser_wait_for_student_form(driver, timeout=final_form_timeout)


def prompt_manual_form_open(
    args: argparse.Namespace,
    row: dict[str, Any],
    index: int,
    total: int,
    institute_name: str,
    same_institute: bool = False,
    institute_index: int | None = None,
    institute_total: int | None = None,
    remaining_after_this: int | None = None,
    previous_saved: str = "",
) -> None:
    sr_no = clean(row.get("Sr No"))
    name = clean(row.get("Student Full Name"))
    dob = clean(row.get("Birth Date"))
    institute_line = ""
    if institute_index is not None and institute_total is not None:
        institute_line = f"Institute count: {institute_index}/{institute_total}\n"
    remaining_line = ""
    if remaining_after_this is not None:
        remaining_line = f"After this child, remaining in this AWC: {remaining_after_this}\n"
    saved_line = f"{previous_saved}\n\n" if previous_saved else ""
    if same_institute:
        message = (
            saved_line +
            f"Row {index}/{total}\n"
            f"{institute_line}"
            f"Sr No: {sr_no}\n"
            f"Child: {name}\n"
            f"DOB: {dob or '(from Excel)'}\n"
            f"Same AWC NAME: {institute_name or '(same institute)'}\n\n"
            f"{remaining_line}"
            "The same institute should already be selected in RBSK.\n\n"
            "In RBSK do only this now:\n"
            "1. Click Add Student/Add Child.\n"
            "2. If RBSK stops on Register Child, click Proceed.\n"
            "3. Make sure the final form is visible with Child Name, Birth Date, Gender, Mother, Father, and Save.\n\n"
            "Only then click OK here. The tool will fill and save this row."
        )
    else:
        message = (
            saved_line +
            f"Row {index}/{total}\n"
            f"{institute_line}"
            f"Sr No: {sr_no}\n"
            f"Child: {name}\n"
            f"DOB: {dob or '(from Excel)'}\n"
            f"AWC NAME: {institute_name or '(use the correct institute for this row)'}\n\n"
            f"{remaining_line}"
            "In RBSK do these steps now:\n"
            "1. Select the above Institute/AWC NAME.\n"
            "2. If Add Student/Add Child is visible, click it directly.\n"
            "3. Click Search only if Add Student/Add Child is not visible yet.\n"
            "4. If RBSK stops on Register Child, click Proceed.\n"
            "5. Make sure the final form is visible with Child Name, Birth Date, Gender, Mother, Father, and Save.\n\n"
            "Only then click OK here. The tool will fill and save this row."
        )
    emit_progress(
        args,
        "manual_prompt",
        step="open_add_child" if same_institute else "select_institute",
        title="Open Add Student Form",
        sr_no=sr_no,
        name=name,
        dob=dob,
        institute=institute_name,
        index=index,
        total=total,
        remaining_after=remaining_after_this,
    )
    callback = getattr(args, "manual_form_open_callback", None)
    if callback:
        callback("Open Add Student Form", message)
    else:
        print("\nManual form open required:")
        print(message)
        safe_input("Press Enter here only after the Add Student/Add Child form is open...")


def prompt_manual_register_proceed(
    args: argparse.Namespace,
    row: dict[str, Any],
    index: int,
    total: int,
    institute_name: str,
    remaining_after_this: int | None = None,
    previous_saved: str = "",
) -> None:
    sr_no = clean(row.get("Sr No"))
    name = clean(row.get("Student Full Name"))
    remaining_line = ""
    if remaining_after_this is not None:
        remaining_line = f"After this child, remaining in this AWC: {remaining_after_this}\n"
    saved_line = f"{previous_saved}\n\n" if previous_saved else ""
    message = (
        saved_line +
        f"Row {index}/{total}\n"
        f"Sr No: {sr_no}\n"
        f"Child: {name}\n"
        f"AWC NAME: {institute_name or '(same institute)'}\n\n"
        f"{remaining_line}"
        "RBSK is on the Register Child screen.\n\n"
        "Do only this now:\n"
        "1. Click Proceed.\n"
        "2. Wait until the final form is visible with Child Name, Birth Date, Gender, Mother, Father, and Save.\n"
        "3. Then click OK here."
    )
    emit_progress(
        args,
        "manual_prompt",
        step="click_proceed",
        title="Click Proceed",
        sr_no=sr_no,
        name=name,
        institute=institute_name,
        index=index,
        total=total,
        remaining_after=remaining_after_this,
    )
    callback = getattr(args, "manual_form_open_callback", None)
    if callback:
        callback("Click Proceed", message)
    else:
        print("\nRegister Child step needs one click:")
        print(message)
        safe_input("Press Enter here after clicking Proceed and seeing the final form...")


def wait_for_manual_student_form(
    driver,
    args: argparse.Namespace,
    row: dict[str, Any],
    index: int,
    total: int,
    institute_name: str,
    same_institute: bool = False,
    institute_index: int | None = None,
    institute_total: int | None = None,
    remaining_after_this: int | None = None,
    previous_saved: str = "",
    force_prompt: bool = False,
):
    if not force_prompt:
        root = browser_find_student_form_root(driver)
        if root:
            return root
        try:
            root = browser_complete_register_child_step(driver, timeout=3.0)
            if root:
                return root
        except Exception:
            pass
    for attempt in range(2):
        if browser_register_child_screen_is_open(driver):
            prompt_manual_register_proceed(
                args,
                row,
                index,
                total,
                institute_name,
                remaining_after_this=remaining_after_this,
                previous_saved=previous_saved,
            )
        else:
            prompt_manual_form_open(
                args,
                row,
                index,
                total,
                institute_name,
                same_institute=same_institute,
                institute_index=institute_index,
                institute_total=institute_total,
                remaining_after_this=remaining_after_this,
                previous_saved=previous_saved,
            )
        browser_wait_for_portal_idle(driver, timeout=20.0)
        root = browser_complete_register_child_step(driver, timeout=8.0)
        if root:
            return root
        root = browser_wait_for_student_form(driver, timeout=8.0)
        if root:
            return root
        if browser_register_child_screen_is_open(driver):
            print("Could not detect the final Add Student/Add Child form yet. RBSK is still on Register Child; click Proceed and confirm again.")
        else:
            print("Could not detect the final Add Student/Add Child form yet. Click Add Child/Add Student and confirm again.")
    raise RuntimeError("Manual handoff failed: Add Student/Add Child form was not detected.")


def browser_click_visible_row_text(driver, text: str) -> bool:
    return bool(
        driver.execute_script(
            """
            const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const nodes = Array.from(document.querySelectorAll('tr, [role="row"], tbody *, div, span')).filter(visible);
            const match = nodes
              .map(el => ({ el, text: norm(el.innerText || el.textContent), r: el.getBoundingClientRect() }))
              .filter(item => item.text && (item.text === target || item.text.includes(target) || target.includes(item.text)))
              .sort((a, b) => {
                const exactA = a.text === target ? 0 : 1;
                const exactB = b.text === target ? 0 : 1;
                if (exactA !== exactB) return exactA - exactB;
                if (a.text.length !== b.text.length) return a.text.length - b.text.length;
                return (a.r.width * a.r.height) - (b.r.width * b.r.height);
              })[0];
            if (!match) return false;
            match.el.scrollIntoView({block: 'center', inline: 'center'});
            match.el.click();
            return true;
            """,
            text,
        )
    )


def browser_wait_for_institute_result(driver, institute_name: str, timeout: float = 8.0) -> bool:
    script = """
    const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
    const visible = el => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
    };
    const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
    const addButton = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'))
      .some(el => {
        if (!visible(el)) return false;
        const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
        const r = clickable.getBoundingClientRect();
        const smallClickable = r.width <= 360 && r.height <= 90;
        const nativeButton = clickable.matches('button, [role="button"], input[type="button"], input[type="submit"], a');
        return (nativeButton || smallClickable) && ['add child', 'add student'].includes(norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')));
      });
    if (!addButton) return false;
    if (!target) return true;
    const rowLike = Array.from(document.querySelectorAll('tr, [role="row"], tbody *, div, span'))
      .filter(visible)
      .some(el => {
        const r = el.getBoundingClientRect();
        const text = norm(el.innerText || el.textContent);
        return r.top > 240 && text && (text === target || text.includes(target) || target.includes(text));
      });
    return rowLike;
    """
    return browser_wait_for_script(driver, script, timeout=timeout)


def browser_institute_dropdown_matches(driver, institute_name: str) -> bool:
    wanted = normalize_match_text(institute_name)
    state = js_react_control_state(driver, ["institutes", "institution"])
    selected = normalize_match_text(state.get("text", ""))
    return bool(wanted and selected and selected != "select institution" and (wanted in selected or selected in wanted))


def js_select_institute_option(driver, institute_name: str, allow_partial: bool = False) -> bool:
    from selenium.webdriver.common.keys import Keys

    wanted = normalize_match_text(institute_name)
    if not wanted:
        return False
    browser_wait_for_portal_idle(driver)
    control = driver.execute_script(
        """
        const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/\\s+/g, ' ').trim();
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const labels = Array.from(document.querySelectorAll('body *'))
          .filter(visible)
          .map(el => ({ el, text: norm(el.innerText || el.textContent), r: el.getBoundingClientRect() }))
          .filter(item => item.text === 'institutes' || item.text === 'institutes *')
          .sort((a, b) => a.r.top - b.r.top || a.r.left - b.r.left);
        if (!labels.length) return null;
        const label = labels[0];
        const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
          .filter(visible)
          .map(el => ({ el, r: el.getBoundingClientRect(), text: norm(el.innerText || el.textContent || el.value), cls: String(el.className || ''), aria: el.getAttribute('aria-disabled') || '' }))
          .filter(item => item.r.width >= 120 && item.r.height >= 25 && item.r.top >= label.r.bottom - 10)
          .filter(item => Math.abs(item.r.left - label.r.left) < 180 || (item.r.left >= label.r.left - 20 && item.r.left <= label.r.left + 420))
          .sort((a, b) => (a.r.top - label.r.bottom) - (b.r.top - label.r.bottom) || a.r.left - b.r.left);
        const enabled = controls.find(item => item.aria !== 'true' && !item.cls.includes('is-disabled'));
        return enabled ? enabled.el : null;
        """
    )
    if not control:
        return False

    for attempt in range(4):
        try:
            driver.execute_script(
                """
                const control = arguments[0];
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const root = control.closest('[class*="container"], [class*="Container"]') || control.parentElement || control;
                const clear = Array.from(root.querySelectorAll('button, [role="button"], div, span'))
                  .filter(visible)
                  .find(el => {
                    const cls = String(el.className || '').toLowerCase();
                    const txt = String(el.innerText || el.textContent || el.getAttribute('aria-label') || '').trim().toLowerCase();
                    return cls.includes('clear') || txt === 'x' || txt === '×' || txt.includes('clear');
                  });
                if (clear) clear.click();
                """,
                control,
            )
        except Exception:
            pass
        browser_sleep(0.2)
        if not browser_click_element(driver, control):
            return False
        browser_sleep(0.25)
        active = None
        try:
            active = driver.execute_script(
                """
                const control = arguments[0];
                const root = control.closest('[class*="container"], [class*="Container"]') || control.parentElement || control;
                return root.querySelector('input') || document.activeElement;
                """,
                control,
            )
            active.click()
            active.send_keys(Keys.CONTROL, "a")
            active.send_keys(Keys.BACKSPACE)
            active.send_keys(institute_name)
        except Exception:
            pass
        browser_sleep(1.0)
        browser_wait_for_portal_idle(driver, timeout=8.0)
        clicked = driver.execute_script(
            """
            const target = arguments[0];
            const allowPartial = arguments[1];
            const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/[()\\[\\]{}.,_-]/g, ' ').replace(/\\s+/g, ' ').trim();
            const wanted = norm(target);
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const candidates = Array.from(document.querySelectorAll('[role="option"], div[class*="option"], div[class*="Option"], li, span'))
              .filter(visible)
              .map(el => ({ el, text: norm(el.innerText || el.textContent || el.getAttribute('aria-label')), r: el.getBoundingClientRect() }))
              .filter(item => item.text && (item.text === wanted || (allowPartial && (item.text.includes(wanted) || wanted.includes(item.text)))))
              .sort((a, b) =>
                (a.text === wanted ? 0 : 1) - (b.text === wanted ? 0 : 1) ||
                a.text.length - b.text.length ||
                a.r.top - b.r.top
              );
            if (!candidates.length) return false;
            const el = candidates[0].el;
            el.scrollIntoView({block:'center', inline:'center'});
            const r = el.getBoundingClientRect();
            const opts = {bubbles:true, cancelable:true, view:window, clientX:r.left+r.width/2, clientY:r.top+r.height/2};
            for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
              el.dispatchEvent(new MouseEvent(type, opts));
            }
            return true;
            """,
            institute_name,
            allow_partial,
        )
        if not clicked:
            try:
                (active or driver.switch_to.active_element).send_keys(Keys.ARROW_DOWN)
                browser_sleep(0.2)
                (active or driver.switch_to.active_element).send_keys(Keys.ENTER)
            except Exception:
                pass
        browser_sleep(1.0)
        browser_wait_for_portal_idle(driver, timeout=10.0)
        if browser_institute_dropdown_matches(driver, institute_name):
            return True
        state = js_react_control_state(driver, ["institutes", "institution"])
        selected = normalize_match_text(state.get("text", ""))
        if selected and selected != "select institution" and (selected == wanted or (allow_partial and (wanted in selected or selected in wanted))):
            return True
    return False


def browser_institute_is_confirmed(driver, institute_name: str) -> bool:
    wanted = normalize_match_text(institute_name)
    if browser_institute_dropdown_matches(driver, institute_name):
        return True
    if not wanted:
        return False
    try:
        return bool(
            driver.execute_script(
                """
                const target = String(arguments[0] || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const addButton = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a, div, span'))
                  .some(el => {
                    if (!visible(el)) return false;
                    const clickable = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a') || el;
                    const r = clickable.getBoundingClientRect();
                    const smallClickable = r.width <= 360 && r.height <= 90;
                    const nativeButton = clickable.matches('button, [role="button"], input[type="button"], input[type="submit"], a');
                    return (nativeButton || smallClickable) && ['add child', 'add student'].includes(norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')));
                  });
                if (!addButton) return false;
                return Array.from(document.querySelectorAll('tr, [role="row"], tbody *, div, span'))
                  .filter(visible)
                  .some(el => {
                    const r = el.getBoundingClientRect();
                    const text = norm(el.innerText || el.textContent);
                    return r.top > 240 && text && (text === target || text.includes(target) || target.includes(text));
                  });
                """,
                wanted,
            )
        )
    except Exception:
        return False


def resolve_institute_alias_path(args: argparse.Namespace) -> Path | None:
    value = clean(getattr(args, "institute_aliases", ""))
    if not value:
        return None
    path = Path(value)
    if path.is_absolute():
        return path
    excel_path = Path(clean(getattr(args, "excel", "")))
    if excel_path.parent and (excel_path.parent / path).exists():
        return excel_path.parent / path
    return Path.cwd() / path


def load_institute_aliases(args: argparse.Namespace) -> dict[str, str]:
    path = resolve_institute_alias_path(args)
    if not path or not path.exists():
        return {}
    rows = load_rows(path, None)
    aliases: dict[str, str] = {}
    for row in rows:
        excel_name = clean(
            row.get("Excel AWC NAME")
            or row.get("Excel AWC Name")
            or row.get("AWC NAME")
            or row.get("Institute Name")
        )
        portal_name = clean(
            row.get("RBSK Institute Name")
            or row.get("Portal Institute Name")
            or row.get("RBSK AWC NAME")
            or row.get("Matched Institute Name")
        )
        if excel_name and portal_name:
            aliases[normalize_match_text(excel_name)] = portal_name
    return aliases


def browser_collect_institute_options(driver) -> list[str]:
    try:
        browser_wait_for_portal_idle(driver)
        control = driver.execute_script(
            """
            const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/\\s+/g, ' ').trim();
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
            };
            const all = Array.from(document.querySelectorAll('body *')).filter(visible);
            const labels = all
              .filter(el => ['institutes', 'institution'].some(h => norm(el.innerText || el.textContent) === h))
              .map(el => ({ el, r: el.getBoundingClientRect() }))
              .sort((a, b) => a.r.top - b.r.top);
            const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
              .filter(visible)
              .map(el => ({ el, r: el.getBoundingClientRect(), text: norm(el.innerText || el.textContent || el.value), cls: String(el.className || ''), aria: el.getAttribute('aria-disabled') || '' }))
              .filter(x => x.r.width >= 120 && x.r.height >= 25);
            for (const label of labels) {
              const candidates = controls
                .filter(c => c.r.top >= label.r.bottom - 8 && Math.abs(c.r.left - label.r.left) < 120)
                .sort((a, b) => (a.r.top - label.r.bottom) - (b.r.top - label.r.bottom));
              if (candidates.length && candidates[0].aria !== 'true' && !candidates[0].cls.includes('is-disabled')) {
                return candidates[0].el;
              }
            }
            return null;
            """
        )
        if not control:
            return []
        if not browser_click_element(driver, control):
            return []
        browser_sleep(1.0)
        options = driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
            };
            const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
            return Array.from(document.querySelectorAll('[role="option"], div[class*="option"], div[class*="Option"], li'))
              .filter(visible)
              .map(el => norm(el.innerText || el.textContent))
              .filter(text => text && !/^select/i.test(text));
            """
        )
        try:
            from selenium.webdriver.common.keys import Keys

            driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        seen: set[str] = set()
        result: list[str] = []
        for option in options or []:
            label = clean(option)
            key = normalize_match_text(label)
            if label and key not in seen:
                seen.add(key)
                result.append(label)
        return result
    except Exception:
        return []


def write_institute_mapping_report(path: Path, rows: list[dict[str, str]]) -> None:
    headers = [
        "Excel AWC NAME",
        "Issue",
        "Matched RBSK Institute Name",
        "Closest RBSK Options",
        "All RBSK Options File",
        "Suggested Action",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({header: row.get(header, "") for header in headers})


def write_institute_options_report(path: Path, options: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    seen: set[str] = set()
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["RBSK Institute Name"])
        writer.writeheader()
        for option in sorted(options, key=lambda value: normalize_match_text(value)):
            label = clean(option)
            key = normalize_match_text(label)
            if not label or key in seen:
                continue
            seen.add(key)
            writer.writerow({"RBSK Institute Name": label})


def closest_institute_options(excel_name: str, options: list[str], limit: int = 8) -> str:
    target = normalize_match_text(excel_name)
    scored: list[tuple[float, str]] = []
    for option in options:
        label = clean(option)
        key = normalize_match_text(label)
        if not label or not key:
            continue
        score = difflib.SequenceMatcher(None, target, key).ratio()
        if target in key or key in target:
            score += 0.2
        scored.append((score, label))
    scored.sort(key=lambda item: (-item[0], normalize_match_text(item[1])))
    return " | ".join(label for score, label in scored[:limit] if score >= 0.35)


def institute_mapping_cache_key(
    rows: list[dict[str, Any]],
    args: argparse.Namespace,
    context: dict[str, Any] | None,
    label: str,
) -> str:
    names = sorted({normalize_match_text(row.get("Institute Name")) for row in rows if clean(row.get("Institute Name"))})
    alias_path = resolve_institute_alias_path(args)
    alias_stamp = ""
    try:
        if alias_path.exists():
            alias_stamp = str(alias_path.stat().st_mtime_ns)
    except Exception:
        alias_stamp = ""
    context_bits = {
        "label": clean(label),
        "login": clean(getattr(args, "default_login_id", "")),
        "state": clean((context or {}).get("state_code")),
        "district": clean((context or {}).get("district_code")),
        "block": clean((context or {}).get("health_block_code")),
        "type": clean((context or {}).get("institution_type")),
        "aliases": alias_stamp,
        "names": names,
    }
    return json.dumps(context_bits, sort_keys=True, ensure_ascii=False)


def prepare_verified_institute_mapping(
    driver,
    rows: list[dict[str, Any]],
    args: argparse.Namespace,
    session: requests.Session | None = None,
    context: dict[str, Any] | None = None,
    label: str = "",
) -> dict[str, str]:
    excel_names: list[str] = []
    seen_excel: set[str] = set()
    for row in rows:
        name = clean(row.get("Institute Name"))
        key = normalize_match_text(name)
        if name and key not in seen_excel:
            seen_excel.add(key)
            excel_names.append(name)
    if not excel_names:
        return {}

    cache_key = institute_mapping_cache_key(rows, args, context, label)
    cached = INSTITUTE_MAPPING_CACHE.get(cache_key)
    if cached:
        emit_progress(args, "institute_mapping", label=label, matched=len(cached), missing=0, total=len(excel_names), report="")
        print(f"Institute mapping cache: using {len(cached)}/{len(excel_names)} verified AWC name(s).")
        return dict(cached)

    options: list[str] = []
    if (
        session is not None
        and context is not None
        and clean(context.get("state_code"))
        and clean(context.get("district_code"))
        and clean(context.get("health_block_code"))
        and clean(context.get("institution_type"))
    ):
        options.extend(item["label"] for item in fetch_institution_options(session, context) if clean(item.get("label")))
    dropdown_options = browser_collect_institute_options(driver) if driver is not None else []
    options.extend(dropdown_options)

    option_index: dict[str, list[str]] = {}
    for option in options:
        text = clean(option)
        key = normalize_match_text(text)
        if not key:
            continue
        labels = option_index.setdefault(key, [])
        if text not in labels:
            labels.append(text)

    aliases = load_institute_aliases(args)
    missing_rows: list[dict[str, str]] = []
    mapping: dict[str, str] = {}
    report_path = Path(args.failed_log).parent / "institute_mapping_missing.csv"
    options_report_path = Path(args.failed_log).parent / "rbsk_institute_options.csv"
    write_institute_options_report(options_report_path, options)

    if not option_index:
        write_institute_mapping_report(
            report_path,
            [
                {
                    "Excel AWC NAME": name,
                    "Issue": "Could not read institute list from RBSK",
                    "Matched RBSK Institute Name": "",
                    "Closest RBSK Options": "",
                    "All RBSK Options File": str(options_report_path),
                    "Suggested Action": "Select State/District/Block and Institute Type AWC, then retry the mapping check.",
                }
                for name in excel_names
            ],
        )
        emit_progress(args, "institute_mapping", label=label, matched=0, missing=len(excel_names), total=len(excel_names), report=str(report_path))
        raise ImportErrorWithHelp(f"Could not read institute list from RBSK. Mapping report: {report_path}")

    for excel_name in excel_names:
        excel_key = normalize_match_text(excel_name)
        alias_name = aliases.get(excel_key)
        target_key = normalize_match_text(alias_name or excel_name)
        matches = option_index.get(target_key, [])
        if len(matches) == 1:
            mapping[excel_key] = matches[0]
            continue
        if len(matches) > 1:
            missing_rows.append(
                {
                    "Excel AWC NAME": excel_name,
                    "Issue": "More than one RBSK institute has the same matched name",
                    "Matched RBSK Institute Name": " | ".join(matches),
                    "Closest RBSK Options": closest_institute_options(excel_name, options),
                    "All RBSK Options File": str(options_report_path),
                    "Suggested Action": "Add an exact row to institute_aliases.csv for this Excel AWC NAME.",
                }
            )
            continue
        if alias_name:
            issue = "Alias name was not found in RBSK list"
            suggested = "Correct RBSK Institute Name in institute_aliases.csv, then retry."
            matched = alias_name
        else:
            issue = "Excel AWC NAME was not found in RBSK list"
            suggested = "Add this row to institute_aliases.csv with the exact RBSK Institute Name."
            matched = ""
        missing_rows.append(
            {
                "Excel AWC NAME": excel_name,
                "Issue": issue,
                "Matched RBSK Institute Name": matched,
                "Closest RBSK Options": closest_institute_options(alias_name or excel_name, options),
                "All RBSK Options File": str(options_report_path),
                "Suggested Action": suggested,
            }
        )

    if missing_rows:
        write_institute_mapping_report(report_path, missing_rows)
        emit_progress(
            args,
            "institute_mapping",
            label=label,
            matched=len(mapping),
            missing=len(missing_rows),
            total=len(excel_names),
            report=str(report_path),
        )
        missing_names = ", ".join(row["Excel AWC NAME"] for row in missing_rows[:10])
        if len(missing_rows) > 10:
            missing_names += f", +{len(missing_rows) - 10} more"
        allow_manual_fallback = bool(
            (getattr(args, "manual_form_open", True) or getattr(args, "unattended_after_start", False))
            and not getattr(args, "mapping_check_only", False)
            and getattr(args, "allow_unmapped_institute_manual_fallback", True)
        )
        if allow_manual_fallback:
            setattr(args, "unmapped_institute_names", {normalize_match_text(row["Excel AWC NAME"]) for row in missing_rows})
            print(
                f"Institute mapping warning: {len(missing_rows)} AWC name(s) need manual selection: {missing_names}"
            )
            print(f"Mapping report: {report_path}")
            if getattr(args, "unattended_after_start", False):
                print("Unattended mode: unmatched AWC blocks will be skipped and logged instead of stopping the run.")
            else:
                print("The run will continue. For each missing AWC, select the institute manually once when prompted.")
            INSTITUTE_MAPPING_CACHE[cache_key] = dict(mapping)
            return mapping
        raise ImportErrorWithHelp(
            f"Institute mapping check failed: {len(missing_rows)} AWC name(s) need review.\n"
            f"Missing AWC name(s): {missing_names}\n"
            f"Open this file in Excel and add/correct aliases: {report_path}"
        )

    emit_progress(args, "institute_mapping", label=label, matched=len(mapping), missing=0, total=len(excel_names), report="")
    print(f"Institute mapping check: {len(mapping)}/{len(excel_names)} AWC name(s) matched.")
    if aliases:
        print(f"Alias file loaded: {resolve_institute_alias_path(args)}")
    INSTITUTE_MAPPING_CACHE[cache_key] = dict(mapping)
    return mapping


def save_button_diagnostic(driver, path: Path) -> None:
    data = driver.execute_script(
        """
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
        const buttons = Array.from(document.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"], a'))
          .filter(visible)
          .map((el, index) => {
            const r = el.getBoundingClientRect();
            return {
              index,
              tag: el.tagName,
              text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')),
              disabled: Boolean(el.disabled || el.getAttribute('aria-disabled') === 'true'),
              className: String(el.className || ''),
              rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
            };
          });
        const controls = Array.from(document.querySelectorAll('input, textarea, select, [role="combobox"]'))
          .filter(visible)
          .map((el, index) => {
            const r = el.getBoundingClientRect();
            const wrapper = el.closest('label, div, fieldset, section') || el.parentElement;
            return {
              index,
              tag: el.tagName,
              type: el.getAttribute('type') || el.getAttribute('role') || '',
              name: el.getAttribute('name') || '',
              id: el.id || '',
              placeholder: el.getAttribute('placeholder') || '',
              aria: el.getAttribute('aria-label') || '',
              value: el.value || norm(el.innerText || el.textContent).slice(0, 120),
              disabled: Boolean(el.disabled || el.readOnly || el.getAttribute('aria-disabled') === 'true'),
              nearText: norm(wrapper?.innerText || '').slice(0, 180),
              rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
            };
          });
        const dropdowns = Array.from(document.querySelectorAll('[class*="control"], [class*="container"], [role="combobox"]'))
          .filter(visible)
          .map((el, index) => {
            const r = el.getBoundingClientRect();
            return {
              index,
              text: norm(el.innerText || el.textContent || el.value).slice(0, 160),
              className: String(el.className || ''),
              rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
            };
          })
          .filter(item => item.text || item.className.includes('select'));
        const saveLabels = ['save child', 'save student', 'save', 'submit'];
        const saveCandidates = buttons
          .filter(item => saveLabels.includes(String(item.text || '').toLowerCase()))
          .map(item => ({ text: item.text, disabled: item.disabled, rect: item.rect, className: item.className }));
        const formOpen = Array.from(document.querySelectorAll('[role="dialog"], .modal, [class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"], form, section, main > div'))
          .filter(visible)
          .some(root => {
            const text = norm(root.innerText || root.textContent).toLowerCase();
            const fields = root.querySelectorAll('input:not([type="hidden"]), textarea, select, [role="combobox"]').length;
            const hasSave = Array.from(root.querySelectorAll('button, [role="button"], input[type="button"], input[type="submit"]'))
              .filter(visible)
              .some(el => saveLabels.includes(norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')).toLowerCase()));
            return fields >= 4 && hasSave && /(child name|birth date|mother|father|mobile)/.test(text);
          });
        return { url: location.href, title: document.title, buttons, controls, dropdowns, saveCandidates, formOpen, bodyText: norm(document.body.innerText).slice(0, 6000) };
        """
    )
    data["busyState"] = browser_portal_busy_state(driver)
    rows = "\n".join(
        "<tr><td>{index}</td><td>{tag}</td><td>{disabled}</td><td>{rect}</td><td>{text}</td><td>{className}</td></tr>".format(
            index=item.get("index", ""),
            tag=item.get("tag", ""),
            disabled=item.get("disabled", ""),
            rect=json.dumps(item.get("rect", {})),
            text=str(item.get("text", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            className=str(item.get("className", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
        )
        for item in data.get("buttons", [])
    )
    control_rows = "\n".join(
        "<tr><td>{index}</td><td>{tag}</td><td>{type}</td><td>{disabled}</td><td>{rect}</td><td>{name}</td><td>{id}</td><td>{placeholder}</td><td>{value}</td><td>{nearText}</td></tr>".format(
            index=item.get("index", ""),
            tag=item.get("tag", ""),
            type=str(item.get("type", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            disabled=item.get("disabled", ""),
            rect=json.dumps(item.get("rect", {})),
            name=str(item.get("name", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            id=str(item.get("id", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            placeholder=str(item.get("placeholder", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            value=str(item.get("value", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            nearText=str(item.get("nearText", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
        )
        for item in data.get("controls", [])
    )
    dropdown_rows = "\n".join(
        "<tr><td>{index}</td><td>{rect}</td><td>{text}</td><td>{className}</td></tr>".format(
            index=item.get("index", ""),
            rect=json.dumps(item.get("rect", {})),
            text=str(item.get("text", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            className=str(item.get("className", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
        )
        for item in data.get("dropdowns", [])
    )
    save_rows = "\n".join(
        "<tr><td>{text}</td><td>{disabled}</td><td>{rect}</td><td>{className}</td></tr>".format(
            text=str(item.get("text", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
            disabled=item.get("disabled", ""),
            rect=json.dumps(item.get("rect", {})),
            className=str(item.get("className", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
        )
        for item in data.get("saveCandidates", [])
    )
    body = str(data.get("bodyText", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    busy_state = str(json.dumps(data.get("busyState", {}), indent=2)).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    html = f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>RBSK Button Diagnostic</title>
  <style>body{{font-family:Arial,sans-serif;margin:24px}}table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #ccc;padding:6px;text-align:left;vertical-align:top}}pre{{white-space:pre-wrap;background:#f5f5f5;padding:12px}}</style>
</head>
<body>
  <h1>RBSK Button Diagnostic</h1>
  <p><strong>URL:</strong> {str(data.get("url", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")}</p>
  <p><strong>Title:</strong> {str(data.get("title", "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")}</p>
  <p><strong>Add Child form detected:</strong> {data.get("formOpen")}</p>
  <h2>Loading State</h2>
  <pre>{busy_state}</pre>
  <h2>Save Button Candidates</h2>
  <table><thead><tr><th>Text</th><th>Disabled</th><th>Rect</th><th>Class</th></tr></thead><tbody>{save_rows}</tbody></table>
  <h2>Buttons</h2>
  <table><thead><tr><th>#</th><th>Tag</th><th>Disabled</th><th>Rect</th><th>Text</th><th>Class</th></tr></thead><tbody>{rows}</tbody></table>
  <h2>Inputs And Selects</h2>
  <table><thead><tr><th>#</th><th>Tag</th><th>Type</th><th>Disabled</th><th>Rect</th><th>Name</th><th>ID</th><th>Placeholder</th><th>Value</th><th>Nearby Text</th></tr></thead><tbody>{control_rows}</tbody></table>
  <h2>Dropdown States</h2>
  <table><thead><tr><th>#</th><th>Rect</th><th>Text</th><th>Class</th></tr></thead><tbody>{dropdown_rows}</tbody></table>
  <h2>Visible Page Text</h2>
  <pre>{body}</pre>
</body>
</html>
"""
    path.write_text(html, encoding="utf-8")


def js_fill_by_label(driver, label_hints: list[str], value: Any, root=None) -> bool:
    script = """
    const hints = arguments[0].map(t => String(t).toLowerCase());
    const value = arguments[1] == null ? '' : String(arguments[1]);
    const scope = arguments[2] || document;
    const visible = el => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
    };
    const setValue = (el, value) => {
      el.focus();
      const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
      if (setter) setter.call(el, value);
      else el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.blur();
      return true;
    };
    const controls = Array.from(scope.querySelectorAll('input, textarea, select')).filter(visible);
    for (const el of controls) {
      const id = el.id ? document.querySelector(`label[for="${CSS.escape(el.id)}"]`) : null;
      const wrapper = el.closest('label, div, fieldset, section') || el.parentElement;
      const hay = [
        el.name, el.id, el.placeholder, el.getAttribute('aria-label'),
        id?.innerText, wrapper?.innerText
      ].filter(Boolean).join(' ').toLowerCase();
      if (hints.some(h => hay.includes(h))) return setValue(el, value);
    }
    return false;
    """
    return bool(driver.execute_script(script, label_hints, "" if value is None else str(value), root))


def js_clear_react_control(driver, label_hints: list[str]) -> bool:
    browser_wait_for_portal_idle(driver)
    control = driver.execute_script(
        """
        const hints = arguments[0].map(t => String(t).toLowerCase());
        const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/\\s+/g, ' ').trim();
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const all = Array.from(document.querySelectorAll('body *')).filter(visible);
        const labels = all
          .filter(el => hints.some(h => norm(el.innerText || el.textContent) === h || norm(el.innerText || el.textContent).startsWith(h + ' *')))
          .map(el => ({ el, r: el.getBoundingClientRect() }))
          .sort((a, b) => a.r.top - b.r.top);
        const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
          .filter(visible)
          .map(el => ({ el, r: el.getBoundingClientRect(), text: norm(el.innerText || el.textContent || el.value) }))
          .filter(x => x.r.width >= 120 && x.r.height >= 25);
        for (const label of labels) {
          const candidates = controls
            .filter(c => c.r.top >= label.r.bottom - 8 && Math.abs(c.r.left - label.r.left) < 100)
            .sort((a, b) => (a.r.top - label.r.bottom) - (b.r.top - label.r.bottom));
          if (candidates.length) return candidates[0].el;
        }
        return null;
        """,
        label_hints,
    )
    if not control:
        return False
    try:
        driver.execute_script(
            """
            const control = arguments[0];
            const visible = el => {
              if (!el) return false;
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
            };
            const root = control.closest('[class*="container"], [class*="Container"]') || control.parentElement || control;
            const clear = Array.from(root.querySelectorAll('button, [role="button"], div, span'))
              .filter(visible)
              .find(el => {
                const cls = String(el.className || '').toLowerCase();
                const txt = String(el.innerText || el.textContent || el.getAttribute('aria-label') || '').trim().toLowerCase();
                return cls.includes('clear') || txt === 'x' || txt === '×' || txt.includes('clear');
              });
            if (clear) clear.click();
            """,
            control,
        )
        browser_sleep(0.2)
        browser_click_element(driver, control)
        from selenium.webdriver.common.keys import Keys

        active = driver.execute_script(
            """
            const control = arguments[0];
            const root = control.closest('[class*="container"], [class*="Container"]') || control.parentElement || control;
            return root.querySelector('input') || document.activeElement;
            """,
            control,
        )
        active.click()
        active.send_keys(Keys.CONTROL, "a")
        active.send_keys(Keys.BACKSPACE)
        browser_sleep(0.2)
        return True
    except Exception:
        return False


def js_select_react_option(driver, label_hints: list[str], option_text: str, allow_partial: bool = True) -> bool:
    from selenium.webdriver.common.keys import Keys

    browser_wait_for_portal_idle(driver)
    control = driver.execute_script(
        """
        const hints = arguments[0].map(t => String(t).toLowerCase());
        const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/\\s+/g, ' ').trim();
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };

        const all = Array.from(document.querySelectorAll('body *')).filter(visible);
        const labels = all
          .filter(el => hints.some(h => norm(el.innerText || el.textContent) === h || norm(el.innerText || el.textContent).startsWith(h + ' *')))
          .map(el => ({ el, r: el.getBoundingClientRect() }))
          .sort((a, b) => a.r.top - b.r.top);

        const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
          .filter(visible)
          .map(el => ({ el, r: el.getBoundingClientRect(), text: norm(el.innerText || el.textContent || el.value) }))
          .filter(x => x.r.width >= 120 && x.r.height >= 25);

        for (const label of labels) {
          const candidates = controls
            .filter(c => c.r.top >= label.r.bottom - 8 && Math.abs(c.r.left - label.r.left) < 90)
            .sort((a, b) => (a.r.top - label.r.bottom) - (b.r.top - label.r.bottom));
          if (candidates.length) return candidates[0].el;
        }

        const fallback = controls.find(c => c.text.includes('select institution') || c.text.includes('select mht') || c.text.includes('select'));
        return fallback ? fallback.el : null;
        """,
        label_hints,
    )
    if not control:
        return False

    target_key = normalize_match_text(option_text)
    for attempt in range(3):
        try:
            if not browser_click_element(driver, control):
                state = browser_portal_busy_state(driver)
                if state.get("busy"):
                    raise RuntimeError("Portal is still loading. Wait or refresh Student Master, then retry.")
                return False
        except Exception as exc:
            state = browser_portal_busy_state(driver)
            if state.get("busy") or "click intercepted" in str(exc).lower():
                raise RuntimeError("Portal is still loading. Wait or refresh Student Master, then retry.") from exc
            raise
        browser_sleep(0.3)
        active = None
        try:
            active = driver.execute_script(
                """
                const control = arguments[0];
                const root = control.closest('[class*="container"], [class*="Container"]') || control.parentElement || control;
                return root.querySelector('input') || document.activeElement;
                """,
                control,
            )
            active.click()
            active.send_keys(Keys.CONTROL, "a")
            active.send_keys(Keys.BACKSPACE)
            active.send_keys(option_text)
        except Exception:
            pass
        browser_sleep(0.9)
        browser_wait_for_portal_idle(driver, timeout=10.0)

        clicked = driver.execute_script(
            """
            const target = arguments[0];
            const allowPartial = arguments[1];
            const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/[()\\[\\]{}.,_-]/g, ' ').replace(/\\s+/g, ' ').trim();
            const wanted = norm(target);
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const nodes = Array.from(document.querySelectorAll('[role="option"], div[class*="option"], div[class*="Option"], li, span'))
              .filter(visible)
              .map(el => {
                const text = norm(el.innerText || el.textContent || el.getAttribute('aria-label'));
                const r = el.getBoundingClientRect();
                return { el, text, r };
              })
              .filter(item => item.text && (item.text === wanted || (allowPartial && (item.text.includes(wanted) || wanted.includes(item.text)))));
            nodes.sort((a, b) =>
              (a.text === wanted ? 0 : 1) - (b.text === wanted ? 0 : 1) ||
              (a.text.length - b.text.length) ||
              (a.r.top - b.r.top)
            );
            if (!nodes.length) return false;
            const el = nodes[0].el;
            el.scrollIntoView({block:'center', inline:'center'});
            const r = el.getBoundingClientRect();
            const opts = {bubbles:true, cancelable:true, view:window, clientX:r.left+r.width/2, clientY:r.top+r.height/2};
            for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
              el.dispatchEvent(new MouseEvent(type, opts));
            }
            return true;
            """,
            option_text,
            allow_partial,
        )
        if not clicked:
            try:
                (active or driver.switch_to.active_element).send_keys(Keys.ENTER)
            except Exception:
                pass
        browser_sleep(0.9)
        browser_wait_for_portal_idle(driver, timeout=10.0)
        state = js_react_control_state(driver, label_hints)
        selected_key = normalize_match_text(state.get("text", ""))
        if target_key and selected_key and (selected_key == target_key or (allow_partial and (target_key in selected_key or selected_key in target_key))):
            return True
        if attempt < 2:
            try:
                js_clear_react_control(driver, label_hints)
            except Exception:
                pass
            browser_sleep(0.3)
    return False


def js_react_control_state(driver, label_hints: list[str]) -> dict[str, Any]:
    return driver.execute_script(
        """
        const hints = arguments[0].map(t => String(t).toLowerCase());
        const norm = s => String(s || '').toLowerCase().replace(/\\.\\.\\./g, '').replace(/\\s+/g, ' ').trim();
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const all = Array.from(document.querySelectorAll('body *')).filter(visible);
        const labels = all
          .filter(el => hints.some(h => norm(el.innerText || el.textContent) === h || norm(el.innerText || el.textContent).startsWith(h + ' *')))
          .map(el => ({ el, r: el.getBoundingClientRect() }))
          .sort((a, b) => a.r.top - b.r.top);
        const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
          .filter(visible)
          .map(el => ({ el, r: el.getBoundingClientRect(), text: norm(el.innerText || el.textContent || el.value), cls: el.className || '', aria: el.getAttribute('aria-disabled') || '' }))
          .filter(x => x.r.width >= 120 && x.r.height >= 25);
        for (const label of labels) {
          const candidates = controls
            .filter(c => c.r.top >= label.r.bottom - 8 && Math.abs(c.r.left - label.r.left) < 100)
            .sort((a, b) => (a.r.top - label.r.bottom) - (b.r.top - label.r.bottom));
          if (candidates.length) {
            const c = candidates[0];
            const disabled = String(c.cls).includes('is-disabled') || c.aria === 'true' || c.text.includes('select institution') && String(c.cls).includes('disabled');
            return { found: true, text: c.text, disabled };
          }
        }
        return { found: false, text: '', disabled: true };
        """,
        label_hints,
    )


def browser_ensure_awc_type(driver) -> None:
    state = js_react_control_state(driver, ["institute type"])
    if state.get("found") and "awc" in state.get("text", "") and not state.get("disabled"):
        return
    if not js_select_react_option(driver, ["institute type"], "AWC", allow_partial=False):
        raise RuntimeError("Institute Type must be AWC before institute selection.")
    browser_sleep(1.0)
    state = js_react_control_state(driver, ["institute type"])
    if "awc" not in state.get("text", ""):
        raise RuntimeError("Institute Type must be AWC before institute selection.")


def browser_require_awc_type_selected(driver) -> None:
    """Live entry keeps Institute Type as a user-confirmed setup step."""
    state = js_react_control_state(driver, ["institute type"])
    text = normalize_match_text(state.get("text", ""))
    if state.get("found") and "awc" in text and not state.get("disabled"):
        return
    visible_text = state.get("text") or "not selected"
    raise RuntimeError(
        "Institute Type must be AWC. On Student Master select Institute Type = AWC once, "
        f"then retry. Current page shows: {visible_text}"
    )


def browser_select_institute(driver, institute_name: str, allow_partial: bool = False) -> None:
    if not institute_name:
        return
    browser_wait_for_portal_idle(driver)
    browser_require_awc_type_selected(driver)
    if browser_institute_dropdown_matches(driver, institute_name):
        return
    institute_state = js_react_control_state(driver, ["institutes"])
    if institute_state.get("disabled"):
        raise RuntimeError("Institutes dropdown is disabled. Select State, District, Block, and wait after Institute Type = AWC.")
    current_text = institute_state.get("text", "")
    if current_text and "select institution" not in current_text and not browser_institute_dropdown_matches(driver, institute_name):
        js_clear_react_control(driver, ["institutes", "institution"])
        browser_sleep(0.4)
    for attempt in range(3):
        if attempt:
            js_clear_react_control(driver, ["institutes", "institution"])
            browser_sleep(0.4)
        selected = js_select_institute_option(driver, institute_name, allow_partial=allow_partial)
        if not selected and attempt == 2:
            selected = js_select_react_option(driver, ["institutes", "institution"], institute_name, allow_partial=allow_partial)
        if not selected:
            if attempt == 2:
                raise RuntimeError(f"Institute selection failed: could not choose {institute_name}")
            continue
        browser_wait_for_portal_idle(driver, timeout=10.0)
        if not browser_institute_dropdown_matches(driver, institute_name):
            browser_sleep(0.5)
            continue
        browser_wait_for_portal_idle(driver, timeout=15.0)
        if browser_institute_dropdown_matches(driver, institute_name):
            return
        browser_sleep(0.5)
    state = js_react_control_state(driver, ["institutes", "institution"])
    visible_text = state.get("text", "") or "Select Institution"
    raise RuntimeError(f"Institute selection failed: expected {institute_name}, page shows {visible_text}. Select State/District/Block and try again.")


def browser_write_institute_diagnostic(driver, path: Path, institute_name: str) -> None:
    try:
        data = driver.execute_script(
            """
            const wanted = arguments[0];
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
            const controls = Array.from(document.querySelectorAll('[role="combobox"], input, div[class*="control"], div[class*="Control"], div[class*="container"]'))
              .filter(visible)
              .map((el, index) => {
                const r = el.getBoundingClientRect();
                return {
                  index,
                  tag: el.tagName,
                  text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label')),
                  cls: String(el.className || ''),
                  ariaDisabled: el.getAttribute('aria-disabled') || '',
                  rect: {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)}
                };
              })
              .filter(item => item.rect.w >= 100 && item.rect.h >= 20);
            const options = Array.from(document.querySelectorAll('[role="option"], div[class*="option"], div[class*="Option"], li, span'))
              .filter(visible)
              .map(el => norm(el.innerText || el.textContent || el.getAttribute('aria-label')))
              .filter(Boolean)
              .slice(0, 200);
            return {
              wanted,
              pageTitle: document.title,
              bodyStart: norm(document.body.innerText || document.body.textContent).slice(0, 2500),
              controls,
              options
            };
            """,
            institute_name,
        )
        path.parent.mkdir(parents=True, exist_ok=True)
        html = [
            "<!doctype html><meta charset='utf-8'><title>Institute Diagnostic</title>",
            "<style>body{font-family:Segoe UI,Arial,sans-serif;padding:20px} pre{white-space:pre-wrap;background:#f3f4f6;padding:12px;border:1px solid #ddd}</style>",
            f"<h1>Institute Diagnostic</h1><p><b>Wanted:</b> {institute_name}</p>",
            "<h2>Controls</h2><pre>" + json.dumps(data.get("controls", []), indent=2, ensure_ascii=False) + "</pre>",
            "<h2>Visible Options</h2><pre>" + json.dumps(data.get("options", []), indent=2, ensure_ascii=False) + "</pre>",
            "<h2>Page Text</h2><pre>" + clean(data.get("bodyStart")) + "</pre>",
        ]
        path.write_text("\n".join(html), encoding="utf-8")
    except Exception:
        pass


def browser_open_add_student(driver, institute_name: str = "", speed_mode: Any = ""):
    browser_wait_for_portal_idle(driver)
    if institute_name and not browser_institute_dropdown_matches(driver, institute_name):
        try:
            browser_select_institute(driver, institute_name)
        except Exception:
            try:
                diag_dir = Path.cwd() / "logs" / "_diagnostics"
                browser_write_institute_diagnostic(driver, diag_dir / f"institute_select_{slugify(institute_name)}.html", institute_name)
            except Exception:
                pass
            raise
    searched = False
    for attempt in range(3):
        browser_wait_for_portal_idle(driver)
        root = browser_find_student_form_root(driver)
        if root:
            return root
        if browser_register_child_screen_is_open(driver):
            root = browser_complete_register_child_step(driver, speed_mode=speed_mode)
            if root:
                return root
            continue
        if browser_click_add_child_button(driver):
            root = browser_complete_register_child_step(driver, speed_mode=speed_mode)
            if root:
                return root
        elif not searched:
            print("Add Child was not visible after institute selection; clicking Search once as fallback.")
            if not browser_click_button(driver, "Search", exact=True):
                js_click_text(driver, ["search"])
            searched = True
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 18.0, 30.0))
            if institute_name:
                browser_wait_for_institute_result(driver, institute_name, timeout=speed_timeout(speed_mode, 3.0, 6.0))
            continue
        else:
            raise RuntimeError("Could not find visible Add Child button after institute selection.")
        if attempt == 0 and institute_name and not browser_child_modal_is_open(driver):
            browser_click_visible_row_text(driver, institute_name)
            browser_sleep(0.5)
    raise RuntimeError("Add Child button clicked, but the real child-entry form did not open.")


def browser_fill_student_form(driver, row: dict[str, Any], root=None, speed_mode: Any = "") -> None:
    browser_wait_for_portal_idle(driver)
    root = root or browser_find_student_form_root(driver)
    if not root:
        raise RuntimeError("Field filling stopped: real Add Child form is not open.")
    payload = {
        "pen": clean(row.get("PEN") or row.get("PEN AWC Number") or row.get("APAAR ID") or row.get("AWC Child ID")),
        "name": clean(row.get("Student Full Name")),
        "dob": excel_date_to_ddmmyyyy(row.get("Birth Date")),
        "age": age_years_months_from_dob(row.get("Birth Date")),
        "gender": canonical_form_gender(row.get("Gender")),
        "mother": clean(row.get("Mother Name")),
        "father": clean(row.get("Father Name")),
        "mobile": only_digits(row.get("Mobile Number"), 10),
        "mobileOwner": clean(row.get("Mobile Owner")) or "Parents",
        "address": clean(row.get("Full Address")),
        "pincode": only_digits(row.get("Pin Code"), 6),
        "blood": clean(row.get("Blood Group")),
        "birthWeight": clean(row.get("Birth Weight (grams)") or row.get("Birth Weight")),
        "birthPlace": clean(row.get("Birth Place")),
    }
    result = driver.execute_script(
        """
        const root = arguments[0];
        const data = arguments[1];
        const visible = el => {
          if (!el) return false;
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const editable = el => visible(el) && !el.disabled && !el.readOnly && el.getAttribute('aria-disabled') !== 'true';
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const setValue = (el, value, clearWhenEmpty = false) => {
          if (!el) return false;
          if ((value == null || String(value) === '') && !clearWhenEmpty) return true;
          if (!editable(el)) return false;
          el.focus();
          const nextValue = value == null ? '' : String(value);
          const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
          if (setter) setter.call(el, nextValue);
          else el.value = nextValue;
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
          el.dispatchEvent(new Event('blur', { bubbles: true }));
          return true;
        };
        const byName = name => root.querySelector(`[name="${CSS.escape(name)}"]`);
        const byPlaceholder = text => Array.from(root.querySelectorAll('input, textarea'))
          .find(el => editable(el) && norm(el.getAttribute('placeholder')).includes(norm(text)));
        const byNearbyText = text => Array.from(root.querySelectorAll('input, textarea, select'))
          .find(el => {
            if (!editable(el)) return false;
            const wrapper = el.closest('label, div, fieldset, section') || el.parentElement;
            return norm(wrapper?.innerText || '').includes(norm(text));
          });
        const setSelect = (name, value) => {
          if (!value) return true;
          const select = byName(name);
          if (!editable(select)) return false;
          const wanted = norm(value);
          const option = Array.from(select.options).find(opt => {
            const optionValue = norm(opt.value);
            const optionText = norm(opt.textContent || opt.label || '');
            return optionValue === wanted || optionText === wanted;
          });
          if (!option) return false;
          select.value = option.value;
          select.dispatchEvent(new Event('input', { bubbles: true }));
          select.dispatchEvent(new Event('change', { bubbles: true }));
          select.dispatchEvent(new Event('blur', { bubbles: true }));
          return true;
        };

        const failures = [];
        const check = (label, ok) => { if (!ok) failures.push(label); };
        setValue(byName('abhaNumber'), '', true);
        setValue(byName('pen_awc_number'), data.pen || '', true);
        check('Child Name', setValue(byName('studentName'), data.name));
        check('Birth Date', setValue(byPlaceholder('dd-mm-yyyy') || byNearbyText('birth date'), data.dob));
        setValue(byName('age'), data.age, true);
        check('Gender', setSelect('gender', data.gender));
        check('Mother Name', setValue(byName('motherName'), data.mother));
        check('Father Name', setValue(byName('fatherName'), data.father));
        check('Mobile Number', setValue(byName('mobile'), data.mobile));
        check('Whose Mobile', setSelect('mobileOwner', data.mobileOwner));
        check('Address', setValue(byName('address'), data.address));
        setValue(byName('pincode'), data.pincode || '', true);
        setSelect('bloodGroup', data.blood);
        check('Birth Weight', setValue(byName('birthWeight'), data.birthWeight));
        setSelect('birthPlace', data.birthPlace);

        return {
          ok: failures.length === 0,
          failures,
          values: {
            studentName: byName('studentName')?.value || '',
            dob: (byPlaceholder('dd-mm-yyyy') || byNearbyText('birth date'))?.value || '',
            age: byName('age')?.value || '',
            gender: byName('gender')?.value || '',
            motherName: byName('motherName')?.value || '',
            fatherName: byName('fatherName')?.value || '',
            mobile: byName('mobile')?.value || '',
            mobileOwner: byName('mobileOwner')?.value || '',
            address: byName('address')?.value || '',
            birthWeight: byName('birthWeight')?.value || ''
          }
        };
        """,
        root,
        payload,
    )
    if payload["dob"] and not browser_type_dob_field(driver, root, payload["dob"]):
        raise RuntimeError(f"Portal did not accept DOB fill: expected {payload['dob']}")
    browser_sleep(speed_wait(speed_mode, 0.25, 0.8))
    if not result.get("ok"):
        raise RuntimeError("Could not fill required field(s): " + ", ".join(result.get("failures", [])))


def canonical_form_gender(value: Any) -> str:
    text = clean(value).lower()
    return GENDER_MAP.get(text, clean(value).upper())


def browser_find_dob_input(driver, root=None):
    root = root or browser_find_student_form_root(driver)
    if not root:
        return None
    return driver.execute_script(
        """
        const root = arguments[0];
        const visible = el => {
          if (!el) return false;
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const fields = Array.from(root.querySelectorAll('input, textarea')).filter(visible);
        return fields.find(el => norm(el.getAttribute('placeholder')).includes('dd-mm-yyyy'))
          || fields.find(el => norm((el.closest('label, div, fieldset, section') || el.parentElement)?.innerText || '').includes('birth date'))
          || null;
        """,
        root,
    )


def browser_type_dob_field(driver, root, dob: str) -> bool:
    if not dob:
        return True
    try:
        from selenium.webdriver.common.keys import Keys
    except Exception:
        Keys = None

    for _attempt in range(2):
        element = browser_find_dob_input(driver, root)
        if not element:
            return False
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'}); arguments[0].focus();",
                element,
            )
            browser_sleep(0.1)
            if Keys is not None:
                element.click()
                element.send_keys(Keys.CONTROL, "a")
                element.send_keys(Keys.BACKSPACE)
                browser_sleep(0.1)
                element.send_keys(dob)
                element.send_keys(Keys.TAB)
            else:
                element.clear()
                element.send_keys(dob)
            browser_sleep(0.25)
        except Exception:
            pass

        try:
            driver.execute_script(
                """
                const el = arguments[0];
                const value = arguments[1];
                if (el.value !== value) {
                  el.focus();
                  const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
                  if (setter) setter.call(el, value);
                  else el.value = value;
                }
                el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
                """,
                element,
                dob,
            )
            browser_sleep(0.25)
            actual = clean(driver.execute_script("return arguments[0].value || '';", element))
            if actual == dob:
                return True
        except Exception:
            pass
    return False


def browser_read_student_form_values(driver, root=None) -> dict[str, str]:
    root = root or browser_find_student_form_root(driver)
    if not root:
        return {}
    return driver.execute_script(
        """
        const root = arguments[0];
        const visible = el => {
          if (!el) return false;
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
        };
        const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
        const byName = name => root.querySelector(`[name="${CSS.escape(name)}"]`);
        const byPlaceholder = text => Array.from(root.querySelectorAll('input, textarea'))
          .find(el => visible(el) && norm(el.getAttribute('placeholder')).toLowerCase().includes(String(text).toLowerCase()));
        const selectValue = name => {
          const el = byName(name);
          if (!el) return { value: '', text: '' };
          const selected = el.selectedOptions && el.selectedOptions.length ? el.selectedOptions[0] : null;
          return { value: norm(el.value), text: norm(selected ? selected.textContent : '') };
        };
        const gender = selectValue('gender');
        return {
          name: norm(byName('studentName')?.value),
          dob: norm((byPlaceholder('dd-mm-yyyy'))?.value),
          genderValue: gender.value,
          genderText: gender.text,
          mother: norm(byName('motherName')?.value),
          father: norm(byName('fatherName')?.value),
          mobile: norm(byName('mobile')?.value),
          address: norm(byName('address')?.value),
          birthWeight: norm(byName('birthWeight')?.value)
        };
        """,
        root,
    )


def browser_verify_student_form_matches_row(driver, row: dict[str, Any], root=None) -> None:
    actual = browser_read_student_form_values(driver, root)
    if not actual:
        raise RuntimeError("Pre-save safety check failed: Add Child form values could not be read.")
    expected = {
        "name": clean(row.get("Student Full Name")),
        "dob": excel_date_to_ddmmyyyy(row.get("Birth Date")),
        "gender": canonical_form_gender(row.get("Gender")),
        "mother": clean(row.get("Mother Name")),
        "father": clean(row.get("Father Name")),
        "mobile": only_digits(row.get("Mobile Number"), 10),
        "address": clean(row.get("Full Address")),
        "birthWeight": clean(row.get("Birth Weight (grams)") or row.get("Birth Weight")),
    }

    mismatches: list[str] = []

    def same_text(left: Any, right: Any) -> bool:
        return normalize_match_text(left) == normalize_match_text(right)

    def same_address(left: Any, right: Any) -> bool:
        left_text = re.sub(r"[()]", "", normalize_match_text(left))
        right_text = re.sub(r"[()]", "", normalize_match_text(right))
        return " ".join(left_text.split()) == " ".join(right_text.split())

    def same_digits(left: Any, right: Any) -> bool:
        return only_digits(left, 20) == only_digits(right, 20)

    checks = [
        ("Child Name", expected["name"], actual.get("name"), same_text),
        ("Birth Date", expected["dob"], actual.get("dob"), same_text),
        ("Mother Name", expected["mother"], actual.get("mother"), same_text),
        ("Father Name", expected["father"], actual.get("father"), same_text),
        ("Mobile Number", expected["mobile"], actual.get("mobile"), same_digits),
        ("Address", expected["address"], actual.get("address"), same_address),
        ("Birth Weight", expected["birthWeight"], actual.get("birthWeight"), same_digits),
    ]
    actual_gender = canonical_form_gender(actual.get("genderValue") or actual.get("genderText"))
    dob_or_gender_mismatch = False
    if expected["gender"] and expected["gender"] != actual_gender:
        dob_or_gender_mismatch = True
        mismatches.append(f"Gender expected '{expected['gender']}' but form has '{actual.get('genderText') or actual.get('genderValue')}'")
    for label, wanted, got, compare in checks:
        if clean(wanted) and not compare(wanted, got):
            if label in {"Birth Date", "Gender"}:
                dob_or_gender_mismatch = True
            mismatches.append(f"{label} expected '{wanted}' but form has '{got or ''}'")
    if mismatches:
        prefix = "Portal did not accept DOB/Gender fill: " if dob_or_gender_mismatch else "Pre-save safety check failed: "
        raise RuntimeError(prefix + "; ".join(mismatches))


def browser_fill_student_form_by_order(driver, row: dict[str, Any], root=None) -> bool:
    root = root or browser_find_student_form_root(driver)
    if not root:
        return False
    payload = {
        "pen": clean(row.get("PEN") or row.get("PEN AWC Number") or row.get("APAAR ID")),
        "name": clean(row.get("Student Full Name")),
        "dob": excel_date_to_iso(row.get("Birth Date")),
        "gender": clean(row.get("Gender")),
        "mother": clean(row.get("Mother Name")),
        "father": clean(row.get("Father Name")),
        "mobile": only_digits(row.get("Mobile Number"), 10),
        "mobileOwner": clean(row.get("Mobile Owner")),
        "address": clean(row.get("Full Address")),
        "pincode": only_digits(row.get("Pin Code"), 6),
        "blood": clean(row.get("Blood Group")),
        "birthWeight": clean(row.get("Birth Weight (grams)") or row.get("Birth Weight")),
    }
    script = """
    const data = arguments[0];
    const root = arguments[1];
    const visible = el => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && !el.disabled && !el.readOnly;
    };
    const setValue = (el, value) => {
      if (!el || value == null || String(value) === '') return false;
      el.focus();
      const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
      if (setter) setter.call(el, String(value));
      else el.value = String(value);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.blur();
      return true;
    };
    const inputs = Array.from(root.querySelectorAll('input, textarea')).filter(visible);
    const selects = Array.from(root.querySelectorAll('select')).filter(visible);
    const textInputs = inputs.filter(el => (el.type || '').toLowerCase() !== 'hidden');
    const trySetByName = (needles, value) => {
      const el = [...inputs, ...selects].find(x => needles.some(n => `${x.name||''} ${x.id||''} ${x.placeholder||''}`.toLowerCase().includes(n)));
      return setValue(el, value);
    };

    trySetByName(['pen', 'apaar'], data.pen);
    const ordered = [
      data.name,
      data.dob,
      data.mother,
      data.father,
      data.mobile,
      data.address,
      data.pincode,
      data.birthWeight
    ].filter(v => v != null && String(v) !== '');
    let idx = 0;
    for (const value of ordered) {
      while (idx < textInputs.length && !setValue(textInputs[idx], value)) idx++;
      idx++;
    }

    const clickOption = (text) => {
      if (!text) return false;
      const target = String(text).toLowerCase();
      const nodes = Array.from(document.querySelectorAll('[role="option"], div, li, span')).filter(visible);
      const match = nodes.find(n => (n.innerText || n.textContent || '').trim().toLowerCase() === target);
      if (match) { match.click(); return true; }
      return false;
    };
    const clickComboNear = (needle, value) => {
      const containers = Array.from(root.querySelectorAll('label, div')).filter(visible);
      const container = containers.find(c => (c.innerText || '').toLowerCase().includes(needle));
      const target = container?.querySelector('[role="combobox"], input, [class*="control"]');
      if (!target) return false;
      target.click();
      setTimeout(() => clickOption(value), 250);
      return true;
    };
    clickComboNear('gender', data.gender);
    clickComboNear('whose mobile', data.mobileOwner);
    clickComboNear('blood', data.blood);
    return textInputs.length;
    """
    count = driver.execute_script(script, payload, root)
    browser_sleep(0.8)
    return bool(count)


def browser_save_student(driver, root=None, speed_mode: Any = "") -> None:
    browser_wait_for_portal_idle(driver)
    root = root or browser_find_student_form_root(driver)
    if not root:
        raise RuntimeError("Save stopped: real Add Child form is not open.")
    labels = ["Save Child", "Save Student", "Save", "Submit", "Create", "Add"]
    clicked = browser_click_button_in_root(driver, root, labels, exact=True)
    if not clicked:
        clicked = browser_click_button_in_root(driver, root, labels, exact=False)
    if not clicked:
        clicked = browser_click_visible_save_button(driver)
    if not clicked:
        if browser_find_visible_save_button(driver):
            raise RuntimeError("Save button is visible, but the browser could not click it. Try scrolling the form bottom into view and retry.")
        raise RuntimeError("Could not find a visible Save/Submit button on the Add Child form.")
    try:
        browser_wait_for_save_result(driver, previous_root=root, timeout=speed_timeout(speed_mode, 7.0, 10.0), speed_mode=speed_mode)
    except RuntimeError as exc:
        if "form stayed open" not in str(exc).lower() or not browser_submit_modal_form(driver):
            raise
        browser_wait_for_save_result(driver, previous_root=root, timeout=speed_timeout(speed_mode, 7.0, 10.0), speed_mode=speed_mode)
    browser_dismiss_success_notice(driver)


def browser_cancel_student_form_if_open(driver) -> None:
    root = browser_find_student_form_root(driver)
    if not root:
        return
    if browser_click_button_in_root(driver, root, ["Cancel", "Close"], exact=True):
        browser_wait_for_portal_idle(driver, timeout=12.0)
        browser_sleep(0.5)


def row_first_value(row: dict[str, Any], names: list[str]) -> str:
    normalized = {normalize_match_text(key): key for key in row.keys()}
    for name in names:
        key = normalized.get(normalize_match_text(name))
        if key is not None and clean(row.get(key)):
            return clean(row.get(key))
    return ""


def screening_student_name(row: dict[str, Any]) -> str:
    return row_first_value(row, ["Student Full Name", "BENEFICIARY NAME", "Child Name", "Student Name", "Name"])


def screening_institute_name(row: dict[str, Any]) -> str:
    return row_first_value(row, ["Institute Name", "School Name", "AWC NAME", "Institution Name", "Institute"])


def screening_mobile(row: dict[str, Any], default: str = "9999999999") -> str:
    custom = clean(row.get("_school_custom_mobile"))
    if custom:
        c_digits = "".join(ch for ch in custom if ch.isdigit())
        if len(c_digits) >= 10:
            return c_digits[-10:]
        if len(c_digits) > 0:
            return c_digits.ljust(10, "9")
    raw = row_first_value(
        row,
        [
            "Mobile Number",
            "Mobile",
            "MOBILE",
            "MOBILE NO",
            "Mobile No",
            "Whose Mobile",
            "Contact",
            "Phone",
            "Phone Number",
            "Mobile_Number",
            "MobileNo",
            "Student Mobile",
            "Parent Mobile",
            "Guardian Mobile",
        ],
    )
    digits = "".join(ch for ch in clean(raw) if ch.isdigit())
    if len(digits) >= 10:
        return digits[-10:]
    if len(digits) > 0:
        return digits.ljust(10, "9")
    return default


def ask_school_mobile_console(school_name: str) -> str:
    try:
        val = safe_input(f"Enter 10-digit Mobile Number for school '{school_name}' (press Enter for default 9999999999): ").strip()
        val_digits = "".join(ch for ch in val if ch.isdigit())
        if len(val_digits) == 10:
            return val_digits
    except Exception:
        pass
    return "9999999999"


def student_age_years(row: dict[str, Any]) -> float | None:
    dob_raw = row_first_value(row, ["Birth Date", "DOB", "Birth Date (DD-MM-YYYY)", "Date of Birth", "BirthDate"])
    iso = excel_date_to_iso(dob_raw)
    if not iso:
        age_text = row_first_value(row, ["ageText", "Age", "Student Age"])
        years_match = re.search(r"(\d+)\s*years?(?:\s*(\d+)\s*months?)?", age_text, flags=re.I)
        if years_match:
            years = int(years_match.group(1) or 0)
            months = int(years_match.group(2) or 0)
            return years + (months / 12.0)
        months_match = re.search(r"(\d+)\s*months?", age_text, flags=re.I)
        if months_match:
            return int(months_match.group(1) or 0) / 12.0
        return None
    try:
        dob = datetime.strptime(iso, "%Y-%m-%d")
        today = datetime.today()
        days = (today - dob).days
        if days < 0:
            return 0.0
        return days / 365.25
    except Exception:
        return None


def compute_age_weight_height(row: dict[str, Any]) -> tuple[str, str, str, str]:
    """Compute healthy WHO reference values for children with smooth linear interpolation for exact ages (e.g. 3 years 3 months).
    Returns (weight_kg, height_cm, head_circumference_cm, muac_cm).
    Reference: WHO Child Growth Standards Reference Chart.
    """
    age_yrs = student_age_years(row)
    if age_yrs is None:
        return "18", "110", "49", "15"

    points = [
        (0.0, 3.4, 50.0, 34.5, None),   # Birth
        (0.5, 7.5, 67.0, 42.5, 13.25),  # 6 Months
        (1.0, 10.0, 76.0, 46.0, 14.0),  # 1 Year
        (2.0, 12.5, 87.0, 48.0, 14.75), # 2 Years
        (3.0, 14.5, 95.0, 49.0, 15.25), # 3 Years
        (4.0, 17.0, 103.0, 50.0, 15.5), # 4 Years
        (5.0, 20.0, 110.5, 51.0, 16.0), # 5 Years
        (6.0, 21.0, 116.0, 51.0, 16.0),
        (12.0, 33.0, 140.0, None, None),
    ]

    if age_yrs <= points[0][0]:
        p = points[0]
        return f"{p[1]:.1f}", f"{int(round(p[2]))}", f"{p[3]:.1f}", f"{p[4]:.1f}" if p[4] else ""

    for i in range(len(points) - 1):
        a0, w0, h0, hc0, m0 = points[i]
        a1, w1, h1, hc1, m1 = points[i + 1]
        if a0 <= age_yrs <= a1:
            t = (age_yrs - a0) / (a1 - a0)
            wt = w0 + t * (w1 - w0)
            ht = h0 + t * (h1 - h0)

            hc_str = ""
            if hc0 is not None and hc1 is not None:
                hc_str = f"{(hc0 + t * (hc1 - hc0)):.1f}"
            elif hc0 is not None:
                hc_str = f"{hc0:.1f}"

            muac_str = ""
            if m0 is not None and m1 is not None:
                muac_str = f"{(m0 + t * (m1 - m0)):.1f}"
            elif m1 is not None and age_yrs >= 0.5:
                muac_str = f"{m1:.1f}"

            return f"{wt:.1f}", f"{int(round(ht))}", hc_str, muac_str

    last = points[-1]
    return f"{last[1]:.1f}", f"{int(round(last[2]))}", f"{last[3]:.1f}" if last[3] else "", f"{last[4]:.1f}" if last[4] else ""


def screening_health_payload(row: dict[str, Any]) -> dict[str, str]:
    calc_wt, calc_ht, calc_hc, calc_muac = compute_age_weight_height(row)
    weight = row_first_value(row, ["Weight", "Weight (kg)", "Weight Kg", "WT", "Weight(kg)"]) or calc_wt
    height = row_first_value(row, ["Height", "Height (cm)", "Height Cm", "Length", "Height/Length", "HT", "Height(cm)"]) or calc_ht
    hc_val = row_first_value(row, ["Head Circumference", "Head Circumference (cm)", "Head Circumference(cm)", "HC"]) or calc_hc
    muac_val = row_first_value(row, ["MUAC", "MUIC", "MUAC (cm)", "MUAC(cm)", "Mid Upper Arm Circumference"]) or calc_muac
    bp = row_first_value(row, ["Blood Pressure", "BP", "Blood Pressure (mmHg)", "BP Reading"]) or "100/60"
    v_left = row_first_value(row, ["Vision Left", "Vision - Left Eye", "Vision Left Eye", "Left Eye"]) or "6/6"
    v_right = row_first_value(row, ["Vision Right", "Vision - Right Eye", "Vision Right Eye", "Right Eye"]) or "6/6"
    hb = row_first_value(row, ["HB Count", "Hb Count", "Hemoglobin", "HB"])
    return {
        "weight": str(weight),
        "height": str(height),
        "headCircumference": str(hc_val),
        "muac": str(muac_val),
        "bloodPressure": str(bp),
        "visionLeft": str(v_left),
        "visionRight": str(v_right),
        "hbCount": clean(hb),
        "mobile": "" if row.get("_is_awc") else screening_mobile(row),
    }


def validate_screening_health_rows(rows: list[dict[str, Any]]) -> None:
    if not rows:
        raise ImportErrorWithHelp("No screening rows found.")
    problems: list[str] = []
    for index, row in enumerate(rows[:25], start=1):
        missing: list[str] = []
        if not screening_student_name(row):
            missing.append("student name")
        if not screening_institute_name(row):
            missing.append("school/institute name")
        health = screening_health_payload(row)
        if not health.get("weight"):
            missing.append("weight")
        if not health.get("height"):
            missing.append("height")
        if missing:
            sr_no = clean(row.get("Sr No")) or str(index)
            problems.append(f"Sr No {sr_no}: missing {', '.join(missing)}")
    if problems:
        raise ImportErrorWithHelp("Screening health Excel is missing required data:\n" + "\n".join(problems[:10]))


def browser_open_screening_page(driver, speed_mode: Any = "") -> None:
    opened = False
    for url in [f"{BASE_URL}/RBSK/childScreening", f"{BASE_URL}/RBSK/screening", f"{BASE_URL}/RBSK/Screening", f"{BASE_URL}/RBSK/#/screening"]:
        try:
            driver.get(url)
            browser_sleep(speed_wait(speed_mode, 0.7, 1.2))
            text = clean(driver.execute_script("return document.body.innerText || document.body.textContent || '';")).lower()
            if "screening" in text and ("month" in text or "financial year" in text):
                opened = True
                break
        except Exception:
            pass
    if opened:
        return
    clicked = driver.execute_script(
        """
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const links = Array.from(document.querySelectorAll('a, button, [role="button"], li, div, span')).filter(visible);
        const byText = links
          .map(el => ({el: el.closest('a, button, [role="button"], li') || el, text: norm(el.innerText || el.textContent || el.getAttribute('aria-label')), r: el.getBoundingClientRect()}))
          .filter(x => x.text.includes('screening'))
          .sort((a, b) => a.text.length - b.text.length);
        if (byText.length) {
          byText[0].el.click();
          return true;
        }
        const side = Array.from(document.querySelectorAll('aside a, aside button, nav a, nav button, .sidebar a, .sidebar button, [class*="side"] a, [class*="side"] button')).filter(visible);
        const fifth = side[4];
        if (fifth) {
          fifth.click();
          return true;
        }
        return false;
        """
    )
    if not clicked:
        raise RuntimeError("Could not open Screening from the fifth menu option.")
    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 14.0))


def show_gui_message(title: str, message: str) -> None:
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title(title)
        root.configure(bg="#f4f7f9")
        root.geometry("500x300+20+20")
        root.minsize(440, 260)
        root.resizable(True, True)
        root.attributes("-topmost", True)

        header = tk.Frame(root, bg="#e5e7eb", height=42)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=title, bg="#e5e7eb", fg="#374151", font=("Segoe UI", 11, "bold")).pack(
            side="left", padx=14, pady=10
        )

        # Reserve the footer before the expanding body so action buttons cannot
        # be pushed below the window when instruction text wraps.
        footer = tk.Frame(root, bg="#f4f7f9")
        footer.pack(side="bottom", fill="x", padx=12, pady=(4, 10))
        tk.Button(footer, text="Minimize", command=root.iconify, width=10).pack(side="left")
        ok_button = tk.Button(footer, text="OK", command=root.destroy, width=10, default="active")
        ok_button.pack(side="right")

        body = tk.Frame(root, bg="#ffffff")
        body.pack(side="top", fill="both", expand=True, padx=12, pady=(10, 4))
        tk.Label(body, text="i", bg="#168bd2", fg="#ffffff", font=("Segoe UI", 18, "bold"), width=2).pack(
            side="left", anchor="n", padx=(2, 12), pady=4
        )
        tk.Label(
            body,
            text=message,
            bg="#ffffff",
            fg="#111827",
            font=("Segoe UI", 9),
            justify="left",
            anchor="nw",
            wraplength=385,
        ).pack(side="left", fill="both", expand=True, pady=4)

        root.bind("<Return>", lambda _event: root.destroy())
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        ok_button.focus_set()
        root.mainloop()
    except Exception:
        print(f"\n{title}: {message}")


def browser_is_logged_out(driver) -> bool:
    try:
        url = str(driver.current_url or "").lower()
        if "login" in url and "student" not in url and "screening" not in url:
            return True
        page_text = clean(driver.execute_script("return document.body.innerText || document.body.textContent || '';")).lower()
        if "captcha" in page_text and ("login id" in page_text or "password" in page_text or "sign in" in page_text):
            return True
        return False
    except Exception:
        return False


def browser_is_open(driver) -> bool:
    if driver is None:
        return False
    try:
        return bool(driver.window_handles)
    except Exception:
        return False


def browser_prompt_school_mobile(driver, school_name: str) -> str:
    if not browser_is_open(driver):
        return ""
    driver.execute_script(
        r"""
        const schoolName = String(arguments[0] || 'Current School/College');
        document.getElementById('rbsk-automation-mobile-prompt')?.remove();
        window.__rbskMobilePromptResult = {status: 'pending', value: ''};

        const host = document.createElement('div');
        host.id = 'rbsk-automation-mobile-prompt';
        host.style.cssText = 'position:fixed;left:18px;top:18px;width:min(380px,calc(100vw - 36px));z-index:2147483647;';
        const shadow = host.attachShadow({mode: 'open'});
        shadow.innerHTML = `
          <style>
            * { box-sizing: border-box; }
            .card { overflow:hidden; border:1px solid #0891b2; border-radius:12px; background:#f8fafc;
              box-shadow:0 18px 45px rgba(15,23,42,.32); color:#172033; font:15px/1.4 "Segoe UI",sans-serif; }
            .head { background:#1697aa; color:#fff; padding:11px 16px; font-weight:700; font-size:17px; }
            .body { padding:14px 16px 10px; }
            .mode { margin:0 0 12px; color:#334155; }
            label { display:block; margin-bottom:6px; font-weight:600; }
            .school { margin-bottom:10px; color:#0f172a; font-weight:700; }
            input { width:100%; height:40px; border:1px solid #94a3b8; border-radius:6px; padding:7px 10px;
              background:#fff; color:#0f172a; font-size:17px; outline:none; }
            input:focus { border-color:#0891b2; box-shadow:0 0 0 3px rgba(8,145,178,.16); }
            .error { min-height:20px; padding-top:4px; color:#dc2626; font-size:13px; }
            .actions { display:flex; justify-content:flex-end; gap:10px; padding:0 16px 14px; }
            button { min-width:92px; border:0; border-radius:6px; padding:9px 14px; font-weight:700; cursor:pointer; }
            .cancel { background:#e2e8f0; color:#334155; }
            .save { background:#0f9b8e; color:#fff; }
          </style>
          <section class="card" role="dialog" aria-modal="false" aria-labelledby="rbsk-mobile-title">
            <div class="head" id="rbsk-mobile-title">School Mobile Number</div>
            <div class="body">
              <p class="mode">Direct Portal Mobile Mode is active.</p>
              <label for="mobile">Enter 10-digit Mobile Number for:</label>
              <div class="school"></div>
              <input id="mobile" type="tel" inputmode="numeric" maxlength="10" autocomplete="tel" aria-describedby="error">
              <div class="error" id="error"></div>
            </div>
            <div class="actions">
              <button class="cancel" type="button">Cancel</button>
              <button class="save" type="button">Save &amp; Continue</button>
            </div>
          </section>`;
        shadow.querySelector('.school').textContent = schoolName;
        const input = shadow.querySelector('#mobile');
        const error = shadow.querySelector('#error');
        input.addEventListener('input', () => {
          input.value = input.value.replace(/\D/g, '').slice(0, 10);
          error.textContent = '';
        });
        shadow.querySelector('.save').addEventListener('click', () => {
          const digits = input.value.replace(/\D/g, '');
          if (digits.length !== 10) {
            error.textContent = 'Enter exactly 10 digits.';
            input.focus();
            return;
          }
          window.__rbskMobilePromptResult = {status: 'saved', value: digits};
          host.remove();
        });
        shadow.querySelector('.cancel').addEventListener('click', () => {
          window.__rbskMobilePromptResult = {status: 'cancelled', value: ''};
          host.remove();
        });
        document.body.appendChild(host);
        input.focus();
        """,
        school_name,
    )
    while browser_is_open(driver):
        try:
            result = driver.execute_script("return window.__rbskMobilePromptResult || {status:'pending',value:''};") or {}
            status = clean(result.get("status")) if isinstance(result, dict) else ""
            if status == "saved":
                return "".join(ch for ch in clean(result.get("value")) if ch.isdigit())[-10:]
            if status == "cancelled":
                return ""
        except Exception:
            if not browser_is_open(driver):
                return ""
        browser_sleep(0.15)
    return ""


def browser_check_and_wait_if_logged_out(driver) -> None:
    if not browser_is_logged_out(driver):
        return
    print("\nSession appears to be logged out. Waiting for re-login in browser...")
    show_gui_message(
        "Logged Out",
        "The RBSK browser session appears to be logged out or expired.\n\n"
        "1. Enter Login ID, Password, and CAPTCHA in the browser.\n"
        "2. Log in and navigate back to the Screening page.\n"
        "3. Click OK here to resume.",
    )
    for _ in range(60):
        if not browser_is_logged_out(driver):
            print("Session restored.")
            return
        time.sleep(2)


def browser_prepare_screening_search(
    driver,
    month: str = "",
    institution_type: str = "School",
    speed_mode: Any = "",
    page_already_prepared: bool = False,
) -> None:
    if not page_already_prepared:
        browser_open_screening_page(driver, speed_mode=speed_mode)
    browser_handle_plan_completed_modal(driver)
    if institution_type:
        js_select_react_option(driver, ["institute type", "institution type"], institution_type, allow_partial=True)
        browser_sleep(speed_wait(speed_mode, 0.25, 0.6))
    if month:
        js_select_react_option(driver, ["month"], month, allow_partial=True)
        browser_sleep(speed_wait(speed_mode, 0.25, 0.6))
    if not month:
        if page_already_prepared:
            if not browser_click_screening_search(driver):
                raise RuntimeError("Could not click Search on the prepared Screening page.")
            browser_wait_for_screening_school_rows(driver, timeout=speed_timeout(speed_mode, 12.0, 30.0))
            return
        print("\nScreening month step:")
        print("Select the Month manually in the RBSK Screening page, then click Search if you want.")
        show_gui_message(
            "Select Month Manually",
            "Please select the Month manually in the browser window, click Search, and click OK here when the school list is visible."
        )
        try:
            browser_wait_for_screening_school_rows(driver, timeout=speed_timeout(speed_mode, 5.0, 10.0))
        except Exception:
            pass
        return
    if not browser_click_screening_search(driver):
        print("Search button was not clicked automatically. If needed, click Search manually.")
        show_gui_message(
            "Click Search Manually",
            "Search button was not clicked automatically. Please click Search in the browser, and click OK here when the school list is visible."
        )
    browser_wait_for_screening_school_rows(driver, timeout=speed_timeout(speed_mode, 12.0, 30.0))


def browser_click_screening_search(driver) -> bool:
    return bool(
        driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], [role="button"]'))
              .filter(visible)
              .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'))}))
              .filter(x => x.text === 'search' || x.text.includes(' search'));
            if (!buttons.length) return false;
            buttons[0].el.click();
            return true;
            """
        )
    )


def browser_selected_screening_month(driver) -> str:
    try:
        return clean(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const selects = Array.from(document.querySelectorAll('select')).filter(visible);
                const monthSelect = selects.find(el => {
                  const id = el.id || '';
                  const label = id ? document.querySelector(`label[for="${CSS.escape(id)}"]`) : null;
                  const nearby = label?.innerText || el.closest('div, td, fieldset')?.innerText || '';
                  return norm(`${el.name || ''} ${id} ${nearby}`).includes('month');
                });
                if (!monthSelect) return '';
                const option = monthSelect.selectedOptions && monthSelect.selectedOptions[0];
                const text = String(option?.textContent || option?.value || monthSelect.value || '').trim();
                return /select/i.test(text) ? '' : text;
                """
            )
        )
    except Exception:
        return ""


def browser_screening_school_list_visible(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                return Array.from(document.querySelectorAll('tbody tr, [role="row"]'))
                  .filter(visible)
                  .some(row => {
                    const cells = Array.from(row.querySelectorAll('td, [role="cell"]')).filter(visible);
                    const text = norm(row.innerText || row.textContent);
                    return cells.length >= 8 && /visit closed|screening in\\s*progress|visit planned|visit completed/.test(text);
                  });
                """
            )
        )
    except Exception:
        return False


def browser_screening_student_list_has_name(driver, student_name: str) -> bool:
    if not student_name:
        return False
    try:
        return bool(
            driver.execute_script(
                """
                const wanted = String(arguments[0] || '').toLowerCase().replace(/[()\\[\\]{}.,_-]+/g, ' ').replace(/\\s+/g, ' ').trim();
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/[()\\[\\]{}.,_-]+/g, ' ').replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent || '');
                if (pageText.includes('screening preview') || (pageText.includes('confirm') && pageText.includes('submit'))) {
                  return false;
                }
                if (pageText.includes('screening for') || pageText.includes('child details') || (pageText.includes('weight') && pageText.includes('height'))) {
                  return false;
                }
                return Array.from(document.querySelectorAll('tbody tr, [role="row"]'))
                  .filter(visible)
                  .some(row => row.querySelectorAll('td, [role="cell"]').length >= 6 && norm(row.innerText || row.textContent).includes(wanted));
                """,
                student_name,
            )
        )
    except Exception:
        return False


def browser_screening_student_list_visible(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent || '');
                if (pageText.includes('screening preview') || (pageText.includes('confirm') && pageText.includes('submit'))) {
                  return false;
                }
                if (pageText.includes('screening for:') || pageText.includes('child details') || (pageText.includes('weight') && pageText.includes('height'))) {
                  return false;
                }
                const headerText = Array.from(document.querySelectorAll('thead th, th, [role="columnheader"]'))
                  .filter(visible)
                  .map(el => norm(el.innerText || el.textContent))
                  .join(' ');
                const studentHeaders = headerText.includes('name')
                  && headerText.includes('gender')
                  && (headerText.includes('dob') || headerText.includes('birth date'))
                  && headerText.includes('age')
                  && (headerText.includes('health status') || headerText.includes('screening actions'));
                if (studentHeaders) return true;
                if (/visit closed|screening in\\s*progress|visit planned|visit completed/.test(pageText) && !/start screening|health|doctor/.test(pageText)) {
                  return false;
                }
                return Array.from(document.querySelectorAll('tbody tr, tr, [role="row"]'))
                  .filter(visible)
                  .some(row => {
                    const cells = Array.from(row.querySelectorAll('td, th, [role="cell"]')).filter(visible);
                    if (cells.length < 3) return false;
                    const text = norm(row.innerText || row.textContent);
                    return /start screening|health|doctor|stethoscope|medical/.test(text)
                      || (/years?|months?/.test(text) && !/visit closed|visit planned|visit completed/.test(text));
                  });
                """
            )
        )
    except Exception:
        return False


def browser_is_screening_health_form_visible(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent || '');
                const hasFormText = text.includes('screening for:') || text.includes('child details') || text.includes('location details') || (text.includes('weight') && text.includes('height'));
                const hasFormInputs = document.querySelectorAll('input, select').length >= 3;
                return hasFormText && hasFormInputs;
                """
            )
        )
    except Exception:
        return False


def browser_screening_preview_visible(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent || '');
                if (pageText.includes('screening preview')) return true;
                return Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a, [role="button"]'))
                  .filter(visible)
                  .some(el => {
                    const text = norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '');
                    return (text.includes('confirm') && text.includes('submit')) || text.includes('back to edit');
                  });
                """
            )
        )
    except Exception:
        return False


def browser_wait_for_screening_school_rows(driver, timeout: float = 30.0) -> None:
    ready = browser_wait_for_script(
        driver,
        """
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        return Array.from(document.querySelectorAll('tbody tr, [role="row"]'))
          .filter(visible)
          .some(row => row.querySelectorAll('td, [role="cell"]').length >= 4);
        """,
        timeout=timeout,
        interval=0.5,
    )
    if not ready:
        raise RuntimeError("The Screening school list did not load after Search. Confirm the Month and click Search again.")


def _browser_handle_mark_attendance_modal_legacy(driver) -> bool:
    try:
        handled = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const bodyText = norm(document.body.innerText || document.body.textContent || '');
                if (!/mark attendance|present to start|attendance/i.test(bodyText)) {
                  return false;
                }
                const dialogs = Array.from(document.querySelectorAll('div, form, [role="dialog"], .modal, [class*="modal"], [class*="Dialog"]'))
                  .filter(visible)
                  .filter(el => /mark attendance|present to start|attendance/i.test(norm(el.innerText || el.textContent || '')));
                const scope = dialogs[0] || document;
                
                // 1. Click inputs if present
                const inputs = Array.from(scope.querySelectorAll('input')).filter(visible);
                for (const input of inputs) {
                  if (!input.checked) {
                    input.click();
                  }
                }

                // Find all card containers inside the modal
                const allDivs = Array.from(scope.querySelectorAll('div, label, li')).filter(visible);
                const memberCards = [];
                for (const el of allDivs) {
                  const text = norm(el.innerText || el.textContent || '');
                  if (/doctor|pharmacist|anm|nurse|medical|female|male/i.test(text) && text.length < 80) {
                    let card = el;
                    while (card && card.parentElement && card.parentElement !== scope) {
                      const r = card.getBoundingClientRect();
                      if (r.height >= 40 && r.height <= 120 && r.width > 120) {
                        break;
                      }
                      card = card.parentElement;
                    }
                    if (card && !memberCards.includes(card)) {
                      memberCards.push(card);
                    }
                  }
                }

                for (const card of memberCards) {
                  let isSelected = false;
                  
                  const cardClass = String(card.className || '').toLowerCase();
                  if (cardClass.includes('selected') || cardClass.includes('checked') || cardClass.includes('active') || card.getAttribute('aria-checked') === 'true') {
                    isSelected = true;
                  }
                  
                  const innerHTML = card.innerHTML.toLowerCase();
                  if (innerHTML.includes('check') || innerHTML.includes('select') || innerHTML.includes('success')) {
                    const hasCheck = Array.from(card.querySelectorAll('svg, i, span')).some(child => {
                      const childClass = String(child.className || '').toLowerCase();
                      const childIcon = String(child.getAttribute('data-icon') || '').toLowerCase();
                      const childText = norm(child.innerText || child.textContent || '');
                      return childClass.includes('check') || childIcon.includes('check') || childText === 'check';
                    });
                    if (hasCheck) {
                      isSelected = true;
                    }
                  }
                  
                  const style = getComputedStyle(card);
                  const bg = style.backgroundColor;
                  const border = style.borderColor;
                  if (bg.includes('rgba(0, 168') || bg.includes('rgba(0, 150') || border.includes('rgb(0, 168') || border.includes('rgb(0, 150') || border.includes('rgb(26, 188')) {
                    isSelected = true;
                  }

                  const input = card.querySelector('input');
                  if (input && input.checked) {
                    isSelected = true;
                  }

                  if (!isSelected) {
                    card.click();
                  }
                }

                // 3. Click Submit & Start button
                const buttons = Array.from(scope.querySelectorAll('button, a, [role="button"], input[type="button"], input[type="submit"]')).filter(visible);
                const submitBtn = buttons.find(el => /submit\\s*&\\s*start|submit|start/i.test(norm(el.innerText || el.textContent || el.value || '')));
                if (submitBtn) {
                  submitBtn.scrollIntoView({block: 'center', inline: 'center'});
                  submitBtn.click();
                  return true;
                }
                return false;
                """
            )
        )
        if handled:
            time.sleep(1.0)
        return handled
    except Exception:
        return False


def browser_handle_mark_attendance_modal(driver) -> bool:
    """Select every available attendance member, verify all, then start screening."""
    try:
        script = """
        const action = arguments[0];
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden'
            && s.display !== 'none' && s.opacity !== '0';
        };
        const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        const candidates = Array.from(document.querySelectorAll(
          '[role="dialog"], .modal, .modal-content, [class*="modal"], [class*="Modal"], '
            + '[class*="dialog"], [class*="Dialog"], form, div'
        )).filter(visible).filter(el => {
          const text = norm(el.innerText || el.textContent || '');
          return text.includes('mark attendance to start screening')
            && text.includes('check the members who are present');
        });
        candidates.sort((a, b) => {
          const ar = a.getBoundingClientRect();
          const br = b.getBoundingClientRect();
          return (ar.width * ar.height) - (br.width * br.height);
        });
        const scope = candidates[0];
        if (!scope) return {found: false, total: 0, selected: 0, submitted: false};

        const inputs = Array.from(scope.querySelectorAll('input[type="checkbox"], input[type="radio"]'))
          .filter(el => !el.disabled && el.getAttribute('aria-disabled') !== 'true');
        const customChecks = Array.from(scope.querySelectorAll('[role="checkbox"]'))
          .filter(visible)
          .filter(el => !el.disabled && el.getAttribute('aria-disabled') !== 'true');
        if (action === 'select') {
          for (const input of inputs) {
            if (!input.checked) input.click();
          }
          for (const control of customChecks) {
            if (control.getAttribute('aria-checked') !== 'true') control.click();
          }
        }

        const total = inputs.length + customChecks.length;
        const selected = inputs.filter(el => el.checked).length
          + customChecks.filter(el => el.getAttribute('aria-checked') === 'true').length;
        if (action !== 'submit' || total === 0 || selected < total) {
          return {found: true, total, selected, submitted: false};
        }

        const buttons = Array.from(scope.querySelectorAll(
          'button, a, [role="button"], input[type="button"], input[type="submit"]'
        )).filter(visible);
        const submitBtn = buttons.find(el => {
          const text = norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || '');
          return text === 'submit & start' || text === 'submit and start';
        });
        if (!submitBtn || submitBtn.disabled || submitBtn.getAttribute('aria-disabled') === 'true') {
          return {found: true, total, selected, submitted: false};
        }
        submitBtn.scrollIntoView({block: 'center', inline: 'center'});
        submitBtn.click();
        return {found: true, total, selected, submitted: true};
        """

        state = driver.execute_script(script, "select") or {}
        if not state.get("found"):
            return False
        for _attempt in range(3):
            browser_sleep(0.1)
            state = driver.execute_script(script, "select") or {}
            total = int(state.get("total") or 0)
            selected = int(state.get("selected") or 0)
            if total > 0 and selected >= total:
                break
        result = driver.execute_script(script, "submit") or {}
        handled = bool(result.get("submitted"))
        if handled:
            print(
                f"Mark Attendance: selected all {result.get('selected')}/{result.get('total')} "
                "members and clicked Submit & Start."
            )
            browser_sleep(0.2)
        return handled
    except Exception:
        return False


def browser_handle_plan_completed_modal(driver) -> bool:
    """Click the exact Plan Completed prompt and keep an immediate page watcher active."""
    try:
        handled = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden'
                    && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const tryProceed = () => {
                  const dialogs = Array.from(document.querySelectorAll(
                    '[role="dialog"], .modal, .modal-content, .swal2-popup, .swal-modal, '
                      + '[class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"], '
                      + '[class*="popup"], [class*="Popup"]'
                  )).filter(visible);
                  const dialog = dialogs.find(el => {
                    const text = norm(el.innerText || el.textContent || '');
                    const plannedGroupComplete = text.includes('planned children for this institution have been screened')
                      || text.includes('planned students for this institution have been screened');
                    return text.includes('plan completed')
                      && plannedGroupComplete
                      && text.includes('do you want to screen more');
                  });
                  if (!dialog) return false;

                  const buttons = Array.from(dialog.querySelectorAll(
                    'button, [role="button"], input[type="button"], input[type="submit"], a'
                  )).filter(visible);
                  const proceed = buttons.find(el => {
                    const text = norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || '');
                    return text === 'yes, proceed!' || text === 'yes, proceed' || text === 'yes proceed';
                  });
                  if (!proceed || proceed.disabled || proceed.getAttribute('aria-disabled') === 'true') return false;
                  proceed.scrollIntoView({block: 'center', inline: 'center'});
                  proceed.click();
                  window.__rbskPlanCompletedClickedAt = Date.now();
                  return true;
                };

                const handledNow = tryProceed();
                if (!window.__rbskPlanCompletedWatcherInstalled && document.body) {
                  let checkQueued = false;
                  const observer = new MutationObserver(() => {
                    if (checkQueued) return;
                    checkQueued = true;
                    requestAnimationFrame(() => {
                      checkQueued = false;
                      tryProceed();
                    });
                  });
                  observer.observe(document.body, {childList: true, subtree: true, characterData: true});
                  window.__rbskPlanCompletedWatcherInstalled = true;
                  window.__rbskPlanCompletedObserver = observer;
                }
                return handledNow;
                """
            )
        )
        if handled:
            print("Plan Completed dialog found; clicked Yes, proceed!")
            browser_sleep(0.1)
        return handled
    except Exception:
        return False


def browser_open_screening_school(driver, institute_name: str, speed_mode: Any = "") -> None:
    browser_handle_mark_attendance_modal(driver)
    if not institute_name:
        raise RuntimeError("School/Institute name is blank in Excel.")
    print(f"Selecting school from Excel Institute Name: {institute_name}")
    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 14.0))
    result = driver.execute_script(
        """
        const wanted = arguments[0];
        const visible = el => {
          const r = el.getBoundingClientRect();
          const s = getComputedStyle(el);
          return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
        };
        const norm = s => String(s || '').toLowerCase().replace(/[()\\[\\]{}.,_-]+/g, ' ').replace(/\\s+/g, ' ').trim();
        const schoolKey = s => norm(s).replace(/\\s+\\d{3,}$/, '').trim();
        const target = schoolKey(wanted);
        const rows = Array.from(document.querySelectorAll('tr, [role="row"]')).filter(visible);
        const match = rows
          .map(row => ({row, cells: Array.from(row.querySelectorAll('td, [role="cell"]')).filter(visible)}))
          .filter(x => x.cells.length >= 8)
          .map(x => ({...x, cellTexts: x.cells.map(cell => schoolKey(cell.innerText || cell.textContent)).filter(Boolean)}))
          .map(x => ({
            ...x,
            exact: x.cellTexts.includes(target),
            partial: x.cellTexts.some(text => text.includes(target) || target.includes(text)),
          }))
          .filter(x => x.exact || x.partial)
          .sort((a, b) => Number(b.exact) - Number(a.exact))[0];
        if (!match) return {status: 'row_missing'};
        match.row.scrollIntoView({block:'center', inline:'center'});
        const actionText = el => `${el.title||''} ${el.getAttribute('aria-label')||''} ${el.className||''} ${el.innerText||''} ${el.textContent||''} ${el.innerHTML||''}`.toLowerCase();
        const isLocked = el => {
          const text = actionText(el);
          const clickable = el.closest('button, a, [role="button"]') || el;
          const style = getComputedStyle(clickable);
          return clickable.disabled
            || clickable.getAttribute('aria-disabled') === 'true'
            || /lock|restricted|restriction|disabled|closed/.test(text)
            || /rgb\\(255,\\s*235|rgb\\(254,\\s*226|rgb\\(252,\\s*165|rgb\\(239,\\s*68/.test(style.backgroundColor || '');
        };
        const interactiveInCell = cell => Array.from(cell.querySelectorAll('button, a, [role="button"], input[type="button"], input[type="submit"], i, svg'))
          .filter(visible)
          .map(el => el.closest('button, a, [role="button"]') || el)
          .filter((el, index, arr) => arr.indexOf(el) === index);
        const actionCells = match.cells
          .map(cell => ({cell, actions: interactiveInCell(cell)}))
          .filter(item => item.actions.length)
          .slice(-2);
        for (const item of actionCells) {
          const safe = item.actions.filter(el => !isLocked(el));
          const preferred = safe.find(el => /view|eye|show|open|student|screen/i.test(actionText(el)));
          const clickable = preferred
            || safe.find(el => el.matches('button, a, [role="button"], input[type="button"], input[type="submit"]'))
            || safe[0];
          if (clickable) {
            clickable.scrollIntoView({block: 'center', inline: 'center'});
            if (typeof clickable.click === 'function') {
              clickable.click();
            } else {
              const r = clickable.getBoundingClientRect();
              const options = {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: r.left + r.width / 2,
                clientY: r.top + r.height / 2,
              };
              for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                clickable.dispatchEvent(new MouseEvent(type, options));
              }
            }
            return {status: 'clicked'};
          }
        }
        return {status: 'action_missing'};
        """,
        institute_name,
    )
    status = clean((result or {}).get("status")) if isinstance(result, dict) else ""
    if status != "clicked":
        if status == "action_missing":
            raise RuntimeError(f"School row found, but its green View action was not available: {institute_name}")
        raise RuntimeError(f"Could not find school row in Screening list: {institute_name}")
    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 16.0))
    browser_handle_mark_attendance_modal(driver)
    browser_handle_plan_completed_modal(driver)


def browser_open_student_health_form(driver, row: dict[str, Any], speed_mode: Any = "") -> None:
    name = screening_student_name(row)
    if not name:
        raise RuntimeError("Student name is blank in Excel.")
    print(f"Finding Excel student and opening Start Screening: {name}")
    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 14.0))
    browser_handle_plan_completed_modal(driver)

    try:
        expanded = driver.execute_script(
            """
            const select = Array.from(document.querySelectorAll('select')).find(s => 
              Array.from(s.options).some(o => /25|50|100/.test(o.value || o.text))
            );
            if (select) {
              const options = Array.from(select.options);
              const maxOpt = options.reduce((best, o) => {
                const val = parseInt(o.value || o.text, 10);
                return (!isNaN(val) && val > best.val) ? {o, val} : best;
              }, {o: null, val: 0});
              if (maxOpt.o && select.value !== maxOpt.o.value) {
                select.value = maxOpt.o.value;
                select.dispatchEvent(new Event('change', {bubbles: true}));
                return true;
              }
            }
            return false;
            """
        )
        if expanded:
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 4.0, 8.0))
            browser_sleep(speed_wait(speed_mode, 0.3, 0.6))
    except Exception:
        pass

    max_page_attempts = 15
    for page_num in range(1, max_page_attempts + 1):
        result = driver.execute_script(
            """
            const wantedName = arguments[0];
            const wantedMobile = arguments[1];
            const norm = s => String(s || '').toLowerCase().replace(/[()\\[\\]{}.,_-]+/g, ' ').replace(/\\s+/g, ' ').trim();
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const targetName = norm(wantedName);
            const rows = Array.from(document.querySelectorAll('tr, [role="row"]')).filter(visible);
            const candidates = rows
              .map(row => ({
                row,
                text: norm(row.innerText || row.textContent),
                cellTexts: Array.from(row.querySelectorAll('td, [role="cell"]')).map(cell => norm(cell.innerText || cell.textContent)).filter(Boolean),
              }))
              .map(item => ({
                ...item,
                exactName: item.cellTexts.includes(targetName),
                nameMatch: item.cellTexts.some(text => text === targetName || text.includes(targetName) || targetName.includes(text)),
                mobileMatch: Boolean(wantedMobile && item.text.includes(wantedMobile)),
              }))
              .filter(item => item.nameMatch)
              .sort((a, b) => Number(b.exactName) - Number(a.exactName) || Number(b.mobileMatch) - Number(a.mobileMatch));
            const match = candidates[0];
            if (!match) return {status: 'student_missing'};
            match.row.scrollIntoView({block:'center', inline:'center'});
            const actionText = el => `${el.title||''} ${el.getAttribute('aria-label')||''} ${el.className||''} ${el.innerText||''} ${el.textContent||''} ${el.innerHTML||''}`.toLowerCase();
            const isLocked = el => {
              const text = actionText(el);
              const clickable = el.closest('button, a, [role="button"]') || el;
              const style = getComputedStyle(clickable);
              return clickable.disabled
                || clickable.getAttribute('aria-disabled') === 'true'
                || /lock|restricted|restriction|disabled|closed|completed|screened|healthy|unhealthy|identified|referral|referred|done|check|fa-check/.test(text)
                || /rgb\\(255,\\s*235|rgb\\(254,\\s*226|rgb\\(252,\\s*165|rgb\\(239,\\s*68/.test(style.backgroundColor || '');
            };
            const titledStart = Array.from(match.row.querySelectorAll('[title]'))
              .filter(visible)
              .find(el => norm(el.getAttribute('title')) === 'start screening' || norm(el.getAttribute('title')).includes('start screening'));
            const actions = Array.from(match.row.querySelectorAll('[title], button, a, [role="button"], i, svg, span'))
              .filter(visible)
              .map(el => el.closest('[title="Start Screening"], button, a, [role="button"]') || el)
              .filter((el, index, arr) => arr.indexOf(el) === index)
              .filter(el => !isLocked(el));
            const health = titledStart
              || actions.find(el => norm(el.getAttribute('title')) === 'start screening' || norm(el.getAttribute('title')).includes('start screening'))
              || actions.find(el => /stethoscope|fa-stethoscope|heartbeat|fa-heartbeat|user-md|medical/i.test(actionText(el)))
              || actions.find(el => /health|doctor|screen/i.test(actionText(el)) && !/check|pencil|edit|view/i.test(actionText(el)));
            if (!health) return {status: 'action_missing'};
            health.scrollIntoView({block: 'center', inline: 'center'});
            if (typeof health.click === 'function') {
              health.click();
            } else {
              const r = health.getBoundingClientRect();
              const options = {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: r.left + r.width / 2,
                clientY: r.top + r.height / 2,
              };
              for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                health.dispatchEvent(new MouseEvent(type, options));
              }
            }
            return {status: 'clicked'};
            """,
            name,
            screening_mobile(row),
        )
        status = clean((result or {}).get("status")) if isinstance(result, dict) else ""
        if status == "clicked":
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 16.0))
            # Verify health form actually loaded after clicking Start Screening
            form_opened = False
            for _verify in range(15):
                if browser_handle_not_eligible_modal(driver):
                    print(f"Student '{name}' is NOT eligible for screening. Skipped! Continuing next entry.")
                    raise RuntimeError(f"NOT_ELIGIBLE: Student '{name}' is not eligible for screening.")
                if browser_is_screening_health_form_visible(driver):
                    form_opened = True
                    break
                browser_sleep(0.5)
            if form_opened:
                return
            # JS click did not trigger the Angular route — retry with Selenium native click
            print(f"Start Screening JS click for '{name}' did not open health form. Retrying with native click...")
            browser_sleep(1.0)
            try:
                from selenium.webdriver.common.action_chains import ActionChains as _AC
                start_el = driver.execute_script(
                    """
                    const wanted = arguments[0];
                    const norm = s => String(s || '').toLowerCase().replace(/[()\\[\\]{}.,_-]+/g, ' ').replace(/\\s+/g, ' ').trim();
                    const visible = el => {
                        const r = el.getBoundingClientRect();
                        const s = getComputedStyle(el);
                        return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                    };
                    const target = norm(wanted);
                    const rows = Array.from(document.querySelectorAll('tr, [role="row"]')).filter(visible);
                    for (const row of rows) {
                        if (!norm(row.innerText || row.textContent).includes(target)) continue;
                        // Priority: [title="Start Screening"] element
                        const titled = Array.from(row.querySelectorAll('[title]')).filter(visible)
                            .find(el => /start screening/i.test(el.getAttribute('title') || ''));
                        if (titled) {
                            const clickable = titled.closest('button, a, [role="button"]') || titled;
                            clickable.scrollIntoView({block: 'center', inline: 'center'});
                            return clickable;
                        }
                        // Fallback: stethoscope icon
                        const steth = Array.from(row.querySelectorAll('i, svg, span, button, a')).filter(visible)
                            .find(el => /stethoscope|fa-stethoscope/i.test((el.className || '') + ' ' + (el.innerHTML || '')));
                        if (steth) {
                            const clickable = steth.closest('button, a, [role="button"]') || steth;
                            clickable.scrollIntoView({block: 'center', inline: 'center'});
                            return clickable;
                        }
                    }
                    return null;
                    """,
                    name,
                )
                if start_el:
                    _AC(driver).move_to_element(start_el).pause(0.3).click(start_el).perform()
                    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 16.0))
                    for _verify2 in range(15):
                        if browser_is_screening_health_form_visible(driver):
                            return
                        browser_sleep(0.5)
            except Exception as retry_exc:
                print(f"Retry click error: {retry_exc}")
            raise RuntimeError(
                f"Clicked Start Screening for '{name}' but the health form did not open. "
                f"Check if this student can be screened in the portal."
            )
        if status == "action_missing":
            raise RuntimeError(f"Student found, but Start Screening was not available: {name}")

        hasNextPage = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const candidates = Array.from(document.querySelectorAll('button, a, [role="button"], li, span, input[type="button"]')).filter(visible);
                const nextBtn = candidates.find(el => {
                  const text = norm(el.innerText || el.textContent || el.getAttribute('aria-label') || el.title || el.value || '');
                  const isNext = (text === 'next' || text === 'next >' || text.endsWith(' next') || text === '>') || el.classList.contains('next') || el.classList.contains('pagination-next');
                  const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true' || el.classList.contains('disabled') || el.parentElement?.classList.contains('disabled');
                  return isNext && !disabled;
                });
                if (!nextBtn) return false;
                nextBtn.scrollIntoView({block: 'center', inline: 'center'});
                nextBtn.click();
                return true;
                """
            )
        )
        if hasNextPage:
            print(f"Student '{name}' not on page {page_num}, clicking Next page...")
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
            browser_sleep(speed_wait(speed_mode, 0.4, 0.8))
        else:
            break

    raise RuntimeError(f"Could not find Excel student in the RBSK student list across pages: {name}")


def browser_fill_screening_health_form(driver, row: dict[str, Any], speed_mode: Any = "") -> None:
    data = screening_health_payload(row)
    calc_wt, calc_ht, calc_hc, calc_muac = compute_age_weight_height(row)
    if not clean(data.get("weight")):
        data["weight"] = calc_wt
    if not clean(data.get("height")):
        data["height"] = calc_ht
    if not clean(data.get("headCircumference")):
        data["headCircumference"] = calc_hc
    if not clean(data.get("muac")):
        data["muac"] = calc_muac
    if not clean(data.get("bloodPressure")):
        data["bloodPressure"] = "100/60"
    if not clean(data.get("visionLeft")):
        data["visionLeft"] = "6/6"
    if not clean(data.get("visionRight")):
        data["visionRight"] = "6/6"
    data["ageText"] = row_first_value(row, ["ageText", "Age", "Student Age"])

    # Safeguard: ensure weight is a valid number between 1 and 120 kg
    try:
        wt_val = float(clean(data.get("weight")) or "0")
        if wt_val < 1.0 or wt_val > 120.0:
            data["weight"] = calc_wt or "25"
    except (ValueError, TypeError):
        data["weight"] = calc_wt or "25"

    # Remove 'mobile' from data before passing to JS — the mobile number is NOT
    # needed by the health form JS and its presence can leak into Angular's model,
    # causing the mobile number (e.g. 9158979553) to end up in the Weight field.
    data.pop("mobile", None)
    data["hasExcelValues"] = bool(
        clean(
            row_first_value(
                row, ["Weight", "Weight (kg)", "Weight Kg", "WT", "Weight(kg)"]
            )
        )
    )

    print(
        "Entering Screening Health values: "
        f"weight={data['weight']}, height={data['height']}, "
        f"head circ={data.get('headCircumference', '')}, muac={data.get('muac', '')}, "
        f"blood pressure={data['bloodPressure']}, left vision={data['visionLeft']}, right vision={data['visionRight']}"
    )
    # Ensure health form page (Screening for / Child Details) is actually loaded
    form_loaded = False
    for _ in range(25):
        if browser_handle_not_eligible_modal(driver):
            print(f"Child '{row.get('Student Full Name', 'Child')}' is NOT eligible for AWC screening (Age >= 6 yrs). Skipped! Continuing next entry.")
            return
        form_loaded = bool(
            driver.execute_script(
                """
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent || '');
                return text.includes('screening for:') || text.includes('child details') || text.includes('location details') || (text.includes('weight') && text.includes('height'));
                """
            )
        )
        if form_loaded:
            break
        browser_sleep(0.4)

    if not form_loaded:
        if browser_handle_not_eligible_modal(driver):
            print(f"Child '{row.get('Student Full Name', 'Child')}' is NOT eligible for AWC screening. Skipped! Continuing next entry.")
            return
        raise RuntimeError("Health form page (Screening for / Child Details) did not open.")

    # Pre-scroll: unlock overflow:hidden on all scrollable containers so form
    # fields below the fold (weight, height, BP, vision) become reachable.
    driver.execute_script("""
        const containers = Array.from(document.querySelectorAll('div, main, section, article, form'));
        for (const el of containers) {
            try {
                const cs = getComputedStyle(el);
                if ((cs.overflow === 'hidden' || cs.overflowY === 'hidden') && el.scrollHeight > el.clientHeight + 10) {
                    el.style.overflow = 'auto';
                }
            } catch(e) {}
        }
    """)
    browser_sleep(0.6)
    ok = driver.execute_script(
        """
        const data = arguments[0];
        const cleanBody = (document.body.innerText || document.body.textContent || '').replace(/[\\s\\u00a0]+/g, ' ');
        let pageAgeYrs = null;

        // Parse Age text from live portal DOM (e.g. "Age: 0 years 3 months" or "Age: 3 months")
        const fullAgeMatch = cleanBody.match(/Age\\s*:\\s*(\\d+)\\s*years?(?:\\s*(\\d+)\\s*months?)?/i)
                          || cleanBody.match(/Age\\s*:\\s*(\\d+)\\s*months?/i)
                          || cleanBody.match(/(\\d+)\\s*years?\\s*(\\d+)\\s*months?/i);
        if (fullAgeMatch) {
          if (fullAgeMatch[2] !== undefined) {
            const y = parseInt(fullAgeMatch[1], 10) || 0;
            const m = parseInt(fullAgeMatch[2], 10) || 0;
            pageAgeYrs = y + (m / 12.0);
          } else {
            const num = parseInt(fullAgeMatch[1], 10) || 0;
            if (/months?/i.test(fullAgeMatch[0])) {
              pageAgeYrs = num / 12.0;
            } else {
              pageAgeYrs = num;
            }
          }
        }
        if (pageAgeYrs == null && data.ageText) {
          const tMatch = String(data.ageText).replace(/[\\s\\u00a0]+/g, ' ').match(/(\\d+)\\s*years?(?:\\s*(\\d+)\\s*months?)?/i);
          if (tMatch) {
            const y = parseInt(tMatch[1], 10) || 0;
            const m = parseInt(tMatch[2], 10) || 0;
            pageAgeYrs = y + (m / 12.0);
          } else {
            const tmMatch = String(data.ageText).match(/(\\d+)\\s*months?/i);
            if (tmMatch) {
              pageAgeYrs = (parseInt(tmMatch[1], 10) || 0) / 12.0;
            }
          }
        }

        if (pageAgeYrs != null) {
          const points = [
            { a: 0.0, w: 3.4, h: 50.0, hc: 34.5, m: null },  // Birth
            { a: 0.5, w: 7.5, h: 67.0, hc: 42.5, m: 13.25 }, // 6 Months
            { a: 1.0, w: 10.0, h: 76.0, hc: 46.0, m: 14.0 }, // 1 Year
            { a: 2.0, w: 12.5, h: 87.0, hc: 48.0, m: 14.75 },// 2 Years
            { a: 3.0, w: 14.5, h: 95.0, hc: 49.0, m: 15.25 },// 3 Years
            { a: 4.0, w: 17.0, h: 103.0, hc: 50.0, m: 15.5 },// 4 Years
            { a: 5.0, w: 20.0, h: 110.5, hc: 51.0, m: 16.0 },// 5 Years
            { a: 6.0, w: 21.0, h: 116.0, hc: 51.0, m: 16.0 },
            { a: 12.0, w: 33.0, h: 140.0, hc: null, m: null }
          ];

          let wt = "18", ht = "103", hc = "50", muac = "15.5";
          if (pageAgeYrs <= points[0].a) {
            wt = points[0].w.toFixed(1);
            ht = Math.round(points[0].h).toString();
            hc = points[0].hc ? points[0].hc.toFixed(1) : "";
            muac = "";
          } else if (pageAgeYrs >= points[points.length - 1].a) {
            const p = points[points.length - 1];
            wt = p.w.toFixed(1);
            ht = Math.round(p.h).toString();
            hc = "";
            muac = "";
          } else {
            for (let i = 0; i < points.length - 1; i++) {
              const p0 = points[i];
              const p1 = points[i + 1];
              if (pageAgeYrs >= p0.a && pageAgeYrs <= p1.a) {
                const t = (pageAgeYrs - p0.a) / (p1.a - p0.a);
                wt = (p0.w + t * (p1.w - p0.w)).toFixed(1);
                ht = Math.round(p0.h + t * (p1.h - p0.h)).toString();
                if (p0.hc !== null && p1.hc !== null) {
                  hc = (p0.hc + t * (p1.hc - p0.hc)).toFixed(1);
                } else if (p0.hc !== null) {
                  hc = p0.hc.toFixed(1);
                } else {
                  hc = "";
                }
                if (p0.m !== null && p1.m !== null) {
                  muac = (p0.m + t * (p1.m - p0.m)).toFixed(1);
                } else if (p1.m !== null && pageAgeYrs >= 0.5) {
                  muac = p1.m.toFixed(1);
                } else {
                  muac = "";
                }
                break;
              }
            }
          }

          if (!data.hasExcelValues) {
            data.weight = wt;
            data.height = ht;
            data.headCircumference = hc;
            data.muac = muac;
          }
        }

        if (!data.bloodPressure || data.bloodPressure === "100/60" || data.bloodPressure === "0") {
          const sys = 100;
          const dia = Math.floor(Math.random() * 11) + 65; // 65 to 75
          data.bloodPressure = `${sys}/${dia}`;
        }
        if (!data.visionLeft) data.visionLeft = "6/6";
        if (!data.visionRight) data.visionRight = "6/6";

        // Uncheck all checked checkboxes
        const uncheckAllCheckboxes = () => {
          const allCheckboxes = Array.from(document.querySelectorAll('input[type="checkbox"]'));
          for (const chk of allCheckboxes) {
            if (chk.checked) chk.click();
          }
        };
        uncheckAllCheckboxes();

        const forceScrollTo = (el) => {
          if (!el) return;
          let ancestor = el.parentElement;
          while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
            try {
              const cs = getComputedStyle(ancestor);
              if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') ancestor.style.overflow = 'auto';
            } catch(e) {}
            ancestor = ancestor.parentElement;
          }
          try { el.scrollIntoView({block:'center', inline:'center', behavior:'instant'}); } catch(e) {}
        };

        const norm = s => String(s || '').toLowerCase().replace(/[\\s\\u00a0]+/g, ' ').trim();

        const setValue = (el, value) => {
          if (!el || value == null || String(value) === '') return false;
          forceScrollTo(el);
          el.focus();
          const proto = HTMLInputElement.prototype;
          const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set || Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
          if (setter) setter.call(el, String(value));
          else el.value = String(value);
          el.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
          el.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
          el.dispatchEvent(new Event('keyup', { bubbles: true, composed: true }));
          el.dispatchEvent(new Event('blur', { bubbles: true, composed: true }));
          return true;
        };

        // Helper to find exact input for a label by matching text
        const findInputByLabel = (labelKeywords) => {
          const allEls = Array.from(document.querySelectorAll('label, div, span, td, th, p, b, strong'));
          for (const labelEl of allEls) {
            const txt = norm(labelEl.innerText || labelEl.textContent || '');
            if (txt.length > 0 && txt.length < 80 && labelKeywords.some(k => txt.startsWith(k) || txt.includes(k))) {
              // 1. Check direct control
              if (labelEl.control && inputs.includes(labelEl.control)) return labelEl.control;
              // 2. Check htmlFor
              if (labelEl.htmlFor) {
                const linked = document.getElementById(labelEl.htmlFor);
                if (linked && inputs.includes(linked)) return linked;
              }
              // 3. Check inside parent container
              const parent = labelEl.closest('div, td, tr, form, section') || labelEl.parentElement;
              if (parent) {
                const nearby = parent.querySelector('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])');
                if (nearby && inputs.includes(nearby)) return nearby;
              }
            }
          }
          return null;
        };

        // All text/number input fields on the screening form (unrestricted query)
        const getInputs = () => Array.from(document.querySelectorAll('input[type="text"], input[type="number"], input[type="tel"], input:not([type]), textarea'))
          .filter(el => {
            const type = String(el.type || '').toLowerCase();
            if (type === 'hidden' || type === 'submit' || type === 'button' || type === 'checkbox' || type === 'radio' || type === 'file') return false;
            const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
            if (/search|filter|mobile|phone|contact/.test(info)) return false;
            return true;
          });

        const inputs = getInputs();

        // 1. Find Weight field (Must NOT match height)
        let wEl = inputs.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('weight') || /\bwt\b/.test(info);
        }) || findInputByLabel(['weight (kg)', 'weight']);
        if (!wEl && inputs.length >= 1) wEl = inputs[0];

        // 2. Find Height field (CRITICAL: Must NOT match 'weight' which contains 'ht'!)
        let hEl = inputs.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('height') || info.includes('length') || /\bht\b/.test(info);
        });
        if (!hEl || hEl === wEl) {
          hEl = findInputByLabel(['height/length (cm)', 'height/length', 'height (cm)', 'height', 'length']);
        }
        if ((!hEl || hEl === wEl) && inputs.length >= 2) hEl = inputs[1];

        // 3. Find Head Circumference field
        let hcEl = inputs.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('head') || info.includes('headcircumference') || /\bhc\b/.test(info);
        }) || findInputByLabel(['head circumference (cm)', 'head circumference', 'head circ']);
        if (!hcEl && inputs.length >= 3) hcEl = inputs[2];

        // 4. Find MUAC field
        let muacEl = inputs.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('muac') || info.includes('muic') || info.includes('arm');
        }) || findInputByLabel(['muac (cm)', 'muac', 'muic', 'arm circumference']);
        if (!muacEl && inputs.length >= 4) muacEl = inputs[3];

        // 5. Find Blood Pressure field
        let bpEl = inputs.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.placeholder||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('blood') || info.includes('bloodpressure') || /\bbp\b/.test(info);
        }) || findInputByLabel(['blood pressure']);

        // Fill inputs cleanly
        const wOk = setValue(wEl, data.weight);
        const hOk = setValue(hEl, data.height);
        const hcOk = !data.headCircumference || setValue(hcEl, data.headCircumference);
        const muacOk = !data.muac || setValue(muacEl, data.muac);
        const bpOk = setValue(bpEl, data.bloodPressure);

        // 6. Fill Vision Selects (Left Eye & Right Eye)
        const allSelects = Array.from(document.querySelectorAll('select'));

        // Find Left Eye Select
        let leftSelect = allSelects.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('left') || info.includes('visionleft');
        });
        if (!leftSelect) {
          const label = Array.from(document.querySelectorAll('label, div, span, td')).find(el => norm(el.innerText || el.textContent).includes('vision - left'));
          if (label) {
            const parent = label.closest('div, td, tr, form') || label.parentElement;
            leftSelect = parent?.querySelector('select');
          }
        }

        // Find Right Eye Select
        let rightSelect = allSelects.find(el => {
          const info = norm(`${el.name||''} ${el.id||''} ${el.getAttribute('formcontrolname')||''}`);
          return info.includes('right') || info.includes('visionright');
        });
        if (!rightSelect) {
          const label = Array.from(document.querySelectorAll('label, div, span, td')).find(el => norm(el.innerText || el.textContent).includes('vision - right'));
          if (label) {
            const parent = label.closest('div, td, tr, form') || label.parentElement;
            rightSelect = parent?.querySelector('select');
          }
        }

        // Fallback: vision selects by 6/6 options
        const visionSelectsByOption = allSelects.filter(s =>
          Array.from(s.options).some(o => (o.textContent || '').includes('6/6') || (o.value || '').includes('6/6'))
        );
        if (!leftSelect && visionSelectsByOption.length >= 1) leftSelect = visionSelectsByOption[0];
        if (!rightSelect && visionSelectsByOption.length >= 2) rightSelect = visionSelectsByOption[1];
        if (!rightSelect && visionSelectsByOption.length >= 1 && visionSelectsByOption[0] !== leftSelect) rightSelect = visionSelectsByOption[0];

        const setSelect = (sel, val) => {
          if (!sel) return false;
          forceScrollTo(sel);
          try { sel.disabled = false; sel.removeAttribute('disabled'); } catch(e) {}
          const target = val || "6/6";
          const opt = Array.from(sel.options).find(o => norm(o.textContent) === norm(target) || norm(o.textContent).includes(norm(target)) || norm(o.value) === norm(target))
            || Array.from(sel.options).find(o => (o.textContent || '').includes('6/6') || (o.value || '').includes('6/6'))
            || sel.options[1];
          if (opt) {
            sel.value = opt.value;
            opt.selected = true;
            sel.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
            sel.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
            sel.dispatchEvent(new Event('blur', { bubbles: true, composed: true }));
            return true;
          }
          return false;
        };

        const leftOk = setSelect(leftSelect, data.visionLeft || "6/6");
        const rightOk = setSelect(rightSelect, data.visionRight || "6/6");

        uncheckAllCheckboxes();

        return {
          weight: wOk,
          height: hOk,
          headCircumference: hcOk,
          muac: muacOk,
          bp: bpOk,
          left: leftOk,
          right: rightOk
        };
        """,
        data,
    )
    browser_sleep(speed_wait(speed_mode, 0.05, 0.2))
    if not ok.get("weight") or not ok.get("height"):
        print(f"Warning: Health form fill state: {ok}")

    # Scroll page down so form fields below the fold and Submit / Preview button are exposed
    driver.execute_script("""
        window.scrollTo(0, document.body.scrollHeight);
        const scrollables = Array.from(document.querySelectorAll('div, main, section, article, form'));
        for (const el of scrollables) {
            try {
                if (el.scrollHeight > el.clientHeight + 10) {
                    el.scrollTop = el.scrollHeight;
                }
            } catch(e) {}
        }
    """)
    browser_sleep(speed_wait(speed_mode, 0.05, 0.2))


def browser_uncheck_other_health_condition(driver) -> None:
    """Uncheck 'Other Health Condition' and all checkboxes on the screening health form.
    
    The RBSK portal uses Angular. Only .click() on the DOM element updates Angular's
    internal model. We MUST click only checkboxes that are currently checked.
    Never click an unchecked checkbox (that would toggle it ON).
    """
    try:
        # Pass 1: click any checked checkbox to uncheck it
        unchecked = int(driver.execute_script("""
            const allChk = document.querySelectorAll('input[type="checkbox"]');
            let changed = 0;
            for (const chk of allChk) {
                if (chk.checked) {
                    chk.click();
                    changed += 1;
                }
            }
            return changed;
        """) or 0)
        if unchecked:
            browser_sleep(0.05)
        # Pass 2: verify and retry — if Angular re-rendered and any are still checked
        driver.execute_script("""
            const allChk = document.querySelectorAll('input[type="checkbox"]');
            for (const chk of allChk) {
                if (chk.checked) {
                    chk.click();
                }
            }
        """)
    except Exception:
        pass


def browser_click_preview_submit_button(driver, speed_mode: Any = "") -> bool:
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys
    except Exception:
        ActionChains = None
        Keys = None

    def real_scroll_down() -> None:
        try:
            driver.execute_script(
                """
                try { document.activeElement && document.activeElement.blur && document.activeElement.blur(); } catch(e) {}
                try { document.body && document.body.focus && document.body.focus(); } catch(e) {}
                """
            )
        except Exception:
            pass
        try:
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {"type": "mouseWheel", "x": 960, "y": 900, "deltaX": 0, "deltaY": 1200, "pointerType": "mouse"},
            )
        except Exception:
            pass
        try:
            if Keys:
                body = driver.find_element("tag name", "body")
                body.send_keys(Keys.PAGE_DOWN)
        except Exception:
            pass
        try:
            if ActionChains:
                ActionChains(driver).scroll_by_amount(0, 1200).perform()
        except Exception:
            pass
        browser_sleep(speed_wait(speed_mode, 0.05, 0.15))

    for _attempt in range(10):
        real_scroll_down()
        clicked = bool(
            driver.execute_script(
                """
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const rectVisible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const clickableOf = el => el?.closest?.('button, input[type="button"], input[type="submit"], a, [role="button"]') || el;
                const bestButtonInside = el => {
                  const direct = clickableOf(el);
                  if (direct && direct.matches?.('button, input[type="button"], input[type="submit"], a, [role="button"]')) return direct;
                  const kids = Array.from(el.querySelectorAll?.('button, input[type="button"], input[type="submit"], a, [role="button"]') || [])
                    .map(child => ({el: child, text: norm(child.innerText || child.textContent || child.value || child.getAttribute('aria-label') || child.title || '')}))
                    .filter(x => x.text === 'preview & submit' || x.text === 'preview and submit' || (x.text.includes('preview') && x.text.includes('submit')));
                  return kids[0]?.el || direct || el;
                };
                const unlockAndScroll = target => {
                  const roots = [document.scrollingElement, document.documentElement, document.body, document.getElementById('root')]
                    .concat(Array.from(document.querySelectorAll('div, main, section, article, form, table, tbody')));
                  for (const el of roots) {
                    if (!el) continue;
                    try {
                      const cs = getComputedStyle(el);
                      if (el.scrollHeight > el.clientHeight + 8) {
                        if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') el.style.overflowY = 'auto';
                        el.scrollTop = el.scrollHeight;
                      }
                    } catch(e) {}
                  }
                  let p = target?.parentElement;
                  while (p) {
                    try {
                      const cs = getComputedStyle(p);
                      if (p.scrollHeight > p.clientHeight + 8) {
                        if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') p.style.overflowY = 'auto';
                        const pr = p.getBoundingClientRect();
                        const tr = target.getBoundingClientRect();
                        p.scrollTop += tr.top - pr.top - Math.floor(p.clientHeight * 0.55);
                      }
                    } catch(e) {}
                    p = p.parentElement;
                  }
                  try { window.scrollTo(0, Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)); } catch(e) {}
                  try { target?.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'}); } catch(e) {}
                };

                const primaryElements = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a, [role="button"]'));
                const fallbackElements = Array.from(document.querySelectorAll('span, div'));
                const matches = primaryElements.concat(fallbackElements)
                  .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}))
                  .filter(x => x.text === 'preview & submit' || x.text === 'preview and submit' || (x.text.includes('preview') && x.text.includes('submit')));
                if (!matches.length) {
                  unlockAndScroll(document.body);
                  return false;
                }
                const raw = matches.find(x => x.el.matches?.('button, input[type="button"], input[type="submit"], a, [role="button"]') && (x.text === 'preview & submit' || x.text === 'preview and submit'))?.el
                  || matches.find(x => x.el.matches?.('button, input[type="button"], input[type="submit"], a, [role="button"]'))?.el
                  || matches.find(x => x.text === 'preview & submit' || x.text === 'preview and submit')?.el
                  || matches[0].el;
                const btn = bestButtonInside(raw);
                unlockAndScroll(btn);
                try { btn.disabled = false; } catch(e) {}
                try { btn.removeAttribute('disabled'); } catch(e) {}
                try { btn.removeAttribute('aria-disabled'); } catch(e) {}
                try { btn.classList.remove('disabled', 'is-disabled'); } catch(e) {}
                let r = btn.getBoundingClientRect();
                if (!rectVisible(btn)) return false;
                const cx = Math.max(2, Math.min(window.innerWidth - 3, r.left + r.width / 2));
                const cy = Math.max(2, Math.min(window.innerHeight - 3, r.top + r.height / 2));
                const topEl = document.elementFromPoint(cx, cy);
                const clickable = clickableOf(topEl || btn);
                const target = clickable?.matches?.('button, input[type="button"], input[type="submit"], a, [role="button"]')
                  ? clickable : btn;
                try { target.focus?.(); } catch(e) {}
                target.click();
                return true;
                """
            )
        )
        if clicked:
            return True
        browser_sleep(speed_wait(speed_mode, 0.25, 0.45))

    for _ in range(4):
        real_scroll_down()

    try:
        if ActionChains:
            for _ in range(4):
                ActionChains(driver).scroll_by_amount(0, 900).perform()
                browser_sleep(speed_wait(speed_mode, 0.15, 0.3))
    except Exception:
        pass

    for _attempt in range(8):
        clicked = bool(
            driver.execute_script(
                """
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && !el.disabled;
                };
                const hardScrollBottom = () => {
                  const roots = [document.scrollingElement, document.documentElement, document.body, document.getElementById('root')]
                    .concat(Array.from(document.querySelectorAll('div, main, section, article, form')));
                  for (const el of roots) {
                    if (!el) continue;
                    try {
                      const cs = getComputedStyle(el);
                      if ((cs.overflow === 'hidden' || cs.overflowY === 'hidden') && el.scrollHeight > el.clientHeight + 10) {
                        el.style.overflow = 'auto';
                      }
                      if (el.scrollHeight > el.clientHeight + 10) {
                        el.scrollTop = el.scrollHeight;
                      }
                    } catch(e) {}
                  }
                  try { window.scrollTo(0, Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)); } catch(e) {}
                };
                const candidates = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a, [role="button"]'))
                  .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}));
                const match = candidates.find(x => x.text.includes('preview') && x.text.includes('submit'))
                  || candidates.find(x => x.text.includes('preview'));
                if (!match) {
                  hardScrollBottom();
                  return false;
                }
                const btn = match.el;
                let parent = btn.parentElement;
                while (parent) {
                  try {
                    if (parent.scrollHeight > parent.clientHeight + 10) {
                      parent.scrollTop = Math.max(parent.scrollTop, btn.offsetTop - Math.floor(parent.clientHeight / 2));
                    }
                  } catch(e) {}
                  parent = parent.parentElement;
                }
                hardScrollBottom();
                try { btn.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'}); } catch(e) {}
                let r = btn.getBoundingClientRect();
                if (r.top < 0 || r.bottom > window.innerHeight) {
                  try { window.scrollTo(0, Math.max(0, window.scrollY + r.top - Math.floor(window.innerHeight / 2))); } catch(e) {}
                }
                try { btn.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'}); } catch(e) {}
                r = btn.getBoundingClientRect();
                if (!visible(btn) || r.bottom < 0 || r.top > window.innerHeight) return false;
                const cx = Math.max(1, Math.min(window.innerWidth - 2, r.left + r.width / 2));
                const cy = Math.max(1, Math.min(window.innerHeight - 2, r.top + r.height / 2));
                const topEl = document.elementFromPoint(cx, cy);
                const clickable = topEl?.closest?.('button, input[type="button"], input[type="submit"], a, [role="button"]') || btn;
                const target = clickable?.matches?.('button, input[type="button"], input[type="submit"], a, [role="button"]')
                  ? clickable : btn;
                try { target.focus?.(); } catch(e) {}
                target.click();
                return true;
                """
            )
        )
        if clicked:
            return True
        browser_sleep(speed_wait(speed_mode, 0.3, 0.6))
    return False


def browser_preview_submit_screening(driver, row: dict[str, Any], speed_mode: Any = "") -> None:
    # SAFETY: Ensure form checkboxes remain unchecked before previewing
    browser_uncheck_other_health_condition(driver)
    
    # 1. Scroll to bottom of page to expose Preview & Submit button.
    driver.execute_script(
        """
        const scrollBottom = () => {
          window.scrollTo({top: document.body.scrollHeight, behavior: 'instant'});
          document.documentElement.scrollTop = document.documentElement.scrollHeight;
          document.body.scrollTop = document.body.scrollHeight;
          const scrollables = Array.from(document.querySelectorAll('body, html, div, section, form, main, article'))
            .filter(el => el && el.scrollHeight > el.clientHeight + 20);
          for (const s of scrollables) {
            try {
              const cs = getComputedStyle(s);
              if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') {
                s.style.overflow = 'auto';
              }
              s.scrollTop = s.scrollHeight;
            } catch(e) {}
          }
        };
        scrollBottom();
        setTimeout(scrollBottom, 120);
        """
    )

    # 2. Click the Preview & Submit button.
    preview_clicked = browser_click_preview_submit_button(driver, speed_mode=speed_mode)
    for attempt in range(6):
        if preview_clicked:
            break
        preview_clicked = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && !el.disabled;
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const scrollBottom = () => {
                  window.scrollTo({top: document.body.scrollHeight, behavior: 'instant'});
                  document.documentElement.scrollTop = document.documentElement.scrollHeight;
                  document.body.scrollTop = document.body.scrollHeight;
                  for (const s of Array.from(document.querySelectorAll('body, html, div, section, form, main, article')).filter(el => el.scrollHeight > el.clientHeight + 20)) {
                    try {
                      const cs = getComputedStyle(s);
                      if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') {
                        s.style.overflow = 'auto';
                      }
                      s.scrollTop = s.scrollHeight;
                    } catch(e) {}
                  }
                };
                scrollBottom();
                const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], [role="button"], a'))
                  .filter(visible)
                  .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}))
                  .filter(x => x.text.includes('preview') || (x.text.includes('submit') && x.text.includes('preview')));
                if (!buttons.length) return false;
                const exact = buttons.find(x => x.text === 'preview & submit' || x.text === 'preview and submit') || buttons[0];
                const btn = exact.el;
                btn.scrollIntoView({block:'center', inline:'center', behavior:'instant'});
                window.scrollBy(0, -80);
                const r = btn.getBoundingClientRect();
                const cx = r.left + r.width / 2;
                const cy = r.top + r.height / 2;
                const topEl = document.elementFromPoint(cx, cy);
                const clickable = topEl?.closest?.('button, input[type="button"], input[type="submit"], [role="button"], a') || btn;
                const opts = {bubbles: true, cancelable: true, view: window, clientX: cx, clientY: cy};
                for (const t of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                  clickable.dispatchEvent(new MouseEvent(t, opts));
                }
                try { clickable.click(); } catch(e) {}
                try { btn.click(); } catch(e) {}
                return true;
                """
            )
        )
        if preview_clicked:
            break
        browser_sleep(speed_wait(speed_mode, 0.4, 0.8))

    if not preview_clicked:
        raise RuntimeError("Could not click Preview button on Screening form.")

    print("Clicked Preview button. Waiting for Screening Preview modal...")

    # 3. Wait for Preview screen. Some RBSK versions show mobile as read-only
    # preview text, while older versions ask for mobile in a popup input.
    is_awc = bool(row.get("_is_awc"))
    mobile = "" if is_awc else (screening_mobile(row) or "9999999999")
    mobile_filled = False
    mobile_input_required = True
    preview_ready = False
    deadline = time.time() + speed_timeout(speed_mode, 12.0, 25.0)
    while time.time() < deadline:
        preview_state = driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const pageText = norm(document.body.innerText || document.body.textContent || '');
            const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a, [role="button"]'))
              .filter(visible)
              .map(el => norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || ''));
            const hasConfirm = buttons.some(text => text.includes('confirm') && text.includes('submit'));
            const hasPreview = pageText.includes('screening preview') || hasConfirm;
            const mobileInputs = Array.from(document.querySelectorAll('input')).filter(el => {
              if (!visible(el) || el.disabled || el.readOnly) return false;
              const info = norm(`${el.placeholder || ''} ${el.name || ''} ${el.id || ''} ${el.getAttribute('aria-label') || ''}`);
              return info.includes('mobile') || info.includes('phone') || info.includes('contact') || info.includes('enter mobile');
            });
            return {hasPreview, hasConfirm, mobileInputCount: mobileInputs.length};
            """
        )
        preview_ready = bool(preview_state and preview_state.get("hasPreview"))
        if preview_ready:
            mobile_input_required = int(preview_state.get("mobileInputCount") or 0) > 0
            break
        browser_sleep(speed_wait(speed_mode, 0.04, 0.2))

    if not preview_ready:
        raise RuntimeError("Preview screen did not open after clicking Preview & Submit.")

    if is_awc and mobile_input_required:
        driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && !el.disabled && !el.readOnly;
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const inputs = Array.from(document.querySelectorAll('input')).filter(el => {
              if (!visible(el)) return false;
              const info = norm(`${el.placeholder || ''} ${el.name || ''} ${el.id || ''} ${el.getAttribute('aria-label') || ''}`);
              return info.includes('mobile') || info.includes('phone') || info.includes('contact') || info.includes('enter mobile');
            });
            for (const input of inputs) {
              input.focus();
              try { input.required = false; } catch(e) {}
              try { input.removeAttribute('required'); } catch(e) {}
              try { input.removeAttribute('pattern'); } catch(e) {}
              try { input.setCustomValidity(''); } catch(e) {}
              const setter = Object.getOwnPropertyDescriptor(input.__proto__, 'value')?.set || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
              if (setter) setter.call(input, '');
              else input.value = '';
              input.classList.remove('ng-invalid', 'is-invalid', 'invalid');
              input.classList.add('ng-valid');
              input.setAttribute('aria-invalid', 'false');
              input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
              input.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
              input.dispatchEvent(new Event('blur', { bubbles: true, composed: true }));
            }
            return inputs.length;
            """
        )

    if not mobile_input_required:
        if is_awc:
            print("Preview screen opened. AWC/AWS mode: no mobile number entered.")
        else:
            print(f"Preview screen opened. Mobile number is already shown on preview: {mobile}")
        mobile_filled = True
    
    deadline = time.time() + speed_timeout(speed_mode, 8.0, 18.0)
    while mobile_input_required and not is_awc and time.time() < deadline:
        mobile_filled = bool(
            driver.execute_script(
                """
                const mobile = arguments[0];
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && !el.disabled && !el.readOnly;
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                
                // Find visible inputs on the page
                const allInputs = Array.from(document.querySelectorAll('input')).filter(visible);
                if (!allInputs.length) return false;

                // Look specifically for an input whose placeholder, name, id, or aria-label contains 'mobile' or 'phone' or 'contact'
                let target = allInputs.find(el => {
                  const info = norm(`${el.placeholder||''} ${el.name||''} ${el.id||''} ${el.getAttribute('aria-label')||''}`);
                  return info.includes('mobile') || info.includes('phone') || info.includes('contact') || info.includes('enter mobile');
                });

                // Fallback: look for mobile label in modals/dialogs
                if (!target) {
                  const modalContainers = Array.from(document.querySelectorAll('.modal, .dialog, [role="dialog"], .modal-content, [class*="modal"], [class*="preview"], [class*="Preview"], [class*="popup"], bs-modal-container')).filter(visible);
                  for (const mc of modalContainers) {
                    const modalInputs = Array.from(mc.querySelectorAll('input')).filter(visible);
                    if (modalInputs.length === 1) {
                      target = modalInputs[0];
                      break;
                    }
                    const labels = Array.from(mc.querySelectorAll('label, div, td, span')).filter(visible);
                    const mobileLabel = labels.find(el => /mobile\\s*number/i.test(el.innerText || el.textContent || ''));
                    if (mobileLabel) {
                      const parent = mobileLabel.closest('div, td, tr, form, p') || mobileLabel.parentElement;
                      if (parent) {
                        target = parent.querySelector('input');
                        if (target && visible(target)) break;
                      }
                    }
                  }
                }
                
                if (!target) {
                  const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a, [role="button"]'))
                    .filter(el => {
                      const r = el.getBoundingClientRect();
                      const s = getComputedStyle(el);
                      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                    })
                    .map(el => String(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '').toLowerCase());
                  return buttons.some(text => text.includes('confirm') && text.includes('submit'));
                }

                target.focus();
                const setter = Object.getOwnPropertyDescriptor(target.__proto__, 'value')?.set || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
                if (setter) setter.call(target, mobile);
                else target.value = mobile;
                
                target.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
                target.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
                target.dispatchEvent(new Event('blur', { bubbles: true, composed: true }));
                target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
                target.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                return String(target.value || '').replace(/\\D/g, '') === String(mobile);
                """,
                mobile,
            )
        )
        if mobile_filled:
            break
        browser_sleep(speed_wait(speed_mode, 0.03, 0.15))
    if not mobile_filled:
        if is_awc:
            print("AWC/AWS mode: mobile number field found, but no number will be entered. Continuing to confirm submission.")
        else:
            raise RuntimeError(f"Could not fill Preview Mobile Number ({mobile}). Modal may not have loaded.")
    
    if mobile_input_required and mobile_filled:
        print(f"Preview Mobile Number filled: {mobile}")

    # 4. Click Confirm & Submit immediately after the mobile change events.
    confirmed = False
    confirm_deadline = time.time() + speed_timeout(speed_mode, 3.0, 8.0)
    while time.time() < confirm_deadline:
        if browser_click_visible_confirmation(driver, force_enable=is_awc):
            confirmed = True
            break
        browser_sleep(speed_wait(speed_mode, 0.04, 0.2))

    if not confirmed:
        raise RuntimeError("Could not click Confirm & Submit in Preview modal.")

    print("Clicked Confirm & Submit immediately. Waiting for confirmation notice...")
    if not browser_wait_for_screening_saved_notice(
        driver,
        timeout=speed_timeout(speed_mode, 20.0, 45.0),
        interval=speed_wait(speed_mode, 0.05, 0.25),
    ):
        raise RuntimeError("RBSK did not show Screening details saved successfully after Confirm & Submit.")

    print("RBSK Screening details saved successfully.")
    browser_dismiss_success_notice(driver)


def browser_wait_for_screening_saved_notice(driver, timeout: float = 35.0, interval: float = 0.5) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if browser_screening_saved_notice_visible(driver):
            return True
        if browser_screening_school_list_visible(driver):
            return True
        browser_sleep(interval)
    return False


def browser_screening_saved_notice_visible(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent);
                if (/screening details saved successfully|details saved successfully|screening saved successfully/.test(text)) {
                  return true;
                }
                return Array.from(document.querySelectorAll('[role="alert"], .toast, [class*="toast"], [class*="Toast"], [class*="notification"], [class*="Notification"]'))
                  .filter(visible)
                  .some(el => /saved successfully|details saved|submitted successfully/.test(norm(el.innerText || el.textContent)));
                """
            )
        )
    except Exception:
        return False


def browser_return_to_screening_list_after_save(driver, row: dict[str, Any], speed_mode: Any = "") -> str:
    name = screening_student_name(row)
    deadline = time.time() + speed_timeout(speed_mode, 18.0, 35.0)

    def current_location() -> str:
        if browser_screening_preview_visible(driver):
            return ""
        if browser_screening_student_list_visible(driver) or browser_screening_student_list_has_name(driver, name):
            return "student_list"
        if browser_screening_school_list_visible(driver):
            return "school_list"
        return ""

    location = current_location()
    if location:
        return location

    clicked_back = bool(
        driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && !el.disabled;
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const buttons = Array.from(document.querySelectorAll('button, [role="button"], a, input[type="button"]'))
              .filter(visible)
              .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label'))}))
              .filter(x => x.text === 'back' || x.text === '← back' || x.text.endsWith(' back'));
            if (!buttons.length) return false;
            buttons[0].el.scrollIntoView({block: 'center', inline: 'center'});
            buttons[0].el.click();
            return true;
            """
        )
    )
    if not clicked_back:
        try:
            driver.back()
        except Exception:
            pass

    while time.time() < deadline:
        location = current_location()
        if location:
            return location
        browser_dismiss_success_notice(driver)
        clicked_back = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && !el.disabled;
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const pageText = norm(document.body.innerText || document.body.textContent || '');
                const buttons = Array.from(document.querySelectorAll('button, [role="button"], a, input[type="button"]'))
                  .filter(visible)
                  .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}));
                let candidate = buttons.find(x => x.text === 'back' || x.text.includes('back to student') || x.text.includes('back to list'));
                if (!candidate && !pageText.includes('screening preview')) {
                  candidate = buttons.find(x => x.text.endsWith(' back') || x.text.startsWith('back '));
                }
                if (!candidate) return false;
                candidate.el.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'});
                candidate.el.click();
                return true;
                """
            )
        )
        if not clicked_back:
            try:
                driver.back()
            except Exception:
                pass
        browser_sleep(speed_wait(speed_mode, 0.1, 0.7))
    return current_location() or "unknown"


def browser_click_visible_confirmation(driver, force_enable: bool = False) -> bool:
    return bool(
        driver.execute_script(
            """
            const forceEnable = Boolean(arguments[0]);
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && (forceEnable || !el.disabled);
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const scrollEverywhere = () => {
              const roots = [document.scrollingElement, document.documentElement, document.body, document.getElementById('root')]
                .concat(Array.from(document.querySelectorAll('.modal, .modal-content, [role="dialog"], bs-modal-container, div, main, section, article, form')));
              for (const el of roots) {
                if (!el) continue;
                try {
                  if (el.scrollHeight > el.clientHeight + 10) el.scrollTop = el.scrollHeight;
                } catch(e) {}
              }
            };
            scrollEverywhere();
            if (forceEnable) {
              try {
                for (const form of Array.from(document.querySelectorAll('form'))) {
                  form.noValidate = true;
                  form.setAttribute('novalidate', 'novalidate');
                }
                const mobileInputs = Array.from(document.querySelectorAll('input')).filter(el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  if (r.width <= 0 || r.height <= 0 || s.visibility === 'hidden' || s.display === 'none') return false;
                  const info = norm(`${el.placeholder || ''} ${el.name || ''} ${el.id || ''} ${el.getAttribute('aria-label') || ''}`);
                  return info.includes('mobile') || info.includes('phone') || info.includes('contact') || info.includes('enter mobile');
                });
                for (const input of mobileInputs) {
                  input.required = false;
                  input.removeAttribute('required');
                  input.removeAttribute('pattern');
                  input.removeAttribute('minlength');
                  input.removeAttribute('maxlength');
                  input.setCustomValidity('');
                  const setter = Object.getOwnPropertyDescriptor(input.__proto__, 'value')?.set || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
                  if (setter) setter.call(input, '');
                  else input.value = '';
                  input.classList.remove('ng-invalid', 'is-invalid', 'invalid');
                  input.classList.add('ng-valid');
                  input.setAttribute('aria-invalid', 'false');
                  input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
                  input.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
                  input.dispatchEvent(new Event('blur', { bubbles: true, composed: true }));
                }
              } catch(e) {}
            }
            const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], [role="button"], a'))
              .filter(visible)
              .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}));
            if (!buttons.length) return false;
            const preferred = buttons.find(x => x.text === 'confirm & submit' || x.text === 'confirm and submit')
              || buttons.find(x => x.text.includes('confirm') && x.text.includes('submit'))
              || buttons.find(x => ['confirm', 'ok', 'yes', 'submit'].some(label => x.text === label));
            if (!preferred) return false;
            const btn = preferred.el;
            if (forceEnable) {
              try { btn.disabled = false; } catch(e) {}
              try { btn.removeAttribute('disabled'); } catch(e) {}
              try { btn.removeAttribute('aria-disabled'); } catch(e) {}
              try { btn.classList.remove('disabled', 'is-disabled'); } catch(e) {}
            }
            btn.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'});
            let r = btn.getBoundingClientRect();
            if (r.top < 0 || r.bottom > window.innerHeight) {
              try { window.scrollTo(0, Math.max(0, window.scrollY + r.top - Math.floor(window.innerHeight / 2))); } catch(e) {}
              btn.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'});
              r = btn.getBoundingClientRect();
            }
            if (!visible(btn) || r.bottom < 0 || r.top > window.innerHeight) return false;
            const cx = Math.max(1, Math.min(window.innerWidth - 2, r.left + r.width / 2));
            const cy = Math.max(1, Math.min(window.innerHeight - 2, r.top + r.height / 2));
            const topEl = document.elementFromPoint(cx, cy);
            const clickable = topEl?.closest?.('button, input[type="button"], input[type="submit"], [role="button"], a') || btn;
            if (forceEnable) {
              for (const el of [btn, clickable]) {
                try { el.disabled = false; } catch(e) {}
                try { el.removeAttribute('disabled'); } catch(e) {}
                try { el.removeAttribute('aria-disabled'); } catch(e) {}
                try { el.classList.remove('disabled', 'is-disabled'); } catch(e) {}
              }
            }
            const target = clickable?.matches?.('button, input[type="button"], input[type="submit"], [role="button"], a')
              ? clickable : btn;
            try { target.focus?.(); } catch(e) {}
            target.click();
            return true;
            """,
            force_enable,
        )
    )


def process_screening_health_rows_browser(driver, rows: list[dict[str, Any]], args: argparse.Namespace, label: str = "Screening Health Entry") -> dict[str, Any]:
    print(f"\n{label} selected rows: {len(rows)}")
    completed = load_progress(Path(args.progress)) if args.resume else set()
    success_headers = ["Sr No", "Student Full Name", "message"]
    failed_headers = ["Sr No", "Student Full Name", "error", "portal_response"]
    speed_mode = getattr(args, "speed_mode", "")
    month = clean(getattr(args, "screening_month", ""))
    if not month and bool(getattr(args, "screening_page_prepared", False)):
        month = browser_selected_screening_month(driver)
        if not month:
            raise RuntimeError("Month is not selected on the Screening page. Select the Month, then start again.")
    institution_type = clean(getattr(args, "screening_institution_type", "")) or "School"
    is_awc = is_awc_institution(institution_type)
    browser_prepare_screening_search(
        driver,
        month=month,
        institution_type=institution_type,
        speed_mode=speed_mode,
        page_already_prepared=bool(getattr(args, "screening_page_prepared", False)),
    )
    saved = failed = skipped = 0
    current_institute = ""
    mobile_source = clean(getattr(args, "mobile_source", "excel"))
    custom_school_mobile = clean(getattr(args, "custom_school_mobile", ""))
    school_mobiles: dict[str, str] = {}
    pending = [row for row in rows if not (args.resume and clean(row.get("Sr No")) in completed)]
    for index, row in enumerate(pending, start=1):
        enforce_entry_license(args)
        row["_is_awc"] = is_awc
        sr_no = clean(row.get("Sr No")) or str(index)
        name = screening_student_name(row)
        institute = screening_institute_name(row)
        if is_awc:
            row["_school_custom_mobile"] = ""
        elif mobile_source == "portal_school":
            school_key = normalize_match_text(institute) or "default"
            if school_key not in school_mobiles:
                if custom_school_mobile and len("".join(ch for ch in custom_school_mobile if ch.isdigit())) >= 10:
                    school_mobiles[school_key] = "".join(ch for ch in custom_school_mobile if ch.isdigit())[-10:]
                else:
                    callback = getattr(args, "prompt_mobile_callback", None)
                    asked_num = callback(institute or "Current School") if callback else ask_school_mobile_console(institute or "Current School")
                    digits = "".join(ch for ch in clean(asked_num) if ch.isdigit())
                    school_mobiles[school_key] = digits[-10:] if len(digits) >= 10 else "9999999999"
            row["_school_custom_mobile"] = school_mobiles[school_key]

        print(f"[{index}/{len(pending)}] Screening health entry: Sr No {sr_no} {name}")
        try:
            if institute and normalize_match_text(institute) != normalize_match_text(current_institute):
                browser_open_screening_school(driver, institute, speed_mode=speed_mode)
                current_institute = institute
            browser_open_student_health_form(driver, row, speed_mode=speed_mode)
            # Safety gate: confirm health form is visible before filling
            if not browser_is_screening_health_form_visible(driver):
                _form_deadline = time.time() + speed_timeout(speed_mode, 10.0, 18.0)
                while time.time() < _form_deadline:
                    if browser_is_screening_health_form_visible(driver):
                        break
                    browser_sleep(0.5)
                if not browser_is_screening_health_form_visible(driver):
                    raise RuntimeError(f"Health form not visible for '{name}' — Start Screening may not have opened.")
            browser_fill_screening_health_form(driver, row, speed_mode=speed_mode)
            browser_preview_submit_screening(driver, row, speed_mode=speed_mode)
            completed.add(sr_no)
            save_progress(Path(args.progress), completed)
            append_csv(Path(args.success_log), {"Sr No": sr_no, "Student Full Name": name, "message": "Screening health submitted"}, success_headers)
            saved += 1
            record_licensed_success(args)
            try:
                location = browser_return_to_screening_list_after_save(driver, row, speed_mode=speed_mode)
                current_institute = institute if location == "student_list" else ""
                if location == "unknown":
                    print("Saved row, but the tool could not confirm the list page. Next row will re-open the school from Screening.")
            except Exception as recover_exc:
                current_institute = ""
                print(f"Saved row, but return-to-list recovery failed: {recover_exc}")
            browser_sleep(float(getattr(args, "delay", 0.25) or 0.25))
        except Exception as exc:
            error_text = str(exc)
            if browser_screening_saved_notice_visible(driver):
                print(f"RBSK reported saved successfully after a wait issue for Sr No {sr_no}; marking completed.")
                completed.add(sr_no)
                save_progress(Path(args.progress), completed)
                append_csv(
                    Path(args.success_log),
                    {"Sr No": sr_no, "Student Full Name": name, "message": "Screening health submitted"},
                    success_headers,
                )
                saved += 1
                record_licensed_success(args)
                try:
                    location = browser_return_to_screening_list_after_save(driver, row, speed_mode=speed_mode)
                    current_institute = institute if location == "student_list" else ""
                except Exception:
                    current_institute = ""
                continue
            failed += 1
            append_csv(Path(args.failed_log), {"Sr No": sr_no, "Student Full Name": name, "error": error_text, "portal_response": ""}, failed_headers)
            try:
                screenshot_path = Path(args.failed_log).parent / f"screening_failure_sr_{sr_no}.png"
                driver.save_screenshot(str(screenshot_path))
            except Exception:
                pass
            print(f"FAILED Sr No {sr_no}: {error_text}")
            try:
                if browser_screening_school_list_visible(driver):
                    current_institute = ""
                elif browser_screening_student_list_has_name(driver, name):
                    current_institute = institute
                else:
                    driver.back()
                    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 16.0))
                    if browser_screening_student_list_has_name(driver, name):
                        current_institute = institute
                    elif browser_screening_school_list_visible(driver):
                        current_institute = ""
                    else:
                        current_institute = ""
                        print("Could not identify the recovery page. Month/Search was not repeated automatically.")
            except Exception:
                current_institute = ""
    skipped = len(rows) - len(pending)
    return {"saved": saved, "duplicates": 0, "failed": failed, "skipped": skipped, "skipped_this_run": 0, "exit_code": 0 if failed == 0 else 1}


def browser_get_active_school_name(driver) -> str:
    try:
        return clean(
            driver.execute_script(
                """
                const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
                const sel = document.querySelector('select[name*="school"], select[name*="institute"], select[id*="school"], select[id*="institute"]');
                if (sel && sel.selectedOptions && sel.selectedOptions[0]) {
                  const txt = norm(sel.selectedOptions[0].textContent);
                  if (txt && !/select/i.test(txt)) return txt;
                }
                const allElements = Array.from(document.querySelectorAll('td, th, span, label, b, strong, p, h1, h2, h3, h4, h5, div'));
                for (const el of allElements) {
                  if (el.children.length > 2) continue;
                  const txt = norm(el.innerText || el.textContent || '');
                  const match = txt.match(/^Institute\\s*:\\s*(.+)/i);
                  if (match && match[1] && match[1].trim().length > 0 && match[1].trim().length < 120) {
                    return match[1].trim();
                  }
                }
                for (const el of allElements) {
                  const txt = norm(el.innerText || el.textContent || '');
                  if (txt.toLowerCase() === 'institute:') {
                    const sibling = el.nextElementSibling || el.parentElement?.nextElementSibling;
                    if (sibling) {
                      const stxt = norm(sibling.innerText || sibling.textContent || '');
                      if (stxt && stxt.length < 120) return stxt;
                    }
                  }
                }
                return '';
                """
            )
        )
    except Exception:
        return ""


def browser_active_institution_is_awc(driver, fallback: bool = False) -> bool:
    """Detect whether the currently opened student list belongs to an AWC/AWS."""
    try:
        result = driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden'
                && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const headers = Array.from(document.querySelectorAll('th, [role="columnheader"]'))
              .filter(visible)
              .map(el => norm(el.innerText || el.textContent || ''));
            if (headers.some(text => text.includes('awc child id'))) return 'awc';
            if (headers.some(text => text.includes('student id') || text.includes('student name'))) return 'school';

            const labels = Array.from(document.querySelectorAll('label, span, b, strong, h1, h2, h3, h4, p'))
              .filter(visible)
              .map(el => norm(el.innerText || el.textContent || ''))
              .filter(text => text.length > 0 && text.length < 100);
            if (labels.some(text => /^institute type\\s*:\\s*(awc|aws)\\b/.test(text))) return 'awc';
            if (labels.some(text => /^institute type\\s*:\\s*(school|college)\\b/.test(text))) return 'school';
            return 'unknown';
            """
        )
        if result == "awc":
            return True
        if result == "school":
            return False
    except Exception:
        pass
    return bool(fallback)


def browser_handle_not_eligible_modal(driver) -> bool:
    """Detect and dismiss 'Not Eligible' modal popups (e.g. Children aged 6 years or above in AWC or Students aged 19 years or above in School)."""
    try:
        return bool(
            driver.execute_script(
                """
                const norm = s => String(s || '').toLowerCase().replace(/[\\s\\u00a0]+/g, ' ').trim();
                const text = norm(document.body.innerText || document.body.textContent || '');
                if (text.includes('not eligible') || text.includes('19 years or above') || text.includes('6 years or above') || text.includes('not eligible for screening')) {
                    const buttons = Array.from(document.querySelectorAll('button, .swal2-confirm, .swal-button--confirm, [role="button"], a, input[type="button"]'))
                      .filter(el => {
                        const r = el.getBoundingClientRect();
                        const s = getComputedStyle(el);
                        return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
                      });
                    const okBtn = buttons.find(el => {
                      const t = norm(el.innerText || el.textContent || el.value || '');
                      return t === 'ok' || t.includes('ok') || t === 'close' || t === 'dismiss';
                    }) || buttons[0];
                    if (okBtn) {
                      okBtn.click();
                      return true;
                    }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def browser_find_and_click_next_unscreened_student(driver, is_awc: bool = False, skipped_names: list[str] | None = None) -> dict[str, str] | None:
    # Schools and AWCs can show the same Plan Completed prompt before student selection.
    browser_handle_plan_completed_modal(driver)
    # First check if 'Not Eligible' modal popup is on page and dismiss it
    if browser_handle_not_eligible_modal(driver):
        print("Dismissed 'Not Eligible' popup from previous action.")
        browser_sleep(0.5)

    skipped_list = [clean(s) for s in (skipped_names or []) if clean(s)]

    try:
        res = driver.execute_script(
            """
            const isAwcMode = arguments[0];
            const skippedList = arguments[1] || [];
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/[\\s\\u00a0]+/g, ' ').trim();

            // Clear any search filter box on table if active
            const searchInputs = document.querySelectorAll('input[type="search"], input[placeholder*="Search"], input[placeholder*="search"], .dataTables_filter input');
            for (const input of searchInputs) {
              if (input.value) {
                input.value = '';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
              }
            }

            const pageText = norm(document.body.innerText || document.body.textContent || '');
            const isAwcPage = isAwcMode || pageText.includes('awc') || pageText.includes('anganwadi');

            const rows = Array.from(document.querySelectorAll('tbody tr, tr, [role="row"]')).filter(visible);
            
            for (const tr of rows) {
              const cells = Array.from(tr.querySelectorAll('td, th, [role="cell"]')).filter(visible);
              if (cells.length < 2) continue;
              const text = norm(tr.innerText || tr.textContent);
                     // Skip header rows
              if (/^\\s*(sr\\.?\\s*no|s\\.?\\s*no|name|action)/.test(text)) continue;
              
              // Skip rows that already have a health status. Action titles can say
              // "Screening not done yet", so do not use broad row text for "done".
              if (/completed|screened|healthy|identified|unhealthy|referral|referred|locked|restriction|closed|already/.test(text)) continue;
              
              // Verify cell texts: if any cell says 'healthy', 'unhealthy', 'completed', 'referral', skip
              const isScreenedCell = cells.some(c => {
                const t = norm(c.innerText || c.textContent);
                return /^(healthy|identified|unhealthy|referral|referred|completed|screened|done)$/i.test(t);
              });
              if (isScreenedCell) continue;

              const nameCell = cells.find(c => {
                const t = norm(c.innerText || c.textContent);
                return t.length > 2 && !/\\d{5,}/.test(t) && !/start|action|male|female|select|view|edit/i.test(t);
              }) || cells[1];
              const studentName = (nameCell ? nameCell.innerText || nameCell.textContent : 'Student').trim();

              // SKIP IF STUDENT IS IN BLACKLIST / SKIPPED SET
              if (skippedList.some(s => norm(s) === norm(studentName) || norm(studentName).includes(norm(s)) || norm(s).includes(norm(studentName)))) {
                continue;
              }
              
              const ageCell = cells.find(c => /years|months/i.test(c.innerText || c.textContent || ''));
              const ageText = ageCell ? (ageCell.innerText || ageCell.textContent).trim() : '';

              // Skip ineligible students by age:
              // - In AWC mode, children aged 6 years or above are NOT eligible for screening in AWC
              // - In School mode, students aged 19 years or above are NOT eligible for screening in School
              if (ageText) {
                const yMatch = norm(ageText).match(/(\\d+)\\s*years?/i);
                if (yMatch) {
                  const years = parseInt(yMatch[1], 10);
                  if (isAwcPage && years >= 6) continue;
                  if (!isAwcPage && years >= 19) continue;
                }
              }

              // Collect all clickable action elements in this row
              const allActions = Array.from(tr.querySelectorAll('[title], button, a, [role="button"], i, svg, span, img'))
                .filter(visible)
                .map(el => el.closest('button, a, [role="button"]') || el)
                .filter((el, i, arr) => arr.indexOf(el) === i)
                .filter(el => {
                  if (el.disabled || el.getAttribute('aria-disabled') === 'true') return false;
                  const info = norm(`${el.title||''} ${el.getAttribute('aria-label')||''} ${el.className||''} ${el.innerText||''} ${el.textContent||''}`);
                  // Filter out completed checkmarks / view buttons / badges (Do NOT match 'success' as buttons use btn-success)
                  if (/completed|edit screening|reset screening|cannot deactivate|screening not done|fa-check|check-circle|bi-check|icon-check/.test(info)) return false;
                  // If it's literally a checkmark badge or view button, skip
                  if (/^(view|details|check|done)$/i.test(info)) return false;
                  return true;
                });
              
              // Priority 1: stethoscope icon / stethoscope class / title containing "start" or "screen" or "stethoscope"
              let startBtn = allActions.find(el => {
                const info = norm(`${el.getAttribute('title')||''} ${el.getAttribute('aria-label')||''} ${el.className||''} ${el.innerHTML||''}`);
                return /start screening|stethoscope/i.test(info) && !/edit|reset|\\bcheck\\b|\\bview\\b|completed|identified|healthy/i.test(info);
              });
              
              // Priority 2: icon classes matching medical/stethoscope icons
              if (!startBtn) startBtn = allActions.find(el => {
                const cls = (el.className || '') + ' ' + Array.from(el.querySelectorAll('i, svg, span')).map(c => c.className || '').join(' ');
                  const info = norm(`${el.getAttribute('title')||''} ${el.getAttribute('aria-label')||''} ${cls}`);
                  return /fa-stethoscope|stethoscope|fa-user-md|fa-medkit/i.test(info) && !/edit|reset|completed|identified|healthy/i.test(info);
              });
              
              // Do not fall back to arbitrary actions: ABHA/PEN/edit/reset controls live in the same cell.
              
              if (startBtn) {
                const nameCell = cells.find(c => {
                  const t = norm(c.innerText || c.textContent);
                  return t.length > 2 && !/\\d{5,}/.test(t) && !/start|action|male|female|select|view|edit/i.test(t);
                }) || cells[1];
                const studentName = (nameCell ? nameCell.innerText || nameCell.textContent : 'Student').trim();
                
                // Extract mobile number from the table row (10-digit number)
                let mobileNumber = '';
                for (const c of cells) {
                  const ct = (c.innerText || c.textContent || '').trim();
                  const digits = ct.replace(/[^0-9]/g, '');
                  if (digits.length === 10 && /^[6-9]/.test(digits)) {
                    mobileNumber = digits;
                    break;
                  }
                }
                
                const clickable = startBtn.closest('button, a, [role="button"]') || startBtn;
                try {
                  let ancestor = clickable.parentElement;
                  while (ancestor && ancestor !== document.body) {
                    const cs = getComputedStyle(ancestor);
                    if (cs.overflow === 'hidden' || cs.overflowY === 'hidden') ancestor.style.overflow = 'auto';
                    ancestor = ancestor.parentElement;
                  }
                } catch(e) {}
                try { clickable.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'}); } catch(e) {}
                
                return {status: 'found', el: clickable, name: studentName, ageText: ageText, mobile: mobileNumber};
              }
            }
            return {status: 'none'};
            """,
            is_awc,
        )
        if isinstance(res, dict) and res.get("status") == "found":
            target_el = res.get("el")
            name = clean(res.get("name")) or "Student"
            age_text = clean(res.get("ageText"))
            mobile = clean(res.get("mobile"))
            
            # Perform 3-tier click execution
            click_done = False
            if target_el:
                try:
                    driver.execute_script(
                        """
                        const el = arguments[0];
                        const r = el.getBoundingClientRect();
                        const opts = {bubbles: true, cancelable: true, view: window, clientX: Math.max(1, r.left + r.width / 2), clientY: Math.max(1, r.top + r.height / 2)};
                        for (const t of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                          try { el.dispatchEvent(new MouseEvent(t, opts)); } catch(e) {}
                        }
                        try { el.click(); } catch(e) {}
                        """,
                        target_el,
                    )
                    click_done = True
                except Exception:
                    pass
                if not click_done:
                    try:
                        target_el.click()
                        click_done = True
                    except Exception:
                        pass
                if not click_done:
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).move_to_element(target_el).click().perform()
                    except Exception:
                        pass

            print(f"Clicked Start Screening for: {name}" + (f" (Mobile: {mobile})" if mobile else ""))
            return {"name": name, "ageText": age_text, "mobile": mobile}
        return None
    except Exception as exc:
        print(f"Error finding unscreened student: {exc}")
        return None


def browser_click_next_student_page(driver, speed_mode: Any = "") -> bool:
    try:
        clicked = bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const candidates = Array.from(document.querySelectorAll('button, a, [role="button"], li, span')).filter(visible);
                const nextBtn = candidates.find(el => {
                  const text = norm(el.innerText || el.textContent || el.getAttribute('aria-label') || el.title || '');
                  const isNext = (text === 'next' || text === 'next >' || text.endsWith(' next') || text === '>') || el.classList.contains('next') || el.classList.contains('pagination-next');
                  const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true' || el.classList.contains('disabled') || el.parentElement?.classList.contains('disabled');
                  return isNext && !disabled;
                });
                if (!nextBtn) return false;
                nextBtn.scrollIntoView({block: 'center', inline: 'center'});
                nextBtn.click();
                return true;
                """
            )
        )
        if clicked:
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
            browser_sleep(speed_wait(speed_mode, 0.4, 0.8))
            return True
        return False
    except Exception:
        return False


def browser_find_and_click_next_inprogress_school(driver) -> dict[str, str] | None:
    try:
        res = driver.execute_script(
            """
            const visible = el => {
              const r = el.getBoundingClientRect();
              const s = getComputedStyle(el);
              return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0';
            };
            const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const rows = Array.from(document.querySelectorAll('tbody tr, tr')).filter(visible);
            
            for (const tr of rows) {
              const cells = Array.from(tr.querySelectorAll('td, th')).filter(visible);
              if (cells.length < 8) continue;
              
              const cellTexts = cells.map(c => norm(c.innerText || c.textContent || ''));
              
              // Only match rows containing "screening inprogress" or "screening in progress"
              const isInprogress = cellTexts.some(txt => txt.includes('screening inprogress') || txt.includes('screening in progress'));
              if (isInprogress) {
                const schoolNameCell = cells.find(c => {
                  const t = norm(c.innerText || c.textContent);
                  return (t.includes('school') || t.includes('college') || t.includes('vidyalaya') || t.includes('awc') || t.includes('primary')) && !t.includes('visit') && !t.includes('screen');
                }) || cells[5];
                
                const schoolName = (schoolNameCell ? schoolNameCell.innerText || schoolNameCell.textContent : 'School').trim();
                
                const actions = Array.from(tr.querySelectorAll('button, a, [role="button"], i, svg, span'))
                  .filter(visible)
                  .map(el => el.closest('button, a, [role="button"]') || el)
                  .filter((el, i, arr) => arr.indexOf(el) === i);
                  
                const viewBtn = actions.find(el => {
                  const t = norm(el.title || el.getAttribute('aria-label') || el.className || el.innerText || el.textContent || '');
                  return t.includes('view') || t.includes('eye') || t.includes('show') || el.querySelector('i.fa-eye, svg') || el.classList.contains('btn-success');
                }) || actions[0];
                
                if (viewBtn) {
                  viewBtn.scrollIntoView({block: 'center', inline: 'center'});
                  viewBtn.click();
                  return {status: 'clicked', schoolName: schoolName};
                }
              }
            }
            return {status: 'none'};
            """
        )
        if isinstance(res, dict) and res.get("status") == "clicked":
            return {"schoolName": clean(res.get("schoolName")) or "School"}
        return None
    except Exception:
        return None


def browser_return_to_school_list(driver) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
                const visible = el => {
                  const r = el.getBoundingClientRect();
                  const s = getComputedStyle(el);
                  return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' && !el.disabled;
                };
                const norm = s => String(s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
                const buttons = Array.from(document.querySelectorAll('button, [role="button"], a, input[type="button"], span, div, li'))
                  .filter(visible)
                  .map(el => ({el, text: norm(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || '')}))
                  .filter(x => x.text === 'back' || x.text === '← back' || x.text.includes('back to school') || x.text.includes('back to list') || x.text === 'screening');
                if (!buttons.length) return false;
                buttons[0].el.scrollIntoView({block: 'center', inline: 'center'});
                buttons[0].el.click();
                return true;
                """
            )
        )
    except Exception:
        return False


def run_direct_portal_screening(driver, args: argparse.Namespace) -> dict[str, Any]:
    speed_mode = getattr(args, "speed_mode", "Fast Safe")
    institution_type = clean(getattr(args, "screening_institution_type", "")) or "School"
    configured_is_awc = is_awc_institution(institution_type)
    active_is_awc = configured_is_awc
    if not browser_is_open(driver):
        print("RBSK browser was closed. Stopping without showing any popups.")
        return {"saved": 0, "failed": 0, "exit_code": 0, "browser_closed": True}
    if configured_is_awc:
        print("\nStarting Direct Portal Screening Mode for AWC (No Excel File Required)...")
        print("Institution Type: AWC/AWS - Mobile number will NOT be asked or entered.")
    else:
        print("\nStarting Direct Portal Screening Mode (No Excel File Required)...")

    month = clean(getattr(args, "screening_month", ""))
    if not month and bool(getattr(args, "screening_page_prepared", False)):
        month = browser_selected_screening_month(driver)

    try:
        browser_prepare_screening_search(
            driver,
            month=month,
            institution_type=institution_type,
            speed_mode=speed_mode,
            page_already_prepared=bool(getattr(args, "screening_page_prepared", False)),
        )
    except Exception:
        if not browser_is_open(driver):
            print("RBSK browser was closed. Stopping without showing any popups.")
            return {"saved": 0, "failed": 0, "exit_code": 0, "browser_closed": True}
        raise

    saved = failed = 0
    custom_school_mobile = clean(getattr(args, "custom_school_mobile", ""))
    active_mobile = ""
    skipped_names: set[str] = set()

    while True:
        if not browser_is_open(driver):
            print("RBSK browser was closed. Stopping without showing any more popups.")
            break

        # Check if session ended / logged out
        browser_check_and_wait_if_logged_out(driver)
        browser_handle_plan_completed_modal(driver)

        # State 1: School/AWC List Page
        if browser_screening_school_list_visible(driver):
            if not browser_is_open(driver):
                break
            if configured_is_awc:
                show_gui_message(
                    "Select AWC",
                    "Month search is complete.\n\n"
                    "1. In the browser, click the AWC you want to screen.\n"
                    "2. Wait until that AWC's child list is visible.\n"
                    "3. Click OK here.\n\n"
                    "Mobile number will NOT be asked or entered."
                )
            else:
                show_gui_message(
                    "Select School / AWC",
                    "Month search is complete.\n\n"
                    "1. In the browser, click the school, college, or AWC you want to screen.\n"
                    "2. Wait until its student/child list is visible.\n"
                    "3. Click OK here.\n\n"
                    "AWC/AWS is detected automatically and continues without mobile. "
                    "For School/College, the existing mobile-number operation continues."
                )
            browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
            browser_handle_plan_completed_modal(driver)
            if not browser_screening_student_list_visible(driver) and not browser_is_screening_health_form_visible(driver):
                if not browser_is_open(driver):
                    break
                show_gui_message(
                    "Student List Not Detected",
                    "The student list is not visible yet.\n\n"
                    "Please click/open one school row in the browser, wait for the student list, then click OK here."
                )
                browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
                if not browser_screening_student_list_visible(driver) and not browser_is_screening_health_form_visible(driver):
                    print("Waiting for school selection. No student list detected yet.")
                    continue

            active_is_awc = browser_active_institution_is_awc(driver, fallback=configured_is_awc)
            active_school = browser_get_active_school_name(driver) or ("Current AWC" if active_is_awc else "Current School/College")
            if active_is_awc:
                active_mobile = ""
                print(f"AWC/AWS detected: continuing {active_school} without asking or entering a mobile number.")
            elif custom_school_mobile and len("".join(ch for ch in custom_school_mobile if ch.isdigit())) >= 10:
                active_mobile = "".join(ch for ch in custom_school_mobile if ch.isdigit())[-10:]
            else:
                if not browser_is_open(driver):
                    break
                asked = browser_prompt_school_mobile(driver, active_school)
                if not browser_is_open(driver):
                    break
                if not asked:
                    print("School mobile entry was cancelled. Stopping direct screening.")
                    break
                digits = "".join(ch for ch in clean(asked) if ch.isdigit())
                active_mobile = digits[-10:] if len(digits) >= 10 else "9999999999"
            if not active_is_awc:
                print(f"Using mobile number {active_mobile} for {active_school}. Starting filling process.")
            browser_handle_mark_attendance_modal(driver)
            browser_sleep(speed_wait(speed_mode, 0.2, 0.6))
            continue

        # State 2: Health Form Page already open
        if browser_is_screening_health_form_visible(driver):
            enforce_entry_license(args)
            active_school = browser_get_active_school_name(driver) or ("Current AWC" if active_is_awc else "Current School/College")
            if not active_mobile:
                if active_is_awc:
                    active_mobile = ""
                elif custom_school_mobile and len("".join(ch for ch in custom_school_mobile if ch.isdigit())) >= 10:
                    active_mobile = "".join(ch for ch in custom_school_mobile if ch.isdigit())[-10:]
                else:
                    if not browser_is_open(driver):
                        break
                    asked = browser_prompt_school_mobile(driver, active_school)
                    if not browser_is_open(driver):
                        break
                    if not asked:
                        print("School mobile entry was cancelled. Stopping direct screening.")
                        break
                    digits = "".join(ch for ch in clean(asked) if ch.isdigit())
                    active_mobile = digits[-10:] if len(digits) >= 10 else "9999999999"
            
            student_name_on_page = driver.execute_script("""
                const text = document.body.innerText || document.body.textContent || '';
                const match = text.match(/Child\\s*Name\\s*:\\s*([^\\n]+)/i) || text.match(/Screening\\s*for\\s*:\\s*([^\\n]+)/i);
                return match ? match[1].trim() : 'Current Student';
            """)
            print(f"Health Form is open for: {student_name_on_page}")
            row = {"Student Full Name": student_name_on_page, "_school_custom_mobile": active_mobile, "_is_awc": active_is_awc}
            try:
                browser_fill_screening_health_form(driver, row, speed_mode=speed_mode)
                browser_uncheck_other_health_condition(driver)
                browser_preview_submit_screening(driver, row, speed_mode=speed_mode)
                saved += 1
                record_licensed_success(args)
                location = browser_return_to_screening_list_after_save(driver, row, speed_mode=speed_mode)
                if location == "school_list":
                    active_mobile = ""
                    active_is_awc = configured_is_awc
                continue
            except Exception as exc:
                if browser_is_logged_out(driver):
                    browser_check_and_wait_if_logged_out(driver)
                    continue
                failed += 1
                if student_name_on_page:
                    skipped_names.add(student_name_on_page)
                print(f"Direct Portal Screening error for {student_name_on_page}: {exc}. Added to skip list.")
                browser_handle_not_eligible_modal(driver)
                try:
                    driver.back()
                except Exception:
                    pass
                browser_sleep(speed_wait(speed_mode, 0.5, 1.2))
                continue

        # State 3: Student List Page -> Pick next unscreened student
        detected_is_awc = browser_active_institution_is_awc(driver, fallback=active_is_awc)
        if detected_is_awc and not active_is_awc:
            active_mobile = ""
        active_is_awc = detected_is_awc
        active_school = browser_get_active_school_name(driver) or ("Current AWC" if active_is_awc else "Current School/College")
        if not active_mobile:
            if active_is_awc:
                active_mobile = ""
            elif custom_school_mobile and len("".join(ch for ch in custom_school_mobile if ch.isdigit())) >= 10:
                active_mobile = "".join(ch for ch in custom_school_mobile if ch.isdigit())[-10:]
            else:
                if not browser_is_open(driver):
                    break
                asked = browser_prompt_school_mobile(driver, active_school)
                if not browser_is_open(driver):
                    break
                if not asked:
                    print("School mobile entry was cancelled. Stopping direct screening.")
                    break
                digits = "".join(ch for ch in clean(asked) if ch.isdigit())
                active_mobile = digits[-10:] if len(digits) >= 10 else "9999999999"
            if not active_is_awc:
                print(f"Using mobile number {active_mobile} for {active_school}. Starting filling process.")

        student_info = browser_find_and_click_next_unscreened_student(driver, is_awc=active_is_awc, skipped_names=list(skipped_names))
        if not student_info:
            has_next = browser_click_next_student_page(driver, speed_mode=speed_mode)
            if has_next:
                continue
            else:
                print(f"All unscreened students completed for {active_school}.")
                active_mobile = ""
                active_is_awc = configured_is_awc
                returned = browser_return_to_school_list(driver)
                if not returned:
                    try:
                        driver.back()
                    except Exception:
                        pass
                browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
                browser_sleep(speed_wait(speed_mode, 0.25, 0.6))
                if not browser_screening_school_list_visible(driver):
                    if not browser_is_open(driver):
                        break
                    show_gui_message(
                        "School Screening Complete",
                        f"Completed screening for all available students in '{active_school}'.\n\n"
                        "To screen another school:\n"
                        "1. In the browser, click the Back link or select another school from the list.\n"
                        "2. Click OK here to continue."
                    )
                    browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 6.0, 12.0))
                continue

        name = student_info.get("name", "Student")
        age_text = student_info.get("ageText", "")
        student_mobile = student_info.get("mobile", "")
        # Use student's own mobile from the table row if available, otherwise fall back to school mobile
        effective_mobile = "" if active_is_awc else (student_mobile if student_mobile and len(student_mobile) == 10 else active_mobile)
        mobile_note = "No mobile" if active_is_awc else f"Mobile: {effective_mobile}"
        print(f"Direct Portal Screening for: {name} (Age: {age_text}) (School: {active_school}) ({mobile_note})")
        row = {"Student Full Name": name, "ageText": age_text, "_school_custom_mobile": effective_mobile, "_is_awc": active_is_awc}
        if not active_is_awc:
            row["Mobile Number"] = effective_mobile

        # Check if clicking caused 'Not Eligible' modal popup (e.g. child >= 6 years in AWC or student >= 19 years in School)
        browser_sleep(speed_wait(speed_mode, 0.15, 0.4))
        if browser_handle_not_eligible_modal(driver):
            skipped_names.add(name)
            print(f"[Not Eligible] {name} (Age: {age_text}) is NOT eligible for screening. Added to skip list and skipped!")
            browser_sleep(speed_wait(speed_mode, 0.2, 0.5))
            continue

        # Wait for health form to load after clicking Start Screening
        browser_wait_for_portal_idle(driver, timeout=speed_timeout(speed_mode, 8.0, 16.0))
        browser_sleep(speed_wait(speed_mode, 0.08, 1.0))
        enforce_entry_license(args)

        try:
            browser_fill_screening_health_form(driver, row, speed_mode=speed_mode)
            browser_uncheck_other_health_condition(driver)
            browser_preview_submit_screening(driver, row, speed_mode=speed_mode)
            saved += 1
            record_licensed_success(args)
            location = browser_return_to_screening_list_after_save(driver, row, speed_mode=speed_mode)
            if location == "student_list":
                print("Returned to student list. Continuing next student entry.")
                continue
            if location == "school_list":
                active_mobile = ""
                active_is_awc = configured_is_awc
                print("Returned to school list. Select the next school to continue.")
                continue
            print("Saved entry, but student list was not confirmed. Trying to continue from the current page.")
        except Exception as exc:
            if browser_is_logged_out(driver):
                browser_check_and_wait_if_logged_out(driver)
                continue
            failed += 1
            if name:
                skipped_names.add(name)
            print(f"Direct Portal Screening error for {name}: {exc}. Added to skip list and continuing next student.")
            browser_handle_not_eligible_modal(driver)
            try:
                driver.back()
            except Exception:
                pass
            browser_sleep(speed_wait(speed_mode, 0.5, 1.2))

    return {"saved": saved, "failed": failed, "exit_code": 0 if failed == 0 else 1}


SKIPPED_HEADERS = ["Sr No", "Student Full Name", "AWC NAME", "status", "reason", "suggested_action"]


def skipped_log_path(args: argparse.Namespace) -> Path:
    value = clean(getattr(args, "skipped_log", ""))
    if value:
        return Path(value)
    return Path(args.failed_log).parent / "skipped.csv"


def append_skipped_row(
    args: argparse.Namespace,
    row: dict[str, Any],
    status: str,
    reason: str,
    suggested_action: str = "",
    institute_name: str = "",
) -> None:
    append_csv(
        skipped_log_path(args),
        {
            "Sr No": clean(row.get("Sr No")),
            "Student Full Name": clean(row.get("Student Full Name")),
            "AWC NAME": clean(institute_name) or clean(row.get("Institute Name")),
            "status": status,
            "reason": reason,
            "suggested_action": suggested_action,
        },
        SKIPPED_HEADERS,
    )


def row_valid_mobile(row: dict[str, Any]) -> str:
    digits = only_digits(row.get("Mobile Number"), 20)
    return digits[:10] if len(digits) >= 10 else ""


def process_student_rows_browser_manual_countdown(driver, rows: list[dict[str, Any]], args: argparse.Namespace, label: str) -> dict[str, Any]:
    print(f"\n{label} selected rows: {len(rows)}")
    completed = load_progress(Path(args.progress)) if args.resume else set()
    groups, skipped_count = browser_pending_institute_groups(rows, completed, args.resume)
    success_headers = ["Sr No", "Student Full Name", "portal_student_code", "message"]
    failed_headers = ["Sr No", "Student Full Name", "error", "portal_response"]
    saved_count = failed_count = duplicate_count = skipped_this_run = 0
    duplicate_skip_remaining = 0
    last_valid_mobile_by_awc: dict[str, str] = {}
    unattended = bool(getattr(args, "unattended_after_start", False))
    pending_total = sum(len(group["rows"]) for group in groups)
    if pending_total == 0:
        print("No pending student rows. All selected rows are already completed.")
        emit_progress(args, "batch_complete", label=label, saved=0, failed=0, skipped=skipped_count, pending=0)
        return {"exit_code": 0, "saved": saved_count, "duplicates": duplicate_count, "failed": failed_count, "skipped": skipped_count, "skipped_this_run": skipped_this_run}

    auto_select_institute = bool(getattr(args, "auto_select_institute", False))
    speed_mode = clean(getattr(args, "speed_mode", "")) or "Fast Safe"
    institution_type = clean(getattr(args, "screening_institution_type", "")) or clean(getattr(args, "institution_type", ""))
    is_awc_mode = is_awc_institution(institution_type)
    same_awc_open_timeout = speed_timeout(speed_mode, 8.0, 20.0)
    initial_same_awc_open_timeout = speed_timeout(speed_mode, 8.0, 25.0)
    print(f"Speed mode: {speed_mode}")
    pending_rows = [row for group in groups for row in group["rows"]]
    excel_previous_mobile_by_sr: dict[str, str] = {}
    excel_last_mobile_by_awc: dict[str, str] = {}
    for source_row in rows:
        source_awc_key = normalize_match_text(clean(source_row.get("Institute Name")))
        source_sr = clean(source_row.get("Sr No"))
        if not source_awc_key or not source_sr:
            continue
        source_mobile = row_valid_mobile(source_row)
        if source_mobile:
            excel_last_mobile_by_awc[source_awc_key] = source_mobile
        elif excel_last_mobile_by_awc.get(source_awc_key):
            excel_previous_mobile_by_sr[source_sr] = excel_last_mobile_by_awc[source_awc_key]
    institute_mapping: dict[str, str] = {}
    if auto_select_institute:
        print("Auto institute mode: checking selected/pending AWC NAME values against the RBSK institute list...")
        institute_mapping = prepare_verified_institute_mapping(
            driver,
            pending_rows,
            args,
            session=getattr(args, "portal_session", None),
            context=getattr(args, "portal_context", None),
            label=label,
        )
        if getattr(args, "mapping_check_only", False):
            print("Mapping check complete. No live records were saved.")
            emit_progress(args, "batch_complete", label=label, saved=0, failed=0, skipped=skipped_count, pending=pending_total)
            return {
                "exit_code": 0,
                "saved": saved_count,
                "duplicates": duplicate_count,
                "failed": failed_count,
                "skipped": skipped_count,
                "skipped_this_run": skipped_this_run,
            }
        print(
            "Auto institute mode: selected block mapping verified. "
            "The tool will select each verified AWC when it changes."
        )
        if load_institute_aliases(args):
            print(f"Alias file loaded for runtime selection: {resolve_institute_alias_path(args)}")

    emit_progress(
        args,
        "batch_start",
        label=label,
        selected=len(rows),
        pending=pending_total,
        skipped=skipped_count,
        groups=[{"name": group["name"], "pending": len(group["rows"])} for group in groups],
    )
    pending_index = 0
    group_count = len(groups)
    for group_index, group in enumerate(groups, start=1):
        institute_name = clean(group["name"]) or "(no institute)"
        institute_key = normalize_match_text(institute_name)
        portal_institute_name = institute_mapping.get(institute_key, institute_name)
        unmapped_names = getattr(args, "unmapped_institute_names", set()) or set()
        manual_for_this_institute = institute_key in unmapped_names
        group_rows = group["rows"]
        group_total = len(group_rows)
        group_start_sr = clean(group.get("start_sr"))
        print("\n" + "-" * 72)
        print(f"AWC {group_index}/{group_count}: {institute_name}")
        print(f"Pending in this AWC: {group_total}")
        emit_progress(
            args,
            "awc_start",
            label=label,
            institute=institute_name,
            institute_index=group_index,
            institute_total=group_count,
            pending=group_total,
            start_sr=group_start_sr,
        )
        previous_saved = ""
        if group_index > 1:
            if auto_select_institute and not manual_for_this_institute:
                previous_saved = f"Institute changed at Sr No {group_start_sr}. Tool will select: {portal_institute_name}."
            else:
                previous_saved = f"Institute changed at Sr No {group_start_sr}. Select new institute: {institute_name}."

        if manual_for_this_institute and unattended:
            reason = "AWC institute name not mapped"
            print(f"Skipping AWC block without waiting: {institute_name} ({group_total} row(s)) - {reason}")
            for institute_index, row in enumerate(group_rows, start=1):
                pending_index += 1
                sr_no = clean(row.get("Sr No"))
                name = clean(row.get("Student Full Name"))
                append_skipped_row(
                    args,
                    row,
                    "skipped_awc_unmapped",
                    reason,
                    "Add this AWC to institute_aliases.csv with the exact RBSK dropdown name, then rerun this block.",
                    institute_name,
                )
                skipped_this_run += 1
                emit_progress(
                    args,
                    "row_skipped",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    status="skipped_awc_unmapped",
                    error=reason,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=group_total - institute_index,
                )
            continue

        group_skip_reason = ""
        for institute_index, row in enumerate(group_rows, start=1):
            pending_index += 1
            sr_no = clean(row.get("Sr No"))
            name = clean(row.get("Student Full Name"))
            remaining_after = group_total - institute_index
            row_started_at = time.perf_counter()
            timing: dict[str, float] = {}
            if group_skip_reason:
                append_skipped_row(
                    args,
                    row,
                    "skipped_awc_auto_failed",
                    group_skip_reason,
                    "Check institute alias/portal state, then rerun this AWC block.",
                    institute_name,
                )
                skipped_this_run += 1
                print(f"[{pending_index}/{pending_total}] Skipped Sr No {sr_no}: {group_skip_reason}")
                emit_progress(
                    args,
                    "row_skipped",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    status="skipped_awc_auto_failed",
                    error=group_skip_reason,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                    row_seconds=time.perf_counter() - row_started_at,
                )
                continue
            if duplicate_skip_remaining > 0:
                reason = "Skipped this run after duplicate window"
                duplicate_skip_remaining -= 1
                append_skipped_row(
                    args,
                    row,
                    "skipped_duplicate_window",
                    reason,
                    "Review later. This row was not submitted and was not marked completed.",
                    institute_name,
                )
                skipped_this_run += 1
                print(f"[{pending_index}/{pending_total}] Skipped Sr No {sr_no} after duplicate window: {name}")
                emit_progress(
                    args,
                    "row_skipped",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    status="skipped_duplicate_window",
                    error=reason,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                    row_seconds=time.perf_counter() - row_started_at,
                )
                continue
            try:
                browser_wait_for_portal_idle(driver)
                print(
                    f"[{pending_index}/{pending_total}] {institute_name} "
                    f"{institute_index}/{group_total} (remaining after this: {remaining_after}) - Sr No {sr_no}: {name}"
                )
                emit_progress(
                    args,
                    "row_start",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                )
                form_root = None
                open_started_at = time.perf_counter()
                if institute_index == 1:
                    if group_index > 1:
                        try:
                            browser_cancel_student_form_if_open(driver)
                        except Exception:
                            pass
                    emit_progress(
                        args,
                        "awc_change_required",
                        label=label,
                        institute=institute_name,
                        sr_no=sr_no,
                        name=name,
                        institute_index=group_index,
                        institute_total=group_count,
                        start_sr=group_start_sr,
                    )
                    if auto_select_institute and not manual_for_this_institute:
                        try:
                            print(f"Auto-selecting institute: {portal_institute_name}")
                            form_root = browser_open_add_student(driver, portal_institute_name, speed_mode=speed_mode)
                            timing["open"] = time.perf_counter() - open_started_at
                        except Exception as auto_exc:
                            print(f"Automatic institute/Add Child failed: {auto_exc}")
                            if unattended:
                                raise RuntimeError(f"Automatic institute/Add Child failed in unattended mode: {auto_exc}")
                            if not getattr(args, "manual_form_open", True):
                                raise
                            print("Manual fallback: select this institute in RBSK, click Add Child if visible, or Search only if needed, then confirm.")
                            previous_saved = (
                                f"Automatic institute selection failed for {portal_institute_name}: {auto_exc}\n"
                                "Please select the institute manually once, click Add Child if visible, or Search only if Add Child is not visible."
                            )
                            form_root = wait_for_manual_student_form(
                                driver,
                                args,
                                row,
                                pending_index,
                                pending_total,
                                institute_name,
                                same_institute=False,
                                institute_index=institute_index,
                                institute_total=group_total,
                                remaining_after_this=remaining_after,
                                previous_saved=previous_saved,
                                force_prompt=True,
                            )
                            timing["open"] = time.perf_counter() - open_started_at
                    else:
                        if manual_for_this_institute:
                            print(f"Manual institute selection required for unmapped AWC: {institute_name}")
                            previous_saved = (
                                f"AWC NAME is not mapped yet: {institute_name}\n"
                                "Please select the correct institute manually once, click Add Child if visible, "
                                "or Search only if Add Child is not visible."
                            )
                        form_root = wait_for_manual_student_form(
                            driver,
                            args,
                            row,
                            pending_index,
                            pending_total,
                            institute_name,
                            same_institute=False,
                            institute_index=institute_index,
                            institute_total=group_total,
                            remaining_after_this=remaining_after,
                            previous_saved=previous_saved,
                            force_prompt=True,
                        )
                        timing["open"] = time.perf_counter() - open_started_at
                else:
                    form_root = browser_find_student_form_root(driver)
                if not form_root:
                    if institute_index > 1:
                        print("Opening Add Student/Add Child automatically for the same AWC...")
                        form_root = browser_try_open_add_child_current_institute(driver, timeout=initial_same_awc_open_timeout, speed_mode=speed_mode)
                    if not form_root:
                        if institute_index > 1:
                            print("Automatic Add Student/Add Child did not open. Asking for one manual click.")
                        if unattended:
                            raise RuntimeError("Automatic Add Student/Add Child did not open in unattended mode.")
                        if not getattr(args, "manual_form_open", True):
                            raise RuntimeError("Automatic Add Student/Add Child did not open, and manual fallback is turned off.")
                        form_root = wait_for_manual_student_form(
                            driver,
                            args,
                            row,
                            pending_index,
                            pending_total,
                            institute_name,
                            same_institute=institute_index > 1,
                            institute_index=institute_index,
                            institute_total=group_total,
                            remaining_after_this=remaining_after,
                            previous_saved=previous_saved,
                        )
                    timing["open"] = time.perf_counter() - open_started_at
                elif "open" not in timing:
                    timing["open"] = time.perf_counter() - open_started_at
                effective_row = row
                fallback_mobile_message = ""
                current_mobile = row_valid_mobile(row)
                if current_mobile:
                    last_valid_mobile_by_awc[institute_key] = current_mobile
                else:
                    fallback_mobile = last_valid_mobile_by_awc.get(institute_key, "") or excel_previous_mobile_by_sr.get(sr_no, "")
                    if fallback_mobile:
                        effective_row = dict(row)
                        effective_row["Mobile Number"] = fallback_mobile
                        fallback_mobile_message = f" Used previous same-AWC mobile: {fallback_mobile}."
                        print(f"[{pending_index}/{pending_total}] Missing mobile for Sr No {sr_no}; using previous same-AWC mobile {fallback_mobile}.")
                    elif is_awc_mode:
                        effective_row = dict(row)
                        effective_row["Mobile Number"] = ""
                        fallback_mobile_message = " AWC/AWS mode left mobile blank."
                        print(f"[{pending_index}/{pending_total}] AWC/AWS mode: missing mobile for Sr No {sr_no}; leaving mobile blank.")
                    else:
                        reason = "Missing mobile and no previous same-AWC mobile"
                        append_skipped_row(
                            args,
                            row,
                            "skipped_missing_mobile",
                            reason,
                            "Add a valid mobile number, or make sure an earlier row in the same AWC has a valid mobile.",
                            institute_name,
                        )
                        skipped_this_run += 1
                        print(f"[{pending_index}/{pending_total}] Skipped Sr No {sr_no}: {reason}")
                        try:
                            browser_cancel_student_form_if_open(driver)
                        except Exception:
                            pass
                        emit_progress(
                            args,
                            "row_skipped",
                            label=label,
                            institute=institute_name,
                            sr_no=sr_no,
                            name=name,
                            status="skipped_missing_mobile",
                            error=reason,
                            pending_index=pending_index,
                            pending_total=pending_total,
                            institute_index=institute_index,
                            institute_total=group_total,
                            remaining_after=remaining_after,
                            row_seconds=time.perf_counter() - row_started_at,
                        )
                        continue
                fill_started_at = time.perf_counter()
                browser_fill_student_form(driver, effective_row, form_root, speed_mode=speed_mode)
                browser_verify_student_form_matches_row(driver, effective_row, form_root)
                timing["fill_check"] = time.perf_counter() - fill_started_at
                save_started_at = time.perf_counter()
                browser_save_student(driver, form_root, speed_mode=speed_mode)
                timing["save"] = time.perf_counter() - save_started_at
                append_csv(
                    Path(args.success_log),
                    {"Sr No": sr_no, "Student Full Name": name, "portal_student_code": "", "message": "Saved through browser form" + fallback_mobile_message},
                    success_headers,
                )
                completed.add(sr_no)
                save_progress(Path(args.progress), completed)
                saved_count += 1
                previous_saved = f"Saved {institute_index}/{group_total} for {institute_name}. Remaining: {remaining_after}."
                print(f"[{pending_index}/{pending_total}] Saved Sr No {sr_no}: {name}")
                print(f"Saved {institute_index}/{group_total} for {institute_name}. Remaining: {remaining_after}.")
                if remaining_after == 0:
                    print(f"AWC complete: {institute_name}")
                    emit_progress(args, "awc_complete", label=label, institute=institute_name, institute_index=group_index, institute_total=group_count)
                else:
                    print(f"Opening next Add Child automatically for {institute_name}...")
                    next_started_at = time.perf_counter()
                    next_root = browser_try_open_add_child_current_institute(driver, timeout=same_awc_open_timeout, speed_mode=speed_mode)
                    timing["next_open"] = time.perf_counter() - next_started_at
                    if next_root:
                        print("Next Add Child form is ready; continuing automatically.")
                    else:
                        print("Next Add Child form did not open yet; the next row will retry automatically.")
                row_seconds = time.perf_counter() - row_started_at
                emit_progress(
                    args,
                    "row_saved",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                    saved=saved_count,
                    failed=failed_count,
                    row_seconds=row_seconds,
                )
                print(
                    "Timing Sr No "
                    f"{sr_no}: total {row_seconds:.1f}s | open {timing.get('open', 0):.1f}s | "
                    f"fill/check {timing.get('fill_check', 0):.1f}s | save {timing.get('save', 0):.1f}s | "
                    f"next form {timing.get('next_open', 0):.1f}s"
                )
            except DuplicateRecordFound as exc:
                duplicate_message = clean(str(exc)) or "Duplicate record already exists in RBSK"
                try:
                    browser_close_blocking_message_dialog(driver)
                    browser_cancel_student_form_if_open(driver)
                    browser_wait_for_portal_idle(driver, timeout=8.0)
                except Exception:
                    pass
                append_csv(
                    Path(args.success_log),
                    {
                        "Sr No": sr_no,
                        "Student Full Name": name,
                        "portal_student_code": "",
                        "message": f"Already exists in RBSK; marked completed. {duplicate_message}",
                    },
                    success_headers,
                )
                completed.add(sr_no)
                save_progress(Path(args.progress), completed)
                duplicate_count += 1
                duplicate_skip_remaining = 20
                previous_saved = f"Sr No {sr_no} already exists in RBSK. Marked completed. Remaining: {remaining_after}."
                print(f"[{pending_index}/{pending_total}] Already exists in RBSK; marked completed Sr No {sr_no}: {name}")
                print(f"Duplicate detail: {duplicate_message}")
                print("Duplicate rule: skipping the next 20 unattempted rows in this run. They are not marked completed.")
                if remaining_after == 0:
                    print(f"AWC complete: {institute_name}")
                    emit_progress(args, "awc_complete", label=label, institute=institute_name, institute_index=group_index, institute_total=group_count)
                row_seconds = time.perf_counter() - row_started_at
                emit_progress(
                    args,
                    "row_saved",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                    saved=saved_count,
                    failed=failed_count,
                    row_seconds=row_seconds,
                    duplicate=True,
                    message=duplicate_message,
                )
            except Exception as exc:
                if browser_session_error(exc):
                    raise RuntimeError(
                        f"Browser was closed or stopped. Restart from Sr No {sr_no}. "
                        "Completed rows already saved in progress.json will be skipped."
                    ) from exc
                error_text = str(exc)
                if error_text.startswith("Pre-save safety check failed") or error_text.startswith("Portal did not accept DOB/Gender fill"):
                    error_text = "Save not attempted because safety check failed: " + error_text
                    print(f"[{pending_index}/{pending_total}] Save not attempted for Sr No {sr_no}: {exc}")
                else:
                    print(f"[{pending_index}/{pending_total}] Could not save Sr No {sr_no}: {exc}")
                if unattended and institute_index == 1 and "automatic institute/add child failed in unattended mode" in error_text.lower():
                    group_skip_reason = "Automatic institute/Add Child failed for this AWC in unattended mode"
                screenshot_path = None
                diagnostic_path = None
                try:
                    failure_dir = Path(args.failed_log).parent
                    failure_dir.mkdir(parents=True, exist_ok=True)
                    screenshot_path = failure_dir / f"failure_sr_{sr_no or pending_index}.png"
                    driver.save_screenshot(str(screenshot_path))
                    print(f"Saved failure screenshot: {screenshot_path}")
                    diagnostic_path = failure_dir / f"failure_sr_{sr_no or pending_index}_buttons.html"
                    save_button_diagnostic(driver, diagnostic_path)
                    print(f"Saved button diagnostic: {diagnostic_path}")
                except Exception:
                    pass
                append_csv(Path(args.failed_log), {"Sr No": sr_no, "Student Full Name": name, "error": error_text, "portal_response": ""}, failed_headers)
                failed_count += 1
                emit_progress(
                    args,
                    "row_failed",
                    label=label,
                    institute=institute_name,
                    sr_no=sr_no,
                    name=name,
                    error=error_text,
                    screenshot=str(screenshot_path) if screenshot_path else "",
                    diagnostic=str(diagnostic_path) if diagnostic_path else "",
                    pending_index=pending_index,
                    pending_total=pending_total,
                    institute_index=institute_index,
                    institute_total=group_total,
                    remaining_after=remaining_after,
                    saved=saved_count,
                    failed=failed_count,
                    row_seconds=time.perf_counter() - row_started_at,
                )
                previous_saved = f"Previous row failed: Sr No {sr_no}. It is logged in failed.csv. Continue with the next child in the same AWC."
                try:
                    browser_close_blocking_message_dialog(driver)
                    browser_cancel_student_form_if_open(driver)
                    browser_wait_for_portal_idle(driver, timeout=8.0)
                except Exception:
                    print("Portal still loading after failure; refreshing Student Master before next row.")
                    driver.get(STUDENTMASTER_URL)
                    browser_wait_for_portal_idle(driver, timeout=30.0)
            time.sleep(args.delay)

    print(f"\n{label} run complete.")
    completed_or_skipped = saved_count + duplicate_count + skipped_this_run
    emit_progress(
        args,
        "batch_complete",
        label=label,
        saved=saved_count,
        duplicates=duplicate_count,
        failed=failed_count,
        skipped=skipped_count,
        skipped_this_run=skipped_this_run,
        pending=max(pending_total - completed_or_skipped, 0),
    )
    return {
        "exit_code": 0 if failed_count == 0 else 1,
        "saved": saved_count,
        "duplicates": duplicate_count,
        "failed": failed_count,
        "skipped": skipped_count,
        "skipped_this_run": skipped_this_run,
    }


def process_student_rows_browser(driver, rows: list[dict[str, Any]], args: argparse.Namespace, label: str = "Browser Form Fill") -> dict[str, Any]:
    manual_form_open = bool(
        getattr(args, "manual_form_open", False)
        or getattr(args, "auto_select_institute", False)
        or (getattr(args, "browser_form_fill", False) and not hasattr(args, "manual_form_open"))
    )
    if manual_form_open:
        return process_student_rows_browser_manual_countdown(driver, rows, args, label)
    print(f"\n{label} selected rows: {len(rows)}")
    completed = load_progress(Path(args.progress)) if args.resume else set()
    success_headers = ["Sr No", "Student Full Name", "portal_student_code", "message"]
    failed_headers = ["Sr No", "Student Full Name", "error", "portal_response"]
    saved_count = failed_count = skipped_count = 0
    current_institute = None
    speed_mode = clean(getattr(args, "speed_mode", "")) or "Fast Safe"

    for index, row in enumerate(rows, start=1):
        sr_no = clean(row.get("Sr No"))
        name = clean(row.get("Student Full Name"))
        if args.resume and sr_no in completed:
            print(f"[{index}/{len(rows)}] Skipping already completed Sr No {sr_no}: {name}")
            skipped_count += 1
            continue
        try:
            browser_wait_for_portal_idle(driver)
            institute_name = clean(row.get("Institute Name"))
            print(f"[{index}/{len(rows)}] Filling Sr No {sr_no}: {name}")
            form_root = None
            if manual_form_open:
                institute_key = normalize_match_text(institute_name)
                same_institute = bool(institute_key and institute_key == current_institute)
                if same_institute:
                    form_root = browser_try_open_add_child_current_institute(driver, speed_mode=speed_mode)
                    if not form_root:
                        print("Same AWC still selected, but Add Child did not open automatically. Asking for one quick manual click.")
                        form_root = wait_for_manual_student_form(driver, args, row, index, len(rows), institute_name, same_institute=True)
                else:
                    form_root = wait_for_manual_student_form(driver, args, row, index, len(rows), institute_name)
                current_institute = institute_key
            else:
                if institute_name and normalize_match_text(institute_name) != current_institute:
                    print(f"Selecting institute: {institute_name}")
                    try:
                        browser_select_institute(driver, institute_name)
                        current_institute = normalize_match_text(institute_name)
                    except Exception as exc:
                        print(f"Automatic institute selection failed: {exc}")
                        print("Switching to manual institute/Add Student handoff for this row.")
                        form_root = wait_for_manual_student_form(driver, args, row, index, len(rows), institute_name)
                        current_institute = normalize_match_text(institute_name)
                if form_root is None:
                    try:
                        form_root = browser_open_add_student(driver, institute_name, speed_mode=speed_mode)
                    except Exception as exc:
                        print(f"Automatic Add Student opening failed: {exc}")
                        print("Switching to manual institute/Add Student handoff for this row.")
                        form_root = wait_for_manual_student_form(driver, args, row, index, len(rows), institute_name)
            browser_fill_student_form(driver, row, form_root, speed_mode=speed_mode)
            browser_verify_student_form_matches_row(driver, row, form_root)
            browser_save_student(driver, form_root, speed_mode=speed_mode)
            append_csv(Path(args.success_log), {"Sr No": sr_no, "Student Full Name": name, "portal_student_code": "", "message": "Saved through browser form"}, success_headers)
            completed.add(sr_no)
            save_progress(Path(args.progress), completed)
            saved_count += 1
            print(f"[{index}/{len(rows)}] Saved Sr No {sr_no}: {name}")
        except DuplicateRecordFound as exc:
            duplicate_message = clean(str(exc)) or "Duplicate record already exists in RBSK"
            try:
                browser_close_blocking_message_dialog(driver)
                browser_cancel_student_form_if_open(driver)
                browser_wait_for_portal_idle(driver, timeout=8.0)
            except Exception:
                pass
            append_csv(
                Path(args.success_log),
                {
                    "Sr No": sr_no,
                    "Student Full Name": name,
                    "portal_student_code": "",
                    "message": f"Already exists in RBSK; marked completed. {duplicate_message}",
                },
                success_headers,
            )
            completed.add(sr_no)
            save_progress(Path(args.progress), completed)
            saved_count += 1
            print(f"[{index}/{len(rows)}] Already exists in RBSK; marked completed Sr No {sr_no}: {name}")
            print(f"Duplicate detail: {duplicate_message}")
        except Exception as exc:
            if browser_session_error(exc):
                raise RuntimeError(
                    f"Browser was closed or stopped. Restart from Sr No {sr_no}. "
                    "Completed rows already saved in progress.json will be skipped."
                ) from exc
            print(f"[{index}/{len(rows)}] Failed Sr No {sr_no}: {exc}")
            try:
                failure_dir = Path(args.failed_log).parent
                failure_dir.mkdir(parents=True, exist_ok=True)
                screenshot_path = failure_dir / f"failure_sr_{sr_no or index}.png"
                driver.save_screenshot(str(screenshot_path))
                print(f"Saved failure screenshot: {screenshot_path}")
                diagnostic_path = failure_dir / f"failure_sr_{sr_no or index}_buttons.html"
                save_button_diagnostic(driver, diagnostic_path)
                print(f"Saved button diagnostic: {diagnostic_path}")
            except Exception:
                pass
            append_csv(Path(args.failed_log), {"Sr No": sr_no, "Student Full Name": name, "error": str(exc), "portal_response": ""}, failed_headers)
            failed_count += 1
            try:
                browser_close_blocking_message_dialog(driver)
                browser_cancel_student_form_if_open(driver)
                browser_wait_for_portal_idle(driver, timeout=8.0)
            except Exception:
                print("Portal still loading after failure; refreshing Student Master before next row.")
                driver.get(STUDENTMASTER_URL)
                browser_wait_for_portal_idle(driver, timeout=30.0)
                current_institute = None
        time.sleep(args.delay)
    print(f"\n{label} run complete.")
    return {"exit_code": 0 if failed_count == 0 else 1, "saved": saved_count, "failed": failed_count, "skipped": skipped_count}


def demo_row_values(row: dict[str, Any]) -> dict[str, str]:
    values: dict[str, str] = {}
    for field_name, column_name in DEMO_FIELDS.items():
        if field_name == "birth_date":
            values[field_name] = excel_date_to_iso(row.get(column_name))
        elif field_name in {"mobile_number", "pin_code"}:
            values[field_name] = only_digits(row.get(column_name), 10 if field_name == "mobile_number" else 6)
        else:
            values[field_name] = clean(row.get(column_name))
    return values


def fill_demo_form(driver, values: dict[str, str]) -> None:
    script = """
    const values = arguments[0];
    const addButton = document.querySelector('#addStudentButton');
    if (addButton) addButton.click();
    const setValue = (el, value) => {
      if (!el) return false;
      el.focus();
      const setter = Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set;
      if (setter) setter.call(el, value);
      else el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    };
    Object.entries(values).forEach(([name, value]) => {
      setValue(document.querySelector(`[name="${name}"]`), value);
    });
    document.querySelector('#studentForm').requestSubmit();
    return document.querySelectorAll('#entryTable tbody tr').length;
    """
    inserted_count = driver.execute_script(script, values)
    print(f"Demo form rows visible: {inserted_count}")


def run_demo_form(args: argparse.Namespace) -> int:
    form_path = Path(DEMO_FORM_FILE).resolve()
    if not form_path.exists():
        raise ImportErrorWithHelp(f"Demo form not found: {form_path}")

    rows = normalize_student_rows(load_rows(Path(args.excel), args.sheet))
    validate_columns(rows)
    rows = select_rows(rows, args.only_sr, args.limit)
    if not rows:
        raise ImportErrorWithHelp("No rows selected for demo.")

    driver = start_browser(args.browser)
    demo_log_dir = Path(args.demo_log_dir)
    success_path = demo_log_dir / "demo_success.csv"
    headers = ["Sr No", "Student Full Name", "Birth Date", "Gender", "Class", "Section", "status"]

    try:
        driver.get(form_path.as_uri())
        time.sleep(1)
        for index, row in enumerate(rows, start=1):
            values = demo_row_values(row)
            print(f"[{index}/{len(rows)}] Demo inserting Sr No {clean(row.get('Sr No'))}: {values['student_full_name']}")
            fill_demo_form(driver, values)
            append_csv(
                success_path,
                {
                    "Sr No": clean(row.get("Sr No")),
                    "Student Full Name": values["student_full_name"],
                    "Birth Date": values["birth_date"],
                    "Gender": values["gender"],
                    "Class": values["class_name"],
                    "Section": values["section"],
                    "status": "shown_on_demo_form",
                },
                headers,
            )
            time.sleep(args.demo_delay)

        print("\nDemo form run complete.")
        print(f"Demo log: {success_path}")
        if args.keep_browser_open:
            input("Press Enter to close the demo browser...")
        return 0
    finally:
        if not args.keep_browser_open:
            driver.quit()


def process_student_rows(
    rows: list[dict[str, Any]],
    session: requests.Session,
    context: dict[str, Any],
    args: argparse.Namespace,
    label: str = "Import",
) -> dict[str, Any]:
    print("\nResolved context:")
    print(json.dumps(context, indent=2))
    print(f"{label} selected rows: {len(rows)}")

    class_map = fetch_option_map(session, 23, context["state_code"]) if not args.no_browser else {}
    section_map = fetch_option_map(session, 25, context["state_code"]) if not args.no_browser else {}
    uses_row_institute = any(clean(row.get("Institute Name")) or clean(row.get("Institution Code") or row.get("Institute Code")) for row in rows)
    institution_map = fetch_institution_map(session, context) if uses_row_institute and not args.no_browser else {}
    if uses_row_institute and args.dry_run and not institution_map:
        placeholder_code = clean(context.get("institution_code")) or "0"
        institution_map = {normalize_match_text(row.get("Institute Name")): placeholder_code for row in rows if clean(row.get("Institute Name"))}
    if uses_row_institute:
        print(f"Per-row institute mode: enabled ({len(institution_map)} institute names loaded from RBSK)")

    completed = load_progress(Path(args.progress)) if args.resume else set()
    success_headers = ["Sr No", "Student Full Name", "portal_student_code", "message"]
    failed_headers = ["Sr No", "Student Full Name", "error", "portal_response"]
    saved_count = 0
    failed_count = 0
    skipped_count = 0

    for index, row in enumerate(rows, start=1):
        sr_no = clean(row.get("Sr No"))
        name = clean(row.get("Student Full Name"))
        if args.resume and sr_no in completed:
            print(f"[{index}/{len(rows)}] Skipping already completed Sr No {sr_no}: {name}")
            skipped_count += 1
            continue

        try:
            payload = build_payload(row, context, class_map, section_map, institution_map)
            errors = validate_payload(payload)
            if errors:
                raise ValueError("; ".join(errors))

            if args.dry_run:
                print(f"\nDry-run payload for Sr No {sr_no}:")
                print(json.dumps(payload, indent=2, ensure_ascii=False))
                continue

            ok, message, response_data = submit_student(session, payload, args.timeout)
            if ok:
                portal_code = response_data.get("status_Code") if isinstance(response_data, dict) else ""
                print(f"[{index}/{len(rows)}] Saved Sr No {sr_no}: {name} -> {portal_code}")
                append_csv(Path(args.success_log), {"Sr No": sr_no, "Student Full Name": name, "portal_student_code": portal_code, "message": message}, success_headers)
                completed.add(sr_no)
                save_progress(Path(args.progress), completed)
                saved_count += 1
            else:
                print(f"[{index}/{len(rows)}] Failed Sr No {sr_no}: {message}")
                append_csv(
                    Path(args.failed_log),
                    {"Sr No": sr_no, "Student Full Name": name, "error": message, "portal_response": json.dumps(response_data, ensure_ascii=False)},
                    failed_headers,
                )
                failed_count += 1
                if "Session expired" in message:
                    return {"exit_code": 2, "saved": saved_count, "failed": failed_count, "skipped": skipped_count}
            time.sleep(args.delay)
        except Exception as exc:
            print(f"[{index}/{len(rows)}] Failed Sr No {sr_no}: {exc}")
            append_csv(Path(args.failed_log), {"Sr No": sr_no, "Student Full Name": name, "error": str(exc), "portal_response": ""}, failed_headers)
            failed_count += 1

    print(f"\n{label} run complete.")
    return {"exit_code": 0, "saved": saved_count, "failed": failed_count, "skipped": skipped_count}


def run_import(args: argparse.Namespace) -> int:
    args.screening_health_entry = True
    args.browser_form_fill = True
    rows = normalize_student_rows(load_rows(Path(args.excel), args.sheet))
    validate_screening_health_rows(rows)
    rows = select_rows(rows, args.only_sr, args.limit)
    if not rows:
        raise ImportErrorWithHelp("No rows selected for import.")

    session: requests.Session | None = None
    storage: dict[str, Any] = {}
    driver = None

    if getattr(args, "browser_form_fill", False) and not args.dry_run:
        driver = capture_logged_in_driver(args)
        print("\nLogin step:")
        print("1. Type Login ID, password, and CAPTCHA manually in the browser.")
        input("Press Enter here only after you can see the logged-in RBSK portal...")
        print("\nScreening context step:")
        print("The tool will open Screening, select School, use your selected month, then process already-added students.")
        session, storage = collect_browser_session(driver)
    elif not args.dry_run or not args.no_browser:
        session, storage = capture_logged_in_session(args)

    parsed_storage = parse_storage(storage)
    context = infer_context(parsed_storage, args)
    uses_row_institute = any(clean(row.get("Institute Name")) or clean(row.get("Institution Code") or row.get("Institute Code")) for row in rows)
    live_browser_form = bool(getattr(args, "browser_form_fill", False) and not args.dry_run)
    ensure_context(
        context,
        require_created_by=not args.dry_run and not live_browser_form,
        require_institution_code=not uses_row_institute and not live_browser_form,
    )

    if session is None:
        session = new_requests_session_or_raise()
    session = apply_ssl_setting(session, args)

    if getattr(args, "browser_form_fill", False) and not args.dry_run:
        args.portal_session = session
        args.portal_context = context
        try:
            result = process_screening_health_rows_browser(driver, rows, args)
        finally:
            if driver:
                driver.quit()
    else:
        raise ImportErrorWithHelp("This script is Screening-only. It will not add/register students or submit Student Master records.")
    return int(result["exit_code"])


def slugify(value: str) -> str:
    text = clean(value) or "batch"
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    return text.strip("._-") or "batch"


def resolve_batch_path(value: str, control_path: Path) -> Path:
    path = Path(clean(value))
    if path.is_absolute():
        return path
    return (control_path.parent / path).resolve()


def batch_arg_namespace(base: argparse.Namespace, row: dict[str, Any], control_path: Path, batch_log_dir: Path) -> argparse.Namespace:
    batch_args = argparse.Namespace(**vars(base))
    batch_args.excel = str(resolve_batch_path(row.get("Excel File", ""), control_path))
    batch_args.sheet = clean(row.get("Sheet")) or base.sheet
    batch_args.state_code = clean(row.get("State Code")) or base.state_code
    batch_args.district_code = clean(row.get("District Code")) or base.district_code
    batch_args.health_block_code = clean(row.get("Health Block Code")) or base.health_block_code
    batch_args.institution_type = clean(row.get("Institution Type")) or base.institution_type
    batch_args.institution_code = clean(row.get("Institution Code")) or base.institution_code
    batch_args.created_by = base.created_by
    batch_args.success_log = str(batch_log_dir / "success.csv")
    batch_args.failed_log = str(batch_log_dir / "failed.csv")
    batch_args.skipped_log = str(batch_log_dir / "skipped.csv")
    batch_args.progress = str(batch_log_dir / "progress.json")
    return batch_args


def validate_batch_control(rows: list[dict[str, Any]]) -> None:
    if not rows:
        raise ImportErrorWithHelp("No rows found in the batch control file.")
    missing = [col for col in BATCH_REQUIRED_COLUMNS if col not in rows[0]]
    if missing:
        raise ImportErrorWithHelp(
            f"Missing required batch control column(s): {', '.join(missing)}\n"
            f"Required columns: {', '.join(BATCH_REQUIRED_COLUMNS)}"
        )


def append_batch_summary(path: Path, row: dict[str, Any]) -> None:
    headers = ["Batch Name", "Login ID", "Excel File", "status", "saved", "duplicates", "failed", "skipped", "skipped_this_run", "message"]
    append_csv(path, row, headers)


def run_batch_import(args: argparse.Namespace) -> int:
    args.screening_health_entry = True
    args.browser_form_fill = True
    control_path = Path(args.batch_control)
    batch_rows = load_rows(control_path, args.batch_sheet)
    validate_batch_control(batch_rows)
    selected_batches = select_rows_by_batch(batch_rows, args.only_batch)
    if args.batch_limit:
        selected_batches = selected_batches[: args.batch_limit]
    if not selected_batches:
        names = ", ".join(available_batch_names(batch_rows)) or "none"
        raise ImportErrorWithHelp(f"No batch rows selected. Leave Only Batch Name blank or use one of: {names}")

    logs_root = Path(args.logs_dir)
    summary_path = logs_root / "batch_summary.csv"
    overall_exit = 0

    print(f"\nBatch control: {control_path}")
    print(f"Selected batches: {len(selected_batches)}")

    for batch_index, batch_row in enumerate(selected_batches, start=1):
        batch_name = clean(batch_row.get("Batch Name")) or f"Batch {batch_index}"
        login_id = clean(batch_row.get("Login ID"))
        batch_log_dir = logs_root / slugify(batch_name)
        batch_args = batch_arg_namespace(args, batch_row, control_path, batch_log_dir)

        print("\n" + "=" * 72)
        print(f"Batch {batch_index}/{len(selected_batches)}: {batch_name}")
        print(f"Login ID: {login_id}")
        print(f"Excel: {batch_args.excel}")

        try:
            rows = normalize_student_rows(load_rows(Path(batch_args.excel), batch_args.sheet))
            batch_args.screening_health_entry = True
            batch_args.browser_form_fill = True
            validate_screening_health_rows(rows)
            rows = select_rows(rows, batch_args.only_sr, batch_args.limit)
            if not rows:
                raise ImportErrorWithHelp("No student rows selected for this batch.")

            session: requests.Session | None = None
            storage: dict[str, Any] = {}
            driver = None
            if getattr(batch_args, "browser_form_fill", False) and not batch_args.dry_run:
                password = None
                if not getattr(batch_args, "manual_login", False):
                    password = getpass.getpass(f"Enter password for {login_id}: ")
                driver = capture_logged_in_driver(batch_args, login_id=login_id, password=password)
                print("\nLogin step:")
                print("1. Type Login ID, password, and CAPTCHA manually in the browser.")
                input("Press Enter here only after you can see the logged-in RBSK portal...")
                print("\nScreening context step:")
                print("The tool will open Screening, select School, use your selected month, then process already-added students.")
                session, storage = collect_browser_session(driver)
            elif not batch_args.dry_run or not batch_args.no_browser:
                password = None
                if not getattr(batch_args, "manual_login", False):
                    password = getpass.getpass(f"Enter password for {login_id}: ")
                session, storage = capture_logged_in_session(batch_args, login_id=login_id, password=password)
            if session is None:
                session = new_requests_session_or_raise()
            session = apply_ssl_setting(session, batch_args)

            context = infer_context(parse_storage(storage), batch_args)
            uses_row_institute = any(clean(row.get("Institute Name")) or clean(row.get("Institution Code") or row.get("Institute Code")) for row in rows)
            live_browser_form = bool(getattr(batch_args, "browser_form_fill", False) and not batch_args.dry_run)
            ensure_context(
                context,
                require_created_by=not batch_args.dry_run and not live_browser_form,
                require_institution_code=not uses_row_institute and not live_browser_form,
            )
            if getattr(batch_args, "browser_form_fill", False) and not batch_args.dry_run:
                batch_args.portal_session = session
                batch_args.portal_context = context
                try:
                    result = process_screening_health_rows_browser(driver, rows, batch_args, label=batch_name)
                finally:
                    if driver:
                        driver.quit()
            else:
                raise ImportErrorWithHelp("This script is Screening-only. It will not add/register students or submit Student Master records.")
            status = "success" if result["exit_code"] == 0 and result["failed"] == 0 else "partial_or_failed"
            append_batch_summary(
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
                if "driver" in locals() and driver:
                    driver.quit()
            except Exception:
                pass
            overall_exit = 1
            print(f"Batch failed: {exc}")
            append_batch_summary(
                summary_path,
                {
                    "Batch Name": batch_name,
                    "Login ID": login_id,
                    "Excel File": clean(batch_row.get("Excel File")),
                    "status": "failed",
                    "saved": 0,
                    "duplicates": 0,
                    "failed": 0,
                    "skipped": 0,
                    "skipped_this_run": 0,
                    "message": str(exc),
                },
            )
            if args.stop_on_batch_error:
                return overall_exit

    print("\nBatch import run complete.")
    print(f"Summary: {summary_path}")
    return overall_exit


def select_rows_by_batch(rows: list[dict[str, Any]], only_batch: str | None) -> list[dict[str, Any]]:
    if not only_batch:
        return rows
    wanted = [slugify(part.strip()).lower() for part in only_batch.split(",") if part.strip()]
    selected = []
    for row in rows:
        batch_name = slugify(clean(row.get("Batch Name"))).lower()
        if any(want == batch_name or want in batch_name for want in wanted):
            selected.append(row)
    return selected


def available_batch_names(rows: list[dict[str, Any]]) -> list[str]:
    return [clean(row.get("Batch Name")) for row in rows if clean(row.get("Batch Name"))]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Bulk import RBSK student master rows after manual CAPTCHA login.")
    parser.add_argument("--excel", default="students_sample_data.xlsx", help="Path to .xlsx or .csv file.")
    parser.add_argument("--sheet", default="Students", help="Excel sheet name. Ignored for CSV.")
    parser.add_argument("--batch-control", help="Path to control .xlsx/.csv with Batch Name, Login ID, and Excel File columns.")
    parser.add_argument("--batch-sheet", default="Batches", help="Control Excel sheet name. Ignored for CSV.")
    parser.add_argument("--batch-limit", type=int, help="Only process the first N selected batch rows.")
    parser.add_argument("--only-batch", help="Comma-separated Batch Name values to process.")
    parser.add_argument("--logs-dir", default="logs", help="Folder for per-batch logs and batch_summary.csv.")
    parser.add_argument("--stop-on-batch-error", action="store_true", help="Stop the batch run when one center fails.")
    parser.add_argument("--demo-form", action="store_true", help="Run the safe local client demo form instead of RBSK.")
    parser.add_argument("--demo-log-dir", default="demo_logs", help="Folder for local demo logs.")
    parser.add_argument("--demo-delay", type=float, default=0.7, help="Delay between demo form rows.")
    parser.add_argument("--browser", choices=["chrome", "edge"], default="chrome")
    parser.add_argument("--dry-run", action="store_true", help="Build and print payloads without submitting.")
    parser.add_argument("--no-browser", action="store_true", help="Do not open browser. Requires all context override arguments.")
    parser.add_argument("--limit", type=int, help="Only process the first N selected rows.")
    parser.add_argument("--only-sr", help="Sr No selection, for example 1 or 1-5 or 1,3,5.")
    parser.add_argument("--resume", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--delay", type=float, default=0.25, help="Delay between portal submissions.")
    parser.add_argument("--timeout", type=int, default=120, help="HTTP timeout in seconds.")
    parser.add_argument("--keep-browser-open", action="store_true")
    parser.add_argument("--manual-login", action="store_true", help="Open RBSK and let the user type ID, password, and CAPTCHA manually.")
    parser.add_argument("--browser-form-fill", action="store_true", help="Use the live RBSK browser form instead of direct API submit.")
    parser.add_argument("--manual-form-open", action="store_true", help="For each row, let the user select institute and open Add Student/Add Child before filling.")
    parser.add_argument("--screening-health-entry", action="store_true", help="Use Screening page: select month/school/student and submit health details.")
    parser.add_argument("--screening-month", help="Month label to select on the Screening page. Leave blank to choose manually.")
    parser.add_argument("--auto-select-institute", action=argparse.BooleanOptionalAction, default=True, help="Automatically select the verified AWC institute when AWC NAME changes.")
    parser.add_argument("--unattended-after-start", action=argparse.BooleanOptionalAction, default=True, help="After live entry starts, never wait for operator popups; skip/log recoverable row or AWC problems and continue.")
    parser.add_argument("--mapping-check-only", action="store_true", help="Check Excel AWC NAME values against RBSK institute names, then stop without saving.")
    parser.add_argument("--institute-aliases", default="institute_aliases.csv", help="Optional CSV with Excel AWC NAME and RBSK Institute Name columns.")
    parser.add_argument("--speed-mode", choices=["Fast Safe", "Conservative"], default="Fast Safe", help="Fast Safe trims script waits while keeping pre-save checks. Conservative keeps longer waits.")
    parser.add_argument("--allow-insecure-ssl", action="store_true", help="Work around RBSK certificate chain errors in Python requests.")
    parser.add_argument("--success-log", default="success.csv")
    parser.add_argument("--failed-log", default="failed.csv")
    parser.add_argument("--skipped-log", default="skipped.csv")
    parser.add_argument("--progress", default="progress.json")
    parser.add_argument("--state-code")
    parser.add_argument("--district-code")
    parser.add_argument("--health-block-code")
    parser.add_argument("--institution-type", default="1", help="1 for AWC, 2 for School. Default: 1.")
    parser.add_argument("--institution-code")
    parser.add_argument("--created-by", help="Usually your login ID, for example MHT-1271251.")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.demo_form:
            return run_demo_form(args)
        if args.batch_control:
            return run_batch_import(args)
        return run_import(args)
    except ImportErrorWithHelp as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nStopped by user.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())

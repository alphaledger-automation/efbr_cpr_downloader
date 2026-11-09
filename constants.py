"""
eFBR CPR Downloader — Constants & Configuration
================================================
All URLs, CSS/XPath selectors, timing values, and column-detection
aliases live here so the rest of the code stays clean.
"""

import os
import pathlib
import platform

# ── Target URLs ───────────────────────────────────────────────────────────────
EFBR_HOME_URL             = "https://e.fbr.gov.pk"
EFBR_DOC_VERIFICATION_URL = "https://e.fbr.gov.pk/DocumentVerification.aspx"

# ── DOM Selectors (exact IDs first, then fallbacks) ──────────────────────────
# Search Parameter dropdown  (actual ID: ctl00_ContentPlaceHolder1_ddlSearchParam)
SEL_SEARCH_PARAM_DROPDOWN = (
    "#ctl00_ContentPlaceHolder1_ddlSearchParam, "
    "select[id*='ddlSearchParam'], "
    "select[id*='SearchParam'], "
    "select"
)

# CPR number text input  (actual ID: ctl00_ContentPlaceHolder1_txtEDNBarCode)
SEL_CPR_INPUT = (
    "#ctl00_ContentPlaceHolder1_txtEDNBarCode, "
    "input[id*='txtEDN'], "
    "input[id*='txtBarCode'], "
    "input[type='text']:visible"
)

# Search button  (actual ID: ctl00_ContentPlaceHolder1_btnSearch)
SEL_SEARCH_BUTTON = (
    "#ctl00_ContentPlaceHolder1_btnSearch, "
    "input[id*='btnSearch'], "
    "input[type='submit'][value='Search']"
)

# Download link on the popup page  (actual ID: lnkBtnPrint)
SEL_DOWNLOAD_LINK = (
    "#lnkBtnPrint, "
    "a[id*='lnkBtnPrint'], "
    "a:has-text('Click Here'), "
    "a:has-text('click here')"
)

# Verification menu link on the home page
SEL_VERIFICATION_MENU = "a:has-text('Verification')"
SEL_DOC_VERIFY_LINK   = "a:has-text('Online Document Verification')"

# CPR dropdown option value (actual value: "CPR")
CPR_OPTION_VALUE = "CPR"

# ── Timing (milliseconds for Playwright) ─────────────────────────────────────
PAGE_LOAD_TIMEOUT_MS   = 30_000
ELEMENT_TIMEOUT_MS     = 15_000
DOWNLOAD_TIMEOUT_MS    = 60_000
POPUP_TIMEOUT_MS       = 30_000

# ── Timing (seconds for Python sleeps) ───────────────────────────────────────
RETRY_DELAY_SECS       = 3
INTER_CPR_DELAY_SECS   = 1        # brief pause between CPRs
MAX_RETRIES            = 2

# ── Result statuses (shown in the log and written to the Excel summary) ─────
STATUS_DOWNLOADED     = "Downloaded"        # PDF saved
STATUS_NOT_EXIST      = "File not exist"    # portal has no receipt for this CPR
STATUS_LOGIN_ISSUE    = "Login issue"       # portal / verification page unreachable
STATUS_NOT_DOWNLOADED = "Not-Downloaded"    # failed, cancelled or never attempted
ALL_STATUSES = [
    STATUS_DOWNLOADED, STATUS_NOT_EXIST, STATUS_LOGIN_ISSUE,
    STATUS_NOT_DOWNLOADED,
]

# ── CPR Column Auto-Detection────────────────────────────────────────────────
CPR_COLUMN_ALIASES = [
    "cpr", "cprs",                                          # ← added "cprs"
    "cpr no", "cpr no.", "cpr_no", "cpr number", "cpr_number",
    "cprno", "cprnumber", "cpr #", "computerized payment receipt",
    "cpr numbers", "cpr nos", "cpr nos.", "cpr_nos",        # ← other plurals
]

# ── Output Paths ─────────────────────────────────────────────────────────────
DOWNLOAD_SUBFOLDER = "eFBR_CPRs"


def get_downloads_dir() -> pathlib.Path:
    """Return the user's Downloads folder (cross-platform)."""
    return pathlib.Path.home() / "Downloads"


def get_output_dir() -> pathlib.Path:
    """Return (and create) the CPR-specific download directory."""
    out = get_downloads_dir() / DOWNLOAD_SUBFOLDER
    out.mkdir(parents=True, exist_ok=True)
    return out


def load_output_prefs() -> dict:
    """Saved output-folder preferences: {"output_dir": str, "keep_name": bool}."""
    import json
    try:
        return json.loads(get_output_prefs_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_output_prefs(output_dir: str, keep_name: bool) -> None:
    import json
    try:
        get_output_prefs_path().write_text(
            json.dumps({"output_dir": output_dir, "keep_name": keep_name}),
            encoding="utf-8")
    except Exception:
        pass


def get_output_prefs_path() -> pathlib.Path:
    return get_settings_path().with_name("output_prefs.json")


def get_settings_path() -> pathlib.Path:
    """Return the per-user file that stores UI preferences (theme, design)."""
    if platform.system() == "Windows":
        base = pathlib.Path(os.environ.get("LOCALAPPDATA", pathlib.Path.home()))
    else:
        base = pathlib.Path.home() / ".local" / "share"
    d = base / "eFBR_CPR_Downloader"
    d.mkdir(parents=True, exist_ok=True)
    return d / "ui_settings.json"


# ── GUI Appearance ────────────────────────────────────────────────────────────
APP_NAME    = "eFBR CPR Receipt Downloader"
APP_VERSION = "1.1.0"
ACCENT      = "#00755E"          # tropical jungle green
SUCCESS_CLR = "#2FA35A"
ERROR_CLR   = "#E05252"
WARN_CLR    = "#D4A017"

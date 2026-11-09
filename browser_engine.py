"""
eFBR CPR Downloader — Playwright Browser Engine
================================================
Handles all interaction with the eFBR Document Verification portal:
  1. Launch Chromium → navigate via Verification menu hover → click
     "Online Document Verification".
  2. For each CPR: select "Computerized Payment Receipt" from dropdown,
     fill CPR input, click Search, catch popup tab, click "Click Here
     to download", save PDF, close popup.
  3. Retry failed CPRs up to MAX_RETRIES times.
  4. Communicate progress/logs back to the GUI via queues.

Designed to run entirely inside a background daemon thread.
"""

import os
import re
import time
import queue
import threading
import datetime
import pathlib
from dataclasses import dataclass
from typing import List, Optional, Tuple

from playwright.sync_api import sync_playwright, Page, BrowserContext, Browser

from constants import (
    EFBR_HOME_URL,
    EFBR_DOC_VERIFICATION_URL,
    SEL_SEARCH_PARAM_DROPDOWN,
    SEL_CPR_INPUT,
    SEL_SEARCH_BUTTON,
    SEL_DOWNLOAD_LINK,
    SEL_VERIFICATION_MENU,
    SEL_DOC_VERIFY_LINK,
    CPR_OPTION_VALUE,
    PAGE_LOAD_TIMEOUT_MS,
    ELEMENT_TIMEOUT_MS,
    DOWNLOAD_TIMEOUT_MS,
    POPUP_TIMEOUT_MS,
    RETRY_DELAY_SECS,
    INTER_CPR_DELAY_SECS,
    MAX_RETRIES,
    STATUS_DOWNLOADED,
    STATUS_NOT_EXIST,
    STATUS_LOGIN_ISSUE,
    STATUS_NOT_DOWNLOADED,
)


# ── Inter-thread message types ───────────────────────────────────────────────

@dataclass
class LogEntry:
    """Single log line sent from the worker to the GUI."""
    timestamp: str
    message: str
    level: str = "info"          # info | ok | err | warn


@dataclass
class ProgressUpdate:
    """Progress snapshot sent from the worker to the GUI."""
    total: int = 0
    processed: int = 0
    current_cpr: str = ""
    is_done: bool = False
    report_path: str = ""


# ── Helper utilities ─────────────────────────────────────────────────────────

def _now() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def _now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _safe_filename(cpr: str) -> str:
    """Sanitise a CPR string so it's safe for use as a filename."""
    return re.sub(r'[<>:"/\\|?*]', "_", cpr.strip())


def _clean_original_name(name: str) -> str:
    """Make the portal-supplied filename safe, keeping its text intact."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (name or "").strip())
    return name.strip(". ")


def _unique_path(folder: pathlib.Path, filename: str, cpr: str) -> pathlib.Path:
    """
    Return a path in *folder* that does not overwrite an existing file.
    On a clash: name_<CPR>.ext, then name_<CPR>_2.ext, ...
    """
    target = folder / filename
    if not target.exists():
        return target
    stem, ext = os.path.splitext(filename)
    cand = folder / f"{stem}_{_safe_filename(cpr)}{ext}"
    n = 2
    while cand.exists():
        cand = folder / f"{stem}_{_safe_filename(cpr)}_{n}{ext}"
        n += 1
    return cand


# ── Download Worker ──────────────────────────────────────────────────────────

class DownloadWorker:
    """
    Playwright-based automation that downloads CPR receipts from eFBR.

    Usage
    -----
    worker = DownloadWorker(cpr_list, output_dir, log_q, prog_q,
                            stop_event, pause_event)
    thread = threading.Thread(target=worker.run, daemon=True)
    thread.start()
    """

    def __init__(
        self,
        cpr_list: List[str],
        output_dir: pathlib.Path,
        log_q: queue.Queue,
        prog_q: queue.Queue,
        stop_event: threading.Event,
        pause_event: threading.Event,
        keep_original_name: bool = True,
        duplicates: Optional[List[dict]] = None,
    ):
        self.cpr_list    = cpr_list
        self.output_dir  = output_dir
        self.log_q       = log_q
        self.prog_q      = prog_q
        self.stop_event  = stop_event
        self.pause_event = pause_event      # SET = paused
        self.keep_original_name = keep_original_name
        self.duplicates = duplicates or []  # skipped repeats (for the report)

        self.results: List[dict] = []       # populated as we go
        self._portal_failed = False         # never reached the search page

        # Playwright objects — created in run()
        self._pw: Optional[object] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None

    # ── Logging / progress helpers ────────────────────────────────────────

    def _log(self, msg: str, level: str = "info"):
        self.log_q.put(LogEntry(timestamp=_now(), message=msg, level=level))

    def _progress(self, processed: int, current_cpr: str = ""):
        self.prog_q.put(ProgressUpdate(
            total=len(self.cpr_list),
            processed=processed,
            current_cpr=current_cpr,
        ))

    def full_results(self) -> List[dict]:
        """
        One row per CPR in the input list, in order.  CPRs that were never
        attempted (cancelled run, fatal error) are included too.
        """
        rows = list(self.results)
        if self._portal_failed:
            status, remark = STATUS_LOGIN_ISSUE, "Could not open the eFBR portal"
        else:
            status, remark = STATUS_NOT_DOWNLOADED, "Not attempted"
        for cpr in self.cpr_list[len(rows):]:
            rows.append({
                "CPR No": cpr,
                "Status": status,
                "Downloaded File Path": remark,
                "Timestamp": "",
            })
        return rows

    # ── Public entry point ────────────────────────────────────────────────

    def run(self):
        """Call this in a daemon thread."""
        try:
            self._run_inner()
        except Exception as exc:
            self._log(f"Fatal error: {exc}", "err")
        finally:
            self._cleanup()
            # Signal completion
            self.prog_q.put(ProgressUpdate(
                total=len(self.cpr_list),
                processed=len(self.results),
                is_done=True,
            ))

    # ── Core loop ─────────────────────────────────────────────────────────

    def _run_inner(self):
        # 1. Launch browser
        self._log("Launching Chromium browser (background)...")
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=True,          # ← runs invisibly, never steals focus
        )
        self._context = self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
        )
        self._page = self._context.new_page()
        self._page.set_default_timeout(ELEMENT_TIMEOUT_MS)

        # 2. Navigate: eFBR home → hover Verification → click Online
        #    Document Verification
        try:
            self._navigate_to_doc_verification()
        except Exception:
            self._portal_failed = True
            raise

        # 3. Process each CPR
        total = len(self.cpr_list)
        processed = 0

        for idx, cpr in enumerate(self.cpr_list):
            # ── Stop check ────────────────────────────────────────────────
            if self.stop_event.is_set():
                self._log("Download cancelled by user.", "warn")
                break

            # ── Pause check ───────────────────────────────────────────────
            while self.pause_event.is_set() and not self.stop_event.is_set():
                time.sleep(0.5)
            if self.stop_event.is_set():
                self._log("Download cancelled by user.", "warn")
                break

            # ── Process this CPR ──────────────────────────────────────────
            self._progress(processed, cpr)
            self._log(f"[{idx + 1}/{total}] Processing {cpr}...")

            status, file_path = self._download_with_retry(cpr)
            processed += 1

            self.results.append({
                "CPR No": cpr,
                "Status": status,
                "Downloaded File Path": file_path,
                "Timestamp": _now_iso(),
            })

            level = {STATUS_DOWNLOADED: "ok", STATUS_NOT_EXIST: "warn"}.get(
                status, "err")
            self._log(f"[{idx + 1}/{total}] {cpr} ... {status}", level)

            self._progress(processed, cpr)

            # Brief pause between CPRs to avoid hammering the server
            if idx < total - 1:
                time.sleep(INTER_CPR_DELAY_SECS)

        self._log(
            f"Finished: {processed}/{total} CPRs processed.", "ok"
        )

    # ── Navigation ────────────────────────────────────────────────────────

    def _navigate_to_doc_verification(self):
        """
        Navigate to the Document Verification page:
        1. Open eFBR home page (lands on authlogin.aspx).
        2. Hover over "Verification" menu (do NOT click — clicking
           triggers a postback that refreshes the page).
        3. Click "Online Document Verification" from the dropdown.
        """
        page = self._page

        self._log("Navigating to eFBR portal...")
        page.goto(EFBR_HOME_URL, wait_until="domcontentloaded",
                  timeout=PAGE_LOAD_TIMEOUT_MS)
        time.sleep(2)

        self._log("Hovering over Verification menu...")
        verif_menu = page.locator(SEL_VERIFICATION_MENU).first
        verif_menu.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        verif_menu.hover()
        time.sleep(1.5)

        self._log("Clicking Online Document Verification...")
        doc_link = page.locator(SEL_DOC_VERIFY_LINK).first
        doc_link.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        doc_link.click()
        time.sleep(2)

        # Verify we landed on the correct page
        if "DocumentVerification" in page.url:
            self._log("Document Verification page loaded.", "ok")
        else:
            self._log(
                f"Unexpected URL after navigation: {page.url}. "
                "Attempting direct URL fallback...", "warn"
            )
            page.goto(EFBR_DOC_VERIFICATION_URL,
                      wait_until="domcontentloaded",
                      timeout=PAGE_LOAD_TIMEOUT_MS)
            time.sleep(2)

    # ── Retry wrapper ─────────────────────────────────────────────────────

    def _download_with_retry(self, cpr: str) -> Tuple[str, str]:
        """Try downloading a CPR receipt up to MAX_RETRIES times."""
        last_error = ""
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return self._download_single_cpr(cpr)
            except Exception as exc:
                last_error = str(exc)
                self._log(
                    f"  Attempt {attempt}/{MAX_RETRIES} failed: {exc}",
                    "warn",
                )
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY_SECS)
                    self._reload_page()

        status = (STATUS_LOGIN_ISSUE if self._off_search_page()
                  else STATUS_NOT_DOWNLOADED)
        return (status, f"All {MAX_RETRIES} attempts failed: {last_error}")

    def _off_search_page(self) -> bool:
        """True if the browser is no longer on the Document Verification
        page (e.g. bounced back to the portal's login page)."""
        try:
            return "documentverification" not in self._page.url.lower()
        except Exception:
            return True

    # ── Single-CPR download ──────────────────────────────────────────────

    def _download_single_cpr(self, cpr: str) -> Tuple[str, str]:
        """
        Execute the full workflow for one CPR number.
        Returns (status, file_path_or_message).
        """
        page = self._page

        # Step 1 — Select "Computerized Payment Receipt" from dropdown
        dropdown = page.locator(SEL_SEARCH_PARAM_DROPDOWN).first
        dropdown.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        dropdown.select_option(value=CPR_OPTION_VALUE, timeout=3_000)
        time.sleep(1)

        # Step 2 — Fill CPR number into the input field
        cpr_input = page.locator(SEL_CPR_INPUT).first
        cpr_input.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        cpr_input.click()
        cpr_input.fill("")                   # clear any previous value
        cpr_input.fill(cpr.strip())          # preserve dashes exactly
        time.sleep(0.5)

        # Step 3 — Click Search and catch the popup / new tab
        search_btn = page.locator(SEL_SEARCH_BUTTON).first
        search_btn.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)

        try:
            with self._context.expect_page(timeout=POPUP_TIMEOUT_MS) as new_page_info:
                search_btn.click()
            popup_page = new_page_info.value
            popup_page.wait_for_load_state(
                "domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS
            )
            time.sleep(2)
        except Exception as exc:
            # No popup opened — might be an inline error on the same page
            body = page.locator("body").inner_text(timeout=5_000)
            if self._looks_like_not_found(body):
                return (STATUS_NOT_EXIST, "No receipt found for this CPR")
            raise RuntimeError(f"Search did not open a result popup: {exc}")

        # Step 4 — Check the popup for "no record" messages
        try:
            popup_body = popup_page.locator("body").inner_text(timeout=5_000)
        except Exception:
            popup_body = ""

        if self._looks_like_not_found(popup_body):
            popup_page.close()
            return (STATUS_NOT_EXIST, "No receipt found for this CPR")

        # Step 5 — Download the PDF via the "Click Here" link
        file_path = self._download_pdf_from_popup(popup_page, cpr)

        # Step 6 — Close popup and return to main page
        try:
            popup_page.close()
        except Exception:
            pass

        if file_path:
            return (STATUS_DOWNLOADED, file_path)
        return (STATUS_NOT_DOWNLOADED, "Could not download PDF from result page")

    # ── PDF download from popup ───────────────────────────────────────────

    def _download_pdf_from_popup(
        self, popup: Page, cpr: str
    ) -> Optional[str]:
        """
        Try multiple strategies to get the PDF from the popup page.
        Returns the saved file path, or None on failure.
        """
        safe = _safe_filename(cpr)
        fallback = f"CPR_{safe}.pdf"

        def pick_target(original: str = "") -> str:
            """Original portal filename if enabled, else CPR_<no>.pdf."""
            self.output_dir.mkdir(parents=True, exist_ok=True)
            name = fallback
            if self.keep_original_name:
                orig = _clean_original_name(original)
                if orig:
                    if not os.path.splitext(orig)[1]:
                        orig += ".pdf"
                    name = orig
                    return str(_unique_path(self.output_dir, name, cpr))
            return str(self.output_dir / name)

        # Strategy 1 — Click the "Click Here to download" link
        try:
            link = popup.locator(SEL_DOWNLOAD_LINK).first
            link.wait_for(state="visible", timeout=10_000)

            with popup.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl:
                link.click()
            target = pick_target(dl.value.suggested_filename)
            dl.value.save_as(target)
            return target
        except Exception:
            pass

        # Strategy 2 — Trigger __doPostBack('lnkBtnPrint','') directly
        try:
            with popup.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl:
                popup.evaluate("__doPostBack('lnkBtnPrint','')")
            target = pick_target(dl.value.suggested_filename)
            dl.value.save_as(target)
            return target
        except Exception:
            pass

        # Strategy 3 — Look for embedded PDF (iframe / embed / object)
        try:
            for tag, attr in [("embed", "src"), ("object", "data"),
                              ("iframe", "src")]:
                elements = popup.locator(
                    f"{tag}[type='application/pdf'], {tag}[src*='.pdf']"
                ).all()
                for el in elements:
                    src = el.get_attribute(attr) or ""
                    if src:
                        resp = popup.context.request.get(src)
                        if resp.ok:
                            target = pick_target(
                                os.path.basename(src.split("?")[0]))
                            with open(target, "wb") as f:
                                f.write(resp.body())
                            return target
        except Exception:
            pass

        return None

    # ── Helpers ───────────────────────────────────────────────────────────

    @staticmethod
    def _looks_like_not_found(text: str) -> bool:
        """Heuristic check for 'no data' messages."""
        lower = text.lower()
        indicators = ["no record", "not found", "no data", "no result",
                      "invalid", "does not exist"]
        return any(ind in lower for ind in indicators)

    def _reload_page(self):
        """Reload the Document Verification page between retries."""
        try:
            self._navigate_to_doc_verification()
        except Exception as exc:
            self._log(f"Page reload failed: {exc}", "warn")

    def _cleanup(self):
        """Shut down Playwright gracefully."""
        for obj in (self._context, self._browser):
            try:
                if obj:
                    obj.close()
            except Exception:
                pass
        try:
            if self._pw:
                self._pw.stop()
        except Exception:
            pass

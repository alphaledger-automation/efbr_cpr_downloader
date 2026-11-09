# eFBR CPR Downloader

Bulk-download Computerized Payment Receipts (CPRs) from the eFBR portal.

Reads CPR numbers from an Excel file and downloads the matching PDF receipt for each one from the Pakistan eFBR portal, then writes an Excel summary showing what happened to every CPR.

---

## Requirements

- Windows 10/11
- Python 3.10+

Dependencies are installed on first run. `run.bat` will install Python itself via `winget` if it is missing.

## Install and run

```bash
git clone https://github.com/alphaledger-automation/efbr_cpr_downloader.git
cd efbr_cpr_downloader
pip install -r requirements.txt
python main.py
```

Or double-click **`run.bat`** on Windows. On the first launch a setup window installs the missing Python packages (`customtkinter`, `playwright`, `pandas`, `openpyxl`, `xlrd`) and downloads Playwright's Chromium build (~150 MB, once). Later launches start immediately.

Prepare the downloads ahead of time, without opening the app:

```bash
python main.py --setup-only
```

`install.bat` does the same thing by double-click.

### Standalone executable

```bash
build_exe.bat
```

Builds `dist\eFBR_CPR_Downloader.exe` (2–3 minutes) using `requirements-build.txt`. The EXE does not bundle Chromium — it downloads Chromium itself on first launch.

## Input

An `.xlsx` or `.xls` file with a column of CPR numbers. Columns named `CPR`, `CPR No`, `CPR Number` and similar are auto-detected; otherwise pick the column from the dropdown.

CPR numbers look like `IT-2024-0101-01234567890`. Dashes are preserved exactly as they appear in your file — only leading and trailing whitespace is stripped.

Repeated CPR numbers are detected on load, listed in the activity log with their Excel row numbers, counted on the dashboard, and written to a `Duplicates` sheet in the summary. Each CPR is downloaded once.

## Output

Downloaded PDFs go to the folder you pick under **Save To** (remembered for next time). Default:

```
C:\Users\<YourName>\Downloads\eFBR_CPRs\
```

**File Name** — tick *Keep original file name* to save each PDF under the name the portal supplies; untick for `CPR_<number>.pdf`. If two files would collide, the CPR number is appended so nothing is overwritten.

A summary workbook `cpr_download_report.xlsx` is written after processing, one row per CPR:

| Status | Meaning |
|---|---|
| `Downloaded` | PDF saved to the output folder |
| `File not exist` | The portal has no receipt for that CPR |
| `Login issue` | The portal / verification page could not be reached |
| `Not-Downloaded` | The download failed, or the CPR was never attempted (e.g. the run was cancelled) |

## Features

- Excel column auto-detection with manual override
- Duplicate detection with a dedicated summary sheet
- Pause and Cancel mid-run
- Each CPR retried up to 2 times before being marked failed
- Completion dialog with per-status counts, plus buttons to open the summary workbook or the output folder
- Four design modes (Classic, Glass, Neumorphic, Deep) and Light / Dark appearance, switchable at any time including mid-run, remembered between launches
- Browser automation on a background thread, so the window stays responsive

## Project layout

| File | Purpose |
|---|---|
| `run.bat` | Windows launcher |
| `install.bat` | Dependency download only |
| `build_exe.bat` | PyInstaller build |
| `main.py` | Entry point; `--setup-only` |
| `bootstrap.py` | Dependency check and installer |
| `gui.py` | Main window |
| `browser_engine.py` | eFBR portal automation |
| `excel_handler.py` | Reads the input workbook |
| `report_writer.py` | Writes the Excel summary |
| `themes.py` | Design modes and palettes |
| `constants.py` | URLs, selectors, status values, preferences |

`USER_GUIDE.txt` holds the end-user walkthrough and troubleshooting notes.

## Notes

- No credentials are stored or required.
- No taxpayer data, sample workbooks or downloaded receipts are included in this repository.
- Setup log: `%LocalAppData%\eFBR_CPR_Downloader\setup.log`

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Alpha Edge Advisory.

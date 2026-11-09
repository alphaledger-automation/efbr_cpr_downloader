"""
eFBR CPR Downloader — Summary Report Writer
=============================================
Generates a styled Excel summary (cpr_download_report.xlsx) after a run.

Sheet layout:
  A  Sr No.
  B  Formatted CPRs
  C  Status            (Downloaded / File not exist / Login issue /
                        Not-Downloaded — with an in-cell dropdown)
  D  File / Remarks
  E  Timestamp
  G:H  status counts
"""

import datetime
import pathlib
from typing import List, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

from constants import (
    ALL_STATUSES,
    STATUS_DOWNLOADED, STATUS_NOT_EXIST, STATUS_LOGIN_ISSUE,
    STATUS_NOT_DOWNLOADED,
)


# ── Styles (dark teal / emerald / jungle green scheme) ───────────────────────
_HEADER_FILL   = PatternFill("solid", fgColor="00483A")
_HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
_BORDER_SIDE   = Side(style="thin", color="7FA39B")
_CELL_BORDER   = Border(
    left=_BORDER_SIDE, right=_BORDER_SIDE,
    top=_BORDER_SIDE,  bottom=_BORDER_SIDE,
)
_ALT_FILL      = PatternFill("solid", fgColor="DCEDE8")
_WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")
_TOTAL_FILL    = PatternFill("solid", fgColor="B9D9D0")

_STATUS_FONTS = {
    STATUS_DOWNLOADED:     Font(color="0B6B43", bold=True, name="Calibri"),
    STATUS_NOT_EXIST:      Font(color="8A6100", bold=True, name="Calibri"),
    STATUS_LOGIN_ISSUE:    Font(color="A93226", bold=True, name="Calibri"),
    STATUS_NOT_DOWNLOADED: Font(color="5A5A5A", bold=True, name="Calibri"),
}

# (header, key in the result dict or None for the serial number, width)
_COLUMNS = [
    ("Sr No.",         None,                   8),
    ("Formatted CPRs", "CPR No",               30),
    ("Status",         "Status",               18),
    ("File / Remarks", "Downloaded File Path", 60),
    ("Timestamp",      "Timestamp",            21),
]
_STATUS_COL = 3
_SUMMARY_COL = 7            # G:H


def _header(cell, text: str) -> None:
    cell.value     = text
    cell.font      = _HEADER_FONT
    cell.fill      = _HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _CELL_BORDER


def write_summary_report(
    results: List[dict],
    output_dir: pathlib.Path,
    duplicates: Optional[List[dict]] = None,
) -> str:
    """
    Write a styled summary Excel workbook.

    Parameters
    ----------
    results    : One dict per CPR with the keys "CPR No", "Status",
                 "Downloaded File Path" and "Timestamp".
    output_dir : Directory in which to save the report.

    Returns
    -------
    Absolute path to the generated report file.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Download Report"

    # ── Header row ────────────────────────────────────────────────────────
    for col_idx, (header, _key, width) in enumerate(_COLUMNS, start=1):
        _header(ws.cell(row=1, column=col_idx), header)
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col_idx)
        ].width = width

    ws.row_dimensions[1].height = 22

    # ── Data rows ─────────────────────────────────────────────────────────
    for row_idx, entry in enumerate(results, start=2):
        fill = _ALT_FILL if row_idx % 2 == 0 else _WHITE_FILL

        for col_idx, (_header_text, key, _width) in enumerate(_COLUMNS, start=1):
            value = row_idx - 1 if key is None else entry.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill      = fill
            cell.border    = _CELL_BORDER
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if col_idx <= _STATUS_COL else "left",
                wrap_text=col_idx > _STATUS_COL,
            )

        # Colour-code the Status column
        status_cell = ws.cell(row=row_idx, column=_STATUS_COL)
        status_cell.font = _STATUS_FONTS.get(
            str(status_cell.value).strip(),
            _STATUS_FONTS[STATUS_NOT_DOWNLOADED],
        )

    last_row = len(results) + 1

    # ── Status dropdown (so a status can be corrected by hand) ────────────
    if results:
        dropdown = DataValidation(
            type="list", formula1='"' + ",".join(ALL_STATUSES) + '"',
            allow_blank=True,
        )
        ws.add_data_validation(dropdown)
        dropdown.add(f"C2:C{last_row}")

    # ── Status counts (live formulas, G:H) ────────────────────────────────
    label_col = openpyxl.utils.get_column_letter(_SUMMARY_COL)
    ws.column_dimensions[label_col].width = 18
    ws.column_dimensions[
        openpyxl.utils.get_column_letter(_SUMMARY_COL + 1)
    ].width = 10
    _header(ws.cell(row=1, column=_SUMMARY_COL), "Summary")
    _header(ws.cell(row=1, column=_SUMMARY_COL + 1), "Count")

    status_range = f"$C$2:$C${max(last_row, 2)}"
    for offset, status in enumerate(ALL_STATUSES, start=2):
        label = ws.cell(row=offset, column=_SUMMARY_COL, value=status)
        count = ws.cell(
            row=offset, column=_SUMMARY_COL + 1,
            value=f"=COUNTIF({status_range},{label_col}{offset})",
        )
        label.font = _STATUS_FONTS[status]
        for cell in (label, count):
            cell.border = _CELL_BORDER
        count.alignment = Alignment(horizontal="center")

    total_row = len(ALL_STATUSES) + 2
    total_label = ws.cell(row=total_row, column=_SUMMARY_COL, value="Total")
    total_count = ws.cell(
        row=total_row, column=_SUMMARY_COL + 1,
        value=f"=COUNTA({status_range})",
    )
    for cell in (total_label, total_count):
        cell.font   = Font(bold=True, name="Calibri")
        cell.fill   = _TOTAL_FILL
        cell.border = _CELL_BORDER
    total_count.alignment = Alignment(horizontal="center")

    # Duplicates skipped (extra rows that were not downloaded again)
    duplicates = duplicates or []
    dup_extra = sum(d["count"] - 1 for d in duplicates)
    dup_label = ws.cell(row=total_row + 1, column=_SUMMARY_COL,
                        value="Duplicates skipped")
    dup_count = ws.cell(row=total_row + 1, column=_SUMMARY_COL + 1,
                        value=dup_extra)
    dup_label.font = _STATUS_FONTS[STATUS_NOT_EXIST]
    for cell in (dup_label, dup_count):
        cell.border = _CELL_BORDER
    dup_count.alignment = Alignment(horizontal="center")
    ws.column_dimensions[label_col].width = 22

    # ── Duplicates sheet ──────────────────────────────────────────────────
    if duplicates:
        ds = wb.create_sheet("Duplicates")
        for ci, (h, w) in enumerate(
            (("Sr No.", 8), ("CPR No", 30), ("Times in file", 14),
             ("Excel rows", 40), ("Action", 34)), start=1):
            _header(ds.cell(row=1, column=ci), h)
            ds.column_dimensions[
                openpyxl.utils.get_column_letter(ci)].width = w
        for ri, d in enumerate(duplicates, start=2):
            vals = (ri - 1, d["cpr"], d["count"],
                    ", ".join(str(r) for r in d["rows"]),
                    "Downloaded once; repeats skipped")
            for ci, v in enumerate(vals, start=1):
                c = ds.cell(row=ri, column=ci, value=v)
                c.border = _CELL_BORDER
                c.fill = _ALT_FILL if ri % 2 == 0 else _WHITE_FILL
                c.alignment = Alignment(
                    vertical="center",
                    horizontal="left" if ci >= 4 else "center")
        ds.freeze_panes = "A2"

    # ── Freeze header ─────────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Save ──────────────────────────────────────────────────────────────
    report_path = output_dir / "cpr_download_report.xlsx"
    try:
        wb.save(str(report_path))
    except PermissionError:
        # The previous report is still open in Excel — save alongside it
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = output_dir / f"cpr_download_report_{stamp}.xlsx"
        wb.save(str(report_path))
    return str(report_path)

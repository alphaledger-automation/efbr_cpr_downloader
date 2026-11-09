"""
eFBR CPR Downloader — Excel Input Handler
==========================================
Reads CPR numbers from .xlsx / .xls files.
Auto-detects the CPR column or lets the user pick one manually.
"""

from typing import List, Optional

import pandas as pd

from constants import CPR_COLUMN_ALIASES


def get_column_names(filepath: str) -> List[str]:
    """Return every column header from the first sheet of an Excel file."""
    df = pd.read_excel(filepath, nrows=0, dtype=str)
    return list(df.columns)


def detect_cpr_column(filepath: str) -> Optional[str]:
    """
    Scan column headers against known CPR aliases.
    Returns the matching column name, or None if no match found.
    """
    df = pd.read_excel(filepath, nrows=0, dtype=str)
    for col in df.columns:
        if str(col).strip().lower() in CPR_COLUMN_ALIASES:
            return col
    return None


def read_cpr_numbers(
    filepath: str,
    column_name: Optional[str] = None,
) -> List[str]:
    """
    Extract CPR numbers from an Excel file.

    Parameters
    ----------
    filepath    : Path to .xlsx or .xls file.
    column_name : Explicit column to read.  When *None*, auto-detection
                  is attempted.

    Returns
    -------
    List of CPR strings — outer whitespace stripped, blanks removed,
    dashes and internal formatting preserved exactly.

    Raises
    ------
    ValueError  If the file format is unsupported, the column is missing,
                or auto-detection fails.
    """
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsx", "xls"):
        raise ValueError(f"Unsupported file format: .{ext}")

    df = pd.read_excel(filepath, dtype=str)

    # Resolve the column
    if column_name:
        if column_name not in df.columns:
            raise ValueError(
                f"Column '{column_name}' not found.\n"
                f"Available: {', '.join(df.columns)}"
            )
        col = column_name
    else:
        col = detect_cpr_column(filepath)
        if col is None:
            raise ValueError(
                "Could not auto-detect the CPR column.\n"
                f"Available columns: {', '.join(df.columns)}"
            )

    # Clean and return
    cprs = (
        df[col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .tolist()
    )
    return cprs


def analyze_cprs(filepath: str, column_name: Optional[str] = None) -> dict:
    """
    Read CPRs and split them into unique values and duplicates.

    Duplicate test: same CPR after stripping outer whitespace and ignoring
    case.  The FIRST occurrence is kept (original text/dashes preserved),
    later ones are reported as duplicates and never downloaded.

    Returns
    -------
    {
      "all":        [every CPR in file order],
      "unique":     [first occurrence of each CPR, in file order],
      "duplicates": [{"cpr": str, "count": int, "rows": [excel row numbers]}],
      "dup_extra":  number of redundant rows skipped (sum of count-1),
    }
    """
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsx", "xls"):
        raise ValueError(f"Unsupported file format: .{ext}")

    df = pd.read_excel(filepath, dtype=str)
    if column_name:
        if column_name not in df.columns:
            raise ValueError(
                f"Column '{column_name}' not found.\n"
                f"Available: {', '.join(df.columns)}"
            )
        col = column_name
    else:
        col = detect_cpr_column(filepath)
        if col is None:
            raise ValueError(
                "Could not auto-detect the CPR column.\n"
                f"Available columns: {', '.join(df.columns)}"
            )

    seen = {}                   # key -> {"cpr", "rows"}
    order = []
    all_cprs = []
    for idx, raw in df[col].items():
        if raw is None or str(raw).strip() in ("", "nan"):
            continue
        cpr = str(raw).strip()
        row_no = int(idx) + 2   # +1 header, +1 for 1-based Excel rows
        all_cprs.append(cpr)
        key = cpr.upper()
        if key not in seen:
            seen[key] = {"cpr": cpr, "rows": [row_no]}
            order.append(key)
        else:
            seen[key]["rows"].append(row_no)

    unique = [seen[k]["cpr"] for k in order]
    duplicates = [
        {"cpr": seen[k]["cpr"], "count": len(seen[k]["rows"]),
         "rows": seen[k]["rows"]}
        for k in order if len(seen[k]["rows"]) > 1
    ]
    return {
        "all": all_cprs,
        "unique": unique,
        "duplicates": duplicates,
        "dup_extra": len(all_cprs) - len(unique),
    }

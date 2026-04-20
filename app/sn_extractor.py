"""
sn_extractor.py
---------------
Reads the M&V Plan Review Sheet Excel template at startup and returns the
authoritative list of question SNs — all SNs in column B from row 22 onwards
that are NOT whole integers (whole integers = section header rows).
"""

from typing import List

import openpyxl


def _is_whole_integer(value) -> bool:
    """Return True if the value is a whole-integer section header (0, 1 … 17)."""
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    try:
        f = float(s)
        return f == int(f) and "." not in s
    except (ValueError, TypeError):
        return False  # unparseable as float means it's a string SN like '6.3.1' → keep it


def extract_expected_sns(template_path: str) -> List[str]:
    """Return question SNs from Sheet '1. M&V plan_V2.0'."""
    return extract_expected_sns_for_sheet(template_path, "1. M&V plan_V2.0")


def extract_expected_sns_for_sheet(template_path: str, sheet_name: str) -> List[str]:
    """
    Open the Excel template and return the authoritative list of question SNs
    for the given sheet name.

    Rules:
    - Scan column B from row 22 onwards.
    - Skip blank cells and whole-integer rows (section headers).
    - Skip any SN that is a dot-prefix of another SN (sub-section headers).

    Returns a list like ["0.1", "1.1", ..., "6.3.1"].
    """
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    ws = wb[sheet_name]

    candidates: List[str] = []
    for row in ws.iter_rows(min_row=22, min_col=2, max_col=2, values_only=True):
        value = row[0]
        if _is_whole_integer(value):
            continue
        candidates.append(str(value).strip())

    wb.close()

    candidate_set = set(candidates)
    questions = [
        sn for sn in candidates
        if not any(other.startswith(sn + ".") for other in candidate_set)
    ]
    return questions

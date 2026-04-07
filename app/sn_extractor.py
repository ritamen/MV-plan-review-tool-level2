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
    """
    Open the Excel template and return the authoritative list of question SNs.

    Rules:
    - Scan column B from row 22 onwards.
    - Skip blank cells and whole-integer rows (section headers 0–17).
    - Skip any SN that is a dot-prefix of another SN in the sheet
      (e.g. '6.3' is skipped because '6.3.1' exists → it's a sub-section header).

    Returns a list like ["0.1", "1.1", ..., "6.3.1", ..., "17.1"].
    """
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    ws = wb["1. M&V plan_V2.0"]

    # First pass: collect every non-blank, non-integer SN as a string
    candidates: List[str] = []
    for row in ws.iter_rows(min_row=22, min_col=2, max_col=2, values_only=True):
        value = row[0]
        if _is_whole_integer(value):
            continue
        candidates.append(str(value).strip())

    wb.close()

    # Second pass: drop any SN that is a dot-prefix of another SN
    # e.g. "6.3" is dropped because "6.3.1" starts with "6.3."
    candidate_set = set(candidates)
    questions = [
        sn for sn in candidates
        if not any(other.startswith(sn + ".") for other in candidate_set)
    ]
    return questions

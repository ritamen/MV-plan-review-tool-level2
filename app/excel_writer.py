"""
excel_writer.py
---------------
All openpyxl logic for writing AI review results into the M&V Plan Review Sheet.
The AI never sees or touches this file — Python handles everything.

Sheet  : "1. M&V plan_V2.0"
Columns written (1-based):
  H  col 8  = Included
  I  col 9  = Active Status
  J  col 10 = Consultant's Comments Round 1

Colour coding:
  Yes  / Approved     -> bg #C6EFCE  font #375623  bold  centered
  No   / Not Approved -> bg #FFC7CE  font #9C0006  bold  centered
  Partial / Incomplete-> bg #FFEB9C  font #9C6500  bold  centered

  Comment non-empty   -> bg #FFFF99  font #000000  wrap  top-aligned
  Comment empty       -> bg #FFFFFF
  Row height          -> max(40, (len//80 + newlines + 1) * 15)

  All written cells: Arial 10, thin border #BBBBBB.
"""

import io
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Column indices (1-based) ─────────────────────────────────────────────────
COL_SN      = 2   # B
COL_INCL    = 8   # H  Included
COL_STATUS  = 9   # I  Active Status
COL_COMMENT = 10  # J  Consultant's Comments Round 1
DATA_START  = 22  # First data row

FONT_NAME = "Arial"
FONT_SIZE = 10
BORDER_COLOR = "BBBBBB"

# ── Style maps ───────────────────────────────────────────────────────────────
INCLUDED_STYLES = {
    "Yes":     {"bg": "C6EFCE", "fc": "375623"},
    "No":      {"bg": "FFC7CE", "fc": "9C0006"},
    "Partial": {"bg": "FFEB9C", "fc": "9C6500"},
}

STATUS_STYLES = {
    "Approved":     {"bg": "C6EFCE", "fc": "375623"},
    "Not Approved": {"bg": "FFC7CE", "fc": "9C0006"},
    "Incomplete":   {"bg": "FFEB9C", "fc": "9C6500"},
}

COMMENT_BG = "FFFF99"
COMMENT_FC = "000000"


def _make_border() -> Border:
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def _make_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def _style_cell(cell, value: str, bg: str, fc: str,
                bold: bool = False, wrap: bool = False,
                align: str = "left") -> None:
    cell.value = value
    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color=fc, bold=bold)
    cell.fill = _make_fill(bg)
    cell.alignment = Alignment(horizontal=align, vertical="top", wrap_text=wrap)
    cell.border = _make_border()


def _is_section_header(value) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    try:
        f = float(s)
        return f == int(f) and "." not in s
    except (ValueError, TypeError):
        return False  # unparseable as float → multi-part SN like "6.3.1", treat as question row


def write_review(template_bytes: bytes, review_by_sn: dict,
                 ref_no: str = "", esp_name: str = "", facility_name: str = "") -> bytes:
    """
    Load the Excel template from bytes, write review results, return as bytes.

    Parameters
    ----------
    template_bytes : bytes
        Raw bytes of the M&V Plan Review Sheet template.
    review_by_sn  : dict
        Dict keyed by SN string -> {"included": ..., "status": ..., "comment": ...}
    ref_no        : str  Reference number written to the review sheet and cover page.
    esp_name      : str  ESP name written to the cover page.
    facility_name : str  Facility name written to the review sheet.

    Returns
    -------
    bytes
        In-memory bytes of the filled workbook (xlsx).
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), keep_links=False)

    # Drop external links, broken named ranges, and legacy drawings that
    # openpyxl cannot round-trip cleanly.
    wb._external_links.clear()

    # Remove named ranges that reference external workbooks (contain "[")
    # — they become invalid once external links are cleared.
    broken = [
        name for name in wb.defined_names
        if "[" in (wb.defined_names[name].attr_text or "")
    ]
    for name in broken:
        del wb.defined_names[name]

    for sheet in wb.worksheets:
        sheet._charts.clear()
        sheet._images.clear()
        if hasattr(sheet, "_drawing") and sheet._drawing is not None:
            sheet._drawing = None

    ws = wb["1. M&V plan_V2.0"]

    # ── Fill review date fields ───────────────────────────────────────────────
    today_str = date.today().strftime("%d/%m/%y")

    # "Last Updated: dd/mm/yy" — find the cell in row 5 that contains the text
    for cell in ws[5]:
        if cell.value and "Last Updated" in str(cell.value):
            cell.value = f"Last Updated: {today_str}"
            break

    # ── Facility Name (row 6, col C) and Ref. No. (row 7, col C) ─────────────
    if facility_name:
        ws.cell(row=6, column=3).value = facility_name
    if ref_no:
        ws.cell(row=7, column=3).value = ref_no

    # ── Cover Page ────────────────────────────────────────────────────────────
    if "Cover Page" in wb.sheetnames:
        cp = wb["Cover Page"]
        # Date (row 12, col B)
        cp.cell(row=12, column=2).value = today_str
        # ESP name (row 13, col B)
        if esp_name:
            cp.cell(row=13, column=2).value = esp_name
        # Date of Last Status (row 16, col F)
        cp.cell(row=16, column=6).value = today_str

    # Round 1 row (row 15): Issued On (col C=3), Received on (col D=4), Reviewed on (col E=5)
    for col in (3, 4, 5):
        cell = ws.cell(row=15, column=col)
        cell.value = today_str

    # Round 1 Assessment (col F=6): Approved only if every question is Approved
    all_approved = all(item.get("status") == "Approved" for item in review_by_sn.values())
    assessment = "Approved" if all_approved else "Not Approved"
    assess_s = STATUS_STYLES[assessment]
    _style_cell(
        ws.cell(row=15, column=6),
        assessment, bg=assess_s["bg"], fc=assess_s["fc"],
        bold=True, align="center"
    )

    for row_idx in range(DATA_START, ws.max_row + 1):
        sn_val = ws.cell(row=row_idx, column=COL_SN).value
        if _is_section_header(sn_val):
            continue

        sn = str(sn_val).strip()
        if sn not in review_by_sn:
            continue

        item     = review_by_sn[sn]
        included = item.get("included", "")
        status   = item.get("status", "")
        comment  = item.get("comment", "") or ""

        # ── Included (col H) ────────────────────────────────────────────────
        inc_s = INCLUDED_STYLES.get(included, {"bg": "FFFFFF", "fc": "000000"})
        _style_cell(
            ws.cell(row=row_idx, column=COL_INCL),
            included, bg=inc_s["bg"], fc=inc_s["fc"],
            bold=True, align="center"
        )

        # ── Active Status (col I) ────────────────────────────────────────────
        st_s = STATUS_STYLES.get(status, {"bg": "FFFFFF", "fc": "000000"})
        _style_cell(
            ws.cell(row=row_idx, column=COL_STATUS),
            status, bg=st_s["bg"], fc=st_s["fc"],
            bold=True, align="center"
        )

        # ── Consultant's Comment (col J) ─────────────────────────────────────
        if comment:
            _style_cell(
                ws.cell(row=row_idx, column=COL_COMMENT),
                comment, bg=COMMENT_BG, fc=COMMENT_FC,
                wrap=True, align="left"
            )
            lines = len(comment) // 80 + comment.count("\n") + 1
            ws.row_dimensions[row_idx].height = max(40, lines * 15)
        else:
            _style_cell(
                ws.cell(row=row_idx, column=COL_COMMENT),
                "", bg="FFFFFF", fc=COMMENT_FC,
                wrap=True, align="left"
            )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

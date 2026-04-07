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
        return True


def write_review(template_bytes: bytes, review_by_sn: dict) -> bytes:
    """
    Load the Excel template from bytes, write review results, return as bytes.

    Parameters
    ----------
    template_bytes : bytes
        Raw bytes of the M&V Plan Review Sheet template.
    review_by_sn  : dict
        Dict keyed by SN string -> {"included": ..., "status": ..., "comment": ...}

    Returns
    -------
    bytes
        In-memory bytes of the filled workbook (xlsx).
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb["1. M&V plan_V2.0"]

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

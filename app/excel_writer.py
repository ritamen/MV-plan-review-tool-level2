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
  Yes  / Approved     -> bg #00B050  font #FFFFFF  bold  centered
  No   / Not Approved -> bg #FF0000  font #FFFFFF  bold  centered
  Partial / Incomplete-> bg #FFFF00  font #000000  bold  centered

  Comment non-empty   -> bg #FFFF99  font #000000  wrap  top-aligned
  Comment empty       -> bg #FFFFFF
  Row height          -> max(40, (len//80 + newlines + 1) * 15)

  All written cells: Trebuchet MS 10, thin border #BBBBBB.
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

FONT_NAME = "Trebuchet MS"
FONT_SIZE = 10
BORDER_COLOR = "BBBBBB"

# ── Style maps ───────────────────────────────────────────────────────────────
INCLUDED_STYLES = {
    "Yes":     {"bg": "E2EFDA", "fc": "00B050"},
    "No":      {"bg": "FFD3D3", "fc": "C00000"},
    "Partial": {"bg": "FFF2CC", "fc": "C65911"},
}

STATUS_STYLES = {
    "Approved":     {"bg": "00B050", "fc": "FFFFFF"},
    "Not Approved": {"bg": "FF0000", "fc": "FFFFFF"},
    "Incomplete":   {"bg": "FFFF00", "fc": "000000"},
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
    vertical = "center" if align == "center" else "top"
    cell.alignment = Alignment(horizontal=align, vertical=vertical, wrap_text=wrap)
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


def _build_regression_comment(regression_results: list) -> str:
    """
    Build a natural-language paragraph appended to SN 6.3.6 stating the
    Python-computed regression values, whether they match reported stats,
    and whether they meet IPMVP/ASHRAE Guideline 14 thresholds.
    """
    if not regression_results:
        return ""

    blocks = []
    for res in regression_results:
        name = res.get("eem_name", "Unknown EEM")

        if res.get("error"):
            blocks.append(
                f"Independent Python regression verification ({name}) could not be completed: {res['error']}."
            )
            continue

        c          = res.get("computed") or {}
        thresholds = res.get("thresholds") or {}
        comparison = res.get("comparison") or {}

        r2        = c.get("r_squared", 0)
        cvrmse    = c.get("cv_rmse", 0)
        nmbe      = c.get("nmbe", 0)
        t_stat    = c.get("t_stat", 0)
        p_val     = c.get("p_value", 0)
        it        = c.get("intercept_t_stat")
        ip        = c.get("intercept_p_value")
        intercept = c.get("intercept", 0)
        slope     = c.get("slope", 0)
        mse       = c.get("model_std_err", 0)
        n         = c.get("n", 0)

        p_str  = "< 0.001" if p_val < 0.001 else f"= {p_val:.3f}"
        ip_str = ("< 0.001" if ip < 0.001 else f"= {ip:.4f}") if ip is not None else "N/A"
        it_str = f"{it:.4f}" if it is not None else "N/A"

        # ── Threshold split ───────────────────────────────────────────────────
        failing = [k for k, v in thresholds.items() if not v.get("passes")]
        passing = [k for k, v in thresholds.items() if v.get("passes")]

        # ── Match sentence ────────────────────────────────────────────────────
        has_reported = any(v.get("reported") is not None for v in comparison.values())
        if has_reported:
            mismatches = res.get("stats_mismatch", [])
            match_clause = (
                "consistent with those reported in the M&V Plan"
                if not mismatches else
                f"differing from the M&V Plan on {', '.join(mismatches)}"
            )
        else:
            match_clause = "consistent with those reported in the M&V Plan"

        # ── Threshold clause ──────────────────────────────────────────────────
        if not failing:
            threshold_clause = "all meet IPMVP/ASHRAE Guideline 14 thresholds."
        else:
            pass_part = f"{', '.join(passing)} {'meet' if len(passing) > 1 else 'meets'} the required thresholds" if passing else ""
            fail_part = f"{', '.join(failing)} {'do' if len(failing) > 1 else 'does'} not meet the required threshold"
            threshold_clause = f"{pass_part + '; ' if pass_part else ''}{fail_part}."

        # ── Assemble paragraph ────────────────────────────────────────────────
        para = (
            f"Independent Python verification ({name}, n = {n}) yields: "
            f"R² = {r2:.4f}, CV(RMSE) = {cvrmse:.2f}%, NMBE = {nmbe:.2f}%, "
            f"slope t-statistic = {t_stat:.4f} (p {p_str}), "
            f"Model Standard Error = {mse:,.2f}, intercept = {intercept:,.2f}, slope = {slope:.6f}. "
            f"These values are {match_clause}, and {threshold_clause}"
        )

        blocks.append(para)

    return "\n\n" + "\n\n".join(blocks)


def write_review(template_bytes: bytes, review_by_sn: dict,
                 ref_no: str = "", client_name: str = "", esp_name: str = "",
                 facility_name: str = "", regression_results: list = None,
                 regression_data_provided: bool = False) -> bytes:
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

    # ── ESP Name (row 10, col D) and Ref. No. (row 11, col D) ───────────────
    if esp_name:
        ws.cell(row=10, column=4).value = esp_name
    if ref_no:
        ws.cell(row=11, column=4).value = ref_no

    # ── Cover Page ────────────────────────────────────────────────────────────
    if "Cover Page" in wb.sheetnames:
        cp = wb["Cover Page"]
        # Client name (row 10, col A)
        if client_name:
            cp.cell(row=10, column=1).value = client_name
        # Project name (row 11, col A)
        if facility_name:
            cp.cell(row=11, column=1).value = f'Energy Efficiency Retrofit Project for "{facility_name}"'
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

    # Build regression verification text once — appended to SN 6.3.6 comment
    regression_block = _build_regression_comment(regression_results or [])

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

        # Append Python regression verification to the 6.3.6 comment
        if sn == "6.3.6":
            if regression_block:
                comment = (comment + regression_block).strip()
            elif included in ("Yes", "Partial") and not regression_data_provided:
                not_provided = (
                    "\n\nRegression data was not provided; independent Python "
                    "verification of the reported regression statistics could not be performed."
                )
                comment = (comment + not_provided).strip()

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

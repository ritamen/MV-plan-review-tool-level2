# streamlit_app.py  –  M&V Plan Review Tool
import base64
import io
import json
import logging
import os
import re
import sys
import traceback
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv
from pypdf import PdfReader

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent / "app"
PROMPT_PATH   = BASE_DIR / "assets" / "MV_Plan_Reviewer_Prompt.txt"
TEMPLATE_PATH = BASE_DIR / "assets" / "M_V_Plan_Review_Sheet.xlsx"
LOGO_PATH     = BASE_DIR / "static" / "arklogo2.png"

load_dotenv(BASE_DIR.parent / ".env")

sys.path.insert(0, str(BASE_DIR))
from excel_writer import write_review
from sn_extractor import extract_expected_sns

# ── Load assets once ──────────────────────────────────────────────────────────
REVIEWER_PROMPT: str   = PROMPT_PATH.read_text(encoding="utf-8")
TEMPLATE_BYTES:  bytes = TEMPLATE_PATH.read_bytes()
EXPECTED_SNS           = extract_expected_sns(str(TEMPLATE_PATH))

# ── API constants ─────────────────────────────────────────────────────────────
MODEL           = "claude-opus-4-5"
MAX_TOKENS      = 8000
THINKING_TOKENS = 5000
TIMEOUT_SECS    = 180
MAX_PDF_PAGES   = 100

VALID_INCLUDED = {"Yes", "No", "Partial"}
VALID_STATUS   = {"Approved", "Not Approved", "Incomplete"}

logging.basicConfig(level=logging.INFO)


# ── Backend helpers ────────────────────────────────────────────────────────────

def img_to_base64(path: str) -> str:
    return base64.b64encode(Path(path).read_bytes()).decode("utf-8")

def _encode_pdf(data: bytes) -> str:
    return base64.standard_b64encode(data).decode("utf-8")

def _pdf_page_count(pdf_bytes: bytes) -> int:
    return len(PdfReader(io.BytesIO(pdf_bytes)).pages)

def _pdf_to_text(pdf_bytes: bytes, label: str) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for i, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ""
        pages.append(f"--- {label} | Page {i} ---\n{text}")
    return "\n\n".join(pages)

def _strip_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()

def _extract_json_text(response) -> str:
    for block in response.content:
        if block.type == "text":
            return block.text
    raise ValueError("No text content block found in API response.")

def _validate_items(items: list) -> list:
    errors = []
    for i, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"Item {i} is not a dict"); continue
        for field in ("sn", "included", "status", "comment"):
            if field not in item:
                errors.append(f"Item {i} missing field '{field}'")
        if "included" in item and item["included"] not in VALID_INCLUDED:
            errors.append(f"Item {i} invalid included={item['included']!r}")
        if "status" in item and item["status"] not in VALID_STATUS:
            errors.append(f"Item {i} invalid status={item['status']!r}")
    return errors

def _parse_and_validate(raw: str):
    cleaned = _strip_fences(raw)
    try:
        items = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        return [], [f"JSON parse error: {exc}"]
    if not isinstance(items, list):
        return [], ["Response is not a JSON array"]
    return items, _validate_items(items)

def _add_pdf_to_content(content: list, pdf_bytes: bytes, label: str) -> None:
    pages = _pdf_page_count(pdf_bytes)
    if pages <= MAX_PDF_PAGES:
        content.append({
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": _encode_pdf(pdf_bytes)},
            "title": label,
        })
    else:
        extracted = _pdf_to_text(pdf_bytes, label)
        content.append({
            "type": "text",
            "text": f"=== {label} ({pages} pages — full text extracted) ===\n\n{extracted}",
        })

def _build_user_content(mv_bytes: bytes, supporting_bytes: list):
    content = []
    _add_pdf_to_content(content, mv_bytes, "M&V Plan")
    for i, b in enumerate(supporting_bytes or [], 1):
        _add_pdf_to_content(content, b, f"Supporting Document {i}")
    content.append({
        "type": "text",
        "text": (
            "Review the M&V Plan against every question in your instructions. "
            "Return a single valid JSON array only — no markdown, no explanation, "
            "no text outside the array. "
            "Each element: { sn, included, status, comment }."
        ),
    })
    return content

def _call_claude(client, user_content: list) -> str:
    import anthropic
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": THINKING_TOKENS},
        system=REVIEWER_PROMPT,
        messages=[{"role": "user", "content": user_content}],
        timeout=TIMEOUT_SECS,
    )
    return _extract_json_text(response)

def _call_claude_retry(client, user_content: list, first_raw: str) -> str:
    import anthropic
    retry_messages = [
        {"role": "user",      "content": user_content},
        {"role": "assistant", "content": first_raw},
        {"role": "user",      "content": (
            "Your previous response was invalid. Return ONLY a JSON array. "
            "Each item: sn (string), included (Yes/No/Partial), "
            "status (Approved/Not Approved/Incomplete — not Approved as Noted), "
            "comment (string, empty if Approved). Nothing else."
        )},
    ]
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": THINKING_TOKENS},
        system=REVIEWER_PROMPT,
        messages=retry_messages,
        timeout=TIMEOUT_SECS,
    )
    return _extract_json_text(response)

def run_mv_review(mv_bytes, supporting_bytes, ref_no, esp_name, mv_filename, debug_chunks=False):
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    client = anthropic.Anthropic(api_key=api_key)
    user_content = _build_user_content(mv_bytes, supporting_bytes)

    raw = _call_claude(client, user_content)
    items, errors = _parse_and_validate(raw)

    if errors:
        raw = _call_claude_retry(client, user_content, raw)
        items, errors = _parse_and_validate(raw)
        if errors:
            raise RuntimeError(
                "The AI returned an invalid response after retry. Please re-run the review."
            )

    review_by_sn = {str(item["sn"]).strip(): item for item in items}
    missing_sns  = sorted(set(EXPECTED_SNS) - set(review_by_sn.keys()))

    approved     = sum(1 for it in items if it.get("status") == "Approved")
    not_approved = sum(1 for it in items if it.get("status") == "Not Approved")
    incomplete   = sum(1 for it in items if it.get("status") == "Incomplete")
    total        = len(items)

    filled_bytes = write_review(TEMPLATE_BYTES, review_by_sn)

    base_name = mv_filename.replace(".pdf", "")
    parts = ["MV_Plan_Review"]
    if ref_no:   parts.append(ref_no.replace(" ", "_"))
    if esp_name: parts.append(esp_name.replace(" ", "_"))
    if not (ref_no or esp_name): parts.append(base_name)
    output_filename = "_".join(parts) + ".xlsx"

    return {
        "total":        total,
        "approved":     approved,
        "not_approved": not_approved,
        "incomplete":   incomplete,
        "missing_sns":  missing_sns,
        "excel_bytes":  filled_bytes,
        "filename":     output_filename,
    }


class StreamlitLogger:
    def __init__(self, container):
        self.container = container
        self.lines = []

    def log(self, msg: str):
        self.lines.append(msg)
        self.container.text("\n".join(self.lines[-250:]))


# ============================================================
# UI  –  identical to streamlit_app 1.py
# ============================================================

# ---------------- Page config ----------------
st.set_page_config(page_title="ARK Energy | M&V Plan Review Tool", layout="wide")

# ---------------- ARK Theme CSS ----------------
st.markdown(
    """
    <style>
    :root{
      --ark-blue: #0D6079;
      --ark-orange: #F79428;
      --ark-black: #000000;
      --card-border: rgba(13,96,121,0.18);
      --shadow: 0 6px 18px rgba(0,0,0,0.08);
    }

    html, body, [class*="css"], * {
        font-family: "Trebuchet MS", Arial, sans-serif !important;
        color: var(--ark-black);
    }
    button, input, textarea, select, label, p, div, span, li {
        font-family: "Trebuchet MS", Arial, sans-serif !important;
    }

    header[data-testid="stHeader"] { display: none; }
    footer { display: none; }

    .block-container {
        padding-top: 6.2rem !important;
        padding-bottom: 1.2rem !important;
        max-width: 98vw !important;
    }

    .ark-nav {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        z-index: 9999;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
        padding: 12px 18px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.35);
    }

    .ark-nav-inner{
        width: 98vw;
        margin: 0 auto;
        border-radius: 14px;
        padding: 10px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
    }

    .ark-nav-left { display:flex; align-items:center; gap:14px; }

    .ark-nav-title {
        color: white !important;
        font-size: 22px !important;
        font-weight: 900;
        line-height: 1.2;
        margin: 0;
    }

    .pill {
        border-radius: 999px;
        padding: 8px 14px;
        font-size: 14px;
        font-weight: 900;
        border: 1px solid rgba(255,255,255,0.25);
        color: white !important;
        background: transparent;
        white-space: nowrap;
    }

    .ark-section { margin-top:10px; margin-bottom:6px; display:flex; align-items:baseline; gap:10px; }
    .ark-section-title { font-size:18px; font-weight:900; color:var(--ark-blue); margin:0; line-height:1; }
    .ark-section-rule { height:2px; background:rgba(13,96,121,0.25); width:100%; margin-top:8px; margin-bottom:24px; }

    /* style only widget labels, not internal uploader button labels */
    [data-testid="stWidgetLabel"] {
        font-size: 15px !important;
        font-weight: 700 !important;
    }

    div.stButton > button[kind="primary"],
    div.stButton > button[kind="primary"] * {
        color: #FFFFFF !important;
    }
    div.stButton > button[kind="primary"] {
        background-color: var(--ark-orange) !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        border-radius: 14px !important;
        padding: 14px 28px !important;
        height: 64px !important;
        border: none !important;
        box-shadow: 0 8px 20px rgba(247,148,40,0.25) !important;
        width: 100% !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stButton > button[kind="primary"]:hover * {
        background-color: var(--ark-blue) !important;
        color: #FFFFFF !important;
    }

    .stat-card {
        background: white; border-radius: 12px; padding: 18px 16px;
        text-align: center; box-shadow: 0 3px 10px rgba(0,0,0,0.06); margin-bottom: 8px;
    }
    .stat-number { font-size: 36px; font-weight: 900; line-height: 1; margin-bottom: 6px; }
    .stat-label  { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: #555; }
    .color-blue   { color: #0D6079; }
    .color-green  { color: #375623; }
    .color-red    { color: #9C0006; }
    .color-orange { color: #9C6500; }

    /* Fix duplicate "upload" text caused by Material Icon falling back to literal text */
    [data-testid="stFileUploaderDropzone"] button [data-testid="stIconMaterial"] {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------- Fixed Header ----------------
logo_path = str(LOGO_PATH)
logo_b64 = img_to_base64(logo_path) if Path(logo_path).exists() else ""

st.markdown(
    f"""
    <div class="ark-nav">
      <div class="ark-nav-inner">
        <div class="ark-nav-left">
          <img src="data:image/png;base64,{logo_b64}" style="height:68px; width:auto; display:block;" />
          <div>
            <div class="ark-nav-title">M&amp;V Plan Review Tool</div>
          </div>
        </div>
        <div class="ark-nav-right">
            <div class="pill">AI Assistant</div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------- Upload Section ----------------
st.markdown(
    """
    <div class="ark-section">
      <div class="ark-section-title">Upload documents</div>
    </div>
    <div class="ark-section-rule"></div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)

with col1:
    main_pdf_upload = st.file_uploader(
        "Main document to be reviewed",
        type=["pdf"],
        accept_multiple_files=False,
        help="This is the document that will be reviewed.",
    )

with col2:
    supporting_uploads = st.file_uploader(
        "Supporting documents",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload any supporting PDFs (standards, scope, annexes, etc.).",
    )

# ---------------- Submission details ----------------
st.markdown(
    """
    <div class="ark-section">
      <div class="ark-section-title">Submission details</div>
    </div>
    <div class="ark-section-rule"></div>
    """,
    unsafe_allow_html=True,
)

ref_no   = st.text_input("Ref. No.", value="")
esp_name = st.text_input("ESP's Name", value="")

debug_chunks = False

# ---------------- Generate button ----------------
btn_left, btn_right = st.columns([7, 3])
with btn_right:
    run_btn = st.button(
        "Generate Comments",
        type="primary",
        disabled=(main_pdf_upload is None),
        use_container_width=True,
    )

# ---------------- Run ----------------
if run_btn:
    if main_pdf_upload is None:
        st.warning("Please upload the main document PDF.")
        st.stop()

    log_box = st.empty()
    logger  = StreamlitLogger(log_box)

    try:
        mv_bytes = main_pdf_upload.read()
        supporting_bytes = [f.read() for f in (supporting_uploads or [])]

        main_name = main_pdf_upload.name
        logger.log(f"📄 Reviewing: {main_name} with {len(supporting_bytes)} supporting document(s).")

        if debug_chunks:
            logger.log("🐛 Debug chunks ON")

        with st.spinner(f"Generating comments for {main_name}…"):
            result = run_mv_review(
                mv_bytes,
                supporting_bytes,
                ref_no,
                esp_name,
                main_name,
                debug_chunks=debug_chunks,
            )

        logger.log(f"✅ Done. {result['total']} questions reviewed.")

        st.markdown(
            """
            <div class="ark-section">
            <div class="ark-section-title">Results</div>
            </div>
            <div class="ark-section-rule"></div>
            """,
            unsafe_allow_html=True,
        )

        st.success("Review completed successfully.")

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(
            f'<div class="stat-card"><div class="stat-number color-blue">{result["total"]}</div>'
            f'<div class="stat-label">Questions Reviewed</div></div>',
            unsafe_allow_html=True,
        )
        c2.markdown(
            f'<div class="stat-card"><div class="stat-number color-green">{result["approved"]}</div>'
            f'<div class="stat-label">Approved</div></div>',
            unsafe_allow_html=True,
        )
        c3.markdown(
            f'<div class="stat-card"><div class="stat-number color-red">{result["not_approved"]}</div>'
            f'<div class="stat-label">Not Approved</div></div>',
            unsafe_allow_html=True,
        )
        c4.markdown(
            f'<div class="stat-card"><div class="stat-number color-orange">{result["incomplete"]}</div>'
            f'<div class="stat-label">Incomplete</div></div>',
            unsafe_allow_html=True,
        )

        if result["missing_sns"]:
            st.warning(
                f"**Warning — missing SNs:** The following questions were not returned by the AI "
                f"and have been left blank in the output: **{', '.join(result['missing_sns'])}**"
            )

        st.download_button(
            "⬇️ Download Excel Output",
            data=result["excel_bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"{type(e).__name__}: {e}")

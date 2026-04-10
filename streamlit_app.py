# streamlit_app.py  –  M&V Plan Review Tool
import base64
import io
import json
import logging
import os
import re
import sys
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
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
MODEL           = "claude-sonnet-4-6"
MAX_TOKENS      = 8000
THINKING_TOKENS = 5000
TIMEOUT_SECS    = 180
MAX_PDF_PAGES   = 100
MAX_SUPPORTING_TEXT_CHARS = 400_000

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

def _pdf_to_text(pdf_bytes: bytes, label: str, max_chars: int | None = None) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    parts = []
    chars_so_far = 0
    stopped_at = None
    for i, page in enumerate(reader.pages, 1):
        chunk = f"--- {label} | Page {i} ---\n{page.extract_text() or ''}"
        if max_chars is not None and chars_so_far + len(chunk) > max_chars:
            stopped_at = i
            break
        parts.append(chunk)
        chars_so_far += len(chunk)
    result = "\n\n".join(parts)
    if stopped_at is not None:
        omitted = total_pages - stopped_at + 1
        result += (
            f"\n\n[NOTE: {label} was truncated after page {stopped_at - 1} of "
            f"{total_pages} ({chars_so_far:,} chars extracted). "
            f"{omitted} page(s) omitted to stay within the token limit.]"
        )
    return result

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

def _add_supporting_to_content(content: list, pdf_bytes: bytes, label: str) -> None:
    pages = _pdf_page_count(pdf_bytes)
    extracted = _pdf_to_text(pdf_bytes, label, max_chars=MAX_SUPPORTING_TEXT_CHARS)
    content.append({
        "type": "text",
        "text": f"=== {label} ({pages} pages — text extracted) ===\n\n{extracted}",
    })

def _build_user_content(mv_bytes: bytes, supporting_bytes: list):
    content = []
    _add_pdf_to_content(content, mv_bytes, "M&V Plan")
    for i, b in enumerate(supporting_bytes or [], 1):
        _add_supporting_to_content(content, b, f"Supporting Document {i}")
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

def run_mv_review(mv_bytes, supporting_bytes, ref_no, client_name, esp_name, mv_filename, facility_name="", debug_chunks=False):
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    client = anthropic.Anthropic(api_key=api_key)
    user_content = _build_user_content(mv_bytes, supporting_bytes)

    try:
        raw = _call_claude(client, user_content)
    except anthropic.BadRequestError as exc:
        msg = str(exc)
        if "prompt is too long" in msg or "token" in msg.lower():
            raise RuntimeError(
                "The combined documents are too large for the AI to process "
                "(exceeded the 200,000-token limit). "
                "Try uploading a shorter IGA document or split it into smaller sections."
            ) from exc
        raise

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

    filled_bytes = write_review(TEMPLATE_BYTES, review_by_sn,
                               ref_no=ref_no, client_name=client_name,
                               esp_name=esp_name, facility_name=facility_name)

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


def _extract_submission_metadata(pdf_bytes: bytes) -> dict:
    """
    Extract Ref. No., client name, ESP name, and Facility Name from the first
    pages of the M&V Plan PDF using a fast Claude call.  Returns a dict with
    keys ref_no, client_name, esp_name, facility_name (empty strings if not found).
    """
    import anthropic
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return {}
    try:
        text = _pdf_to_text(pdf_bytes, "M&V Plan", max_chars=12_000)
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=400,
            messages=[{
                "role": "user",
                "content": (
                    "Read the following text extracted from an M&V Plan document "
                    "and extract exactly four fields.\n"
                    "Return ONLY a JSON object — no markdown, no explanation — with "
                    "these keys:\n"
                    "  ref_no        – the reference / document number\n"
                    "  client_name   – the name of the client or owner (the organisation that commissioned the work, not the ESP)\n"
                    "  esp_name      – the name of the Energy Service Provider (ESP / ESCO / contractor)\n"
                    "  facility_name – the name of the facility or project site\n"
                    "Use an empty string for any field you cannot find.\n\n"
                    f"{text}"
                ),
            }],
        )
        raw = response.content[0].text.strip()
        return json.loads(_strip_fences(raw))
    except Exception as e:
        logging.warning("Metadata extraction failed: %s", e)
        return {}


class StreamlitLogger:
    def __init__(self, container):
        self.container = container
        self.lines = []

    def log(self, msg: str):
        self.lines.append(msg)
        self.container.text("\n".join(self.lines[-250:]))


# ============================================================
# UI
# ============================================================

st.set_page_config(page_title="ARK Energy | M&V Plan Review Tool", layout="wide")

# ---------------- CSS — identical to Engineering Document Reviewer ----------------
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
        top: 0; left: 0; right: 0;
        z-index: 9999;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
        padding: 12px 18px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.35);
    }
    .ark-nav-inner{
        width: 98vw; margin: 0 auto; border-radius: 14px; padding: 10px 14px;
        display: flex; align-items: center; justify-content: space-between;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
    }
    .ark-nav-left { display:flex; align-items:center; gap:14px; }
    .ark-nav-title { color:white !important; font-size:22px !important; font-weight:900; line-height:1.2; margin:0; }
    .ark-nav-right { display:flex; align-items:center; gap:10px; }
    .pill {
        border-radius:999px; padding:8px 14px; font-size:14px; font-weight:900;
        border:1px solid rgba(255,255,255,0.25); color:white !important; background:transparent; white-space:nowrap;
    }

    .ark-section { margin-top:10px; margin-bottom:6px; display:flex; align-items:baseline; gap:10px; }
    .ark-section-title { font-size:18px; font-weight:900; color:var(--ark-blue); margin:0; line-height:1; }
    .ark-section-rule { height:2px; background:rgba(13,96,121,0.25); width:100%; margin-top:8px; margin-bottom:14px; }

    [data-testid="stFileUploaderLabel"] { padding-top: 4px !important; }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #ebebeb !important;
        border: 1px solid #c8c8c8 !important;
        border-radius: 6px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #e2e2e2 !important;
        border-color: #b0b0b0 !important;
    }

    /* Text inputs — match grey style */
    [data-testid="stTextInput"] input {
        background-color: #e8e8e8 !important;
        border: 1px solid #c8c8c8 !important;
        border-radius: 6px !important;
        color: #111 !important;
    }
    [data-testid="stTextInput"] input:focus {
        background-color: #e0e0e0 !important;
        border-color: #0D6079 !important;
        box-shadow: none !important;
    }

    label { font-size:15px !important; font-weight:700 !important; }

    div.stButton > button[kind="primary"],
    div.stButton > button[kind="primary"] * { color:#FFFFFF !important; }
    div.stButton > button[kind="primary"] {
        background-color: var(--ark-orange) !important;
        color: #FFFFFF !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
        border: none !important;
        box-shadow: 0 8px 20px rgba(247,148,40,0.25) !important;
        width: 100% !important;
    }
    div.stButton > button[kind="primary"]:hover { background-color:var(--ark-blue) !important; color:#FFFFFF !important; }

    /* ── File uploader: hide the stIconMaterial text node that causes "uploadupload" ── */
    [data-testid="stFileUploaderDropzone"] button [data-testid="stIconMaterial"] {
        display: none !important;
    }

    .stat-card { background:white; border-radius:12px; padding:18px 16px; text-align:center; box-shadow:0 3px 10px rgba(0,0,0,0.06); margin-bottom:8px; }
    .stat-number { font-size:36px; font-weight:900; line-height:1; margin-bottom:6px; }
    .stat-label  { font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; color:#555; }
    .color-blue   { color:#0D6079; }
    .color-green  { color:#375623; }
    .color-red    { color:#9C0006; }
    .color-orange { color:#9C6500; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------- JS: nuke stIconMaterial text nodes from the parent document ----------------
# The iframe trick: inject a script into the main page via st.markdown with a postMessage bridge
st.markdown("""
<script>
(function fixUploaders() {
    function hide() {
        // Walk every stFileUploaderDropzone button and hide stIconMaterial children
        document.querySelectorAll('[data-testid="stFileUploaderDropzone"] button').forEach(btn => {
            btn.querySelectorAll('[data-testid="stIconMaterial"]').forEach(el => {
                el.style.cssText = 'display:none!important;width:0!important;height:0!important;overflow:hidden!important;font-size:0!important;';
            });
            // Also nuke bare text nodes that are just "upload"
            btn.childNodes.forEach(node => {
                if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().toLowerCase() === 'upload') {
                    node.textContent = '';
                }
            });
        });
    }
    // Run immediately and on every DOM mutation
    hide();
    new MutationObserver(hide).observe(document.body, { childList: true, subtree: true });
})();
</script>
""", unsafe_allow_html=True)

# ---------------- Fixed Header ----------------
logo_path = str(LOGO_PATH)
logo_b64  = img_to_base64(logo_path) if Path(logo_path).exists() else ""

st.markdown(
    f"""
    <div class="ark-nav">
      <div class="ark-nav-inner">
        <div class="ark-nav-left">
          <img src="data:image/png;base64,{logo_b64}"
            style="height:68px; width:auto; display:block;" />
          <div><div class="ark-nav-title">M&amp;V Plan Review Tool</div></div>
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
    <div class="ark-section"><div class="ark-section-title">Upload documents</div></div>
    <div class="ark-section-rule"></div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)
with col1:
    main_pdf_upload = st.file_uploader(
        "M&V plan",
        type=["pdf"],
        accept_multiple_files=False,
        help="This is the document that will be reviewed.",
    )
with col2:
    supporting_uploads = st.file_uploader(
        "Supporting documents",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload supporting IGA document",
    )

# Also inject via component to reach inside Streamlit's shadow-dom-like iframes
components.html("""
<script>
(function fixUploadersFromComponent() {
    function hide(root) {
        root.querySelectorAll('[data-testid="stIconMaterial"]').forEach(el => {
            el.style.cssText = 'display:none!important;width:0!important;overflow:hidden!important;font-size:0!important;';
        });
        root.querySelectorAll('[data-testid="stFileUploaderDropzone"] button').forEach(btn => {
            btn.childNodes.forEach(node => {
                if (node.nodeType === 3 && node.textContent.trim().toLowerCase() === 'upload') {
                    node.textContent = '';
                }
            });
        });
    }
    function run() {
        // Target parent window
        try { hide(window.parent.document.body); } catch(e) {}
        try { hide(window.top.document.body); } catch(e) {}
    }
    run();
    try {
        new MutationObserver(run).observe(window.parent.document.body, { childList:true, subtree:true });
    } catch(e) {}
})();
</script>
""", height=0)

# ---------------- Auto-extract metadata from uploaded M&V Plan ----------------
if main_pdf_upload is not None:
    file_key = f"{main_pdf_upload.name}_{main_pdf_upload.size}"
    if st.session_state.get("_meta_file_key") != file_key:
        with st.spinner("Extracting submission details from document…"):
            pdf_data = main_pdf_upload.read()
            main_pdf_upload.seek(0)          # rewind so it can be read again later
            meta = _extract_submission_metadata(pdf_data)
        st.session_state["_meta_file_key"]    = file_key
        st.session_state["_meta_ref_no"]      = meta.get("ref_no", "")
        st.session_state["_meta_client_name"] = meta.get("client_name", "")
        st.session_state["_meta_esp_name"]    = meta.get("esp_name", "")
        st.session_state["_meta_facility"]    = meta.get("facility_name", "")
else:
    # Clear cached metadata when file is removed
    for k in ("_meta_file_key", "_meta_ref_no", "_meta_client_name", "_meta_esp_name", "_meta_facility"):
        st.session_state.pop(k, None)

# ---------------- Submission details ----------------
st.markdown(
    """
    <div class="ark-section"><div class="ark-section-title">Submission details</div></div>
    <div class="ark-section-rule"></div>
    """,
    unsafe_allow_html=True,
)

ref_no        = st.text_input("Ref. No.",       value=st.session_state.get("_meta_ref_no", ""))
client_name   = st.text_input("Client Name",   value=st.session_state.get("_meta_client_name", ""))
esp_name      = st.text_input("ESP's Name",    value=st.session_state.get("_meta_esp_name", ""))
facility_name = st.text_input("Facility Name", value=st.session_state.get("_meta_facility", ""))

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
        mv_bytes         = main_pdf_upload.read()
        supporting_bytes = [f.read() for f in (supporting_uploads or [])]
        main_name        = main_pdf_upload.name

        logger.log(f"📄 Reviewing: {main_name} with {len(supporting_bytes)} supporting document(s).")

        with st.spinner(f"Generating comments for {main_name}…"):
            result = run_mv_review(
                mv_bytes,
                supporting_bytes,
                ref_no,
                client_name,
                esp_name,
                main_name,
                facility_name=facility_name,
            )

        logger.log(f"✅ Done. {result['total']} questions reviewed.")

        st.markdown(
            """
            <div class="ark-section"><div class="ark-section-title">Results</div></div>
            <div class="ark-section-rule"></div>
            """,
            unsafe_allow_html=True,
        )
        st.success("Review completed successfully.")

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="stat-card"><div class="stat-number color-blue">{result["total"]}</div><div class="stat-label">Questions Reviewed</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-card"><div class="stat-number color-green">{result["approved"]}</div><div class="stat-label">Approved</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-card"><div class="stat-number color-red">{result["not_approved"]}</div><div class="stat-label">Not Approved</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="stat-card"><div class="stat-number color-orange">{result["incomplete"]}</div><div class="stat-label">Incomplete</div></div>', unsafe_allow_html=True)

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
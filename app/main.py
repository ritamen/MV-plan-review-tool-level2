"""
main.py
-------
FastAPI backend for the M&V Plan Review Tool.

Responsibilities:
- Serve the static frontend
- Receive uploaded PDFs (M&V Plan required, IGA optional)
- Load reviewer prompt and call Claude API (extended thinking)
- Parse + validate the JSON response; retry once on failure
- Write results into the Excel template (in memory)
- Return the filled Excel as a download + results summary as JSON
"""

import base64
import io
import json
import logging
import os
import re
import sys
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Ensure the app/ directory is on sys.path so sibling modules resolve correctly
sys.path.insert(0, str(Path(__file__).resolve().parent))

import anthropic
from dotenv import load_dotenv
from pypdf import PdfReader
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from excel_writer import write_review
from sn_extractor import extract_expected_sns

# ── Environment & paths ───────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR.parent / ".env")

PROMPT_PATH    = BASE_DIR / "assets" / "MV_Plan_Reviewer_Prompt.txt"
TEMPLATE_PATH  = BASE_DIR / "assets" / "M_V_Plan_Review_Sheet.xlsx"
STATIC_DIR     = BASE_DIR / "static"

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)

# ── Load assets at startup ────────────────────────────────────────────────────
try:
    REVIEWER_PROMPT: str = PROMPT_PATH.read_text(encoding="utf-8")
    logger.info("Reviewer prompt loaded (%d chars)", len(REVIEWER_PROMPT))
except FileNotFoundError:
    logger.critical("Reviewer prompt not found at %s", PROMPT_PATH)
    raise

try:
    TEMPLATE_BYTES: bytes = TEMPLATE_PATH.read_bytes()
    EXPECTED_SNS: List[str] = extract_expected_sns(str(TEMPLATE_PATH))
    logger.info("Excel template loaded; %d question SNs found", len(EXPECTED_SNS))
except FileNotFoundError:
    logger.critical("Excel template not found at %s", TEMPLATE_PATH)
    raise

# ── API constants ─────────────────────────────────────────────────────────────
MODEL          = "claude-opus-4-5"
MAX_TOKENS     = 8000
THINKING_TOKENS = 5000
TIMEOUT_SECS   = 180

# Character budget for IGA text extraction.
# Native PDFs cost ~1 500 tokens/page regardless of content density.
# Text extraction is far cheaper for text-heavy IGA documents.
# The IGA is ALWAYS sent as extracted text (never native PDF) so the budget
# is predictable.  At ~4 chars/token, 400 000 chars ≈ 100 000 tokens,
# leaving ~100 000 tokens for the system prompt, M&V Plan, and output.
MAX_IGA_TEXT_CHARS = 400_000

VALID_INCLUDED = {"Yes", "No", "Partial"}
VALID_STATUS   = {"Approved", "Not Approved", "Incomplete"}

# ── FastAPI app ───────────────────────────────────────────────────────────────
app = FastAPI(title="M&V Plan Review Tool")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


# ── Helpers ───────────────────────────────────────────────────────────────────

MAX_PDF_PAGES = 100  # Anthropic API hard limit (total across all PDF blocks)


def _encode_pdf(data: bytes) -> str:
    return base64.standard_b64encode(data).decode("utf-8")


def _pdf_page_count(pdf_bytes: bytes) -> int:
    return len(PdfReader(io.BytesIO(pdf_bytes)).pages)


def _pdf_to_text(pdf_bytes: bytes, label: str, max_chars: Optional[int] = None) -> str:
    """
    Extract text from every page of a PDF using pypdf.

    If max_chars is set, extraction stops once the accumulated text would
    exceed that limit, and a truncation notice is appended so the model is
    aware that some pages were not included.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    parts: list[str] = []
    chars_so_far = 0
    stopped_at: Optional[int] = None

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
            f"{total_pages} (≈{chars_so_far:,} chars extracted). "
            f"{omitted} page(s) were omitted to stay within the token limit.]"
        )
    return result


def _strip_fences(text: str) -> str:
    """Remove leading/trailing markdown code fences."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def _extract_json_text(response) -> str:
    """
    From a Claude API response, extract the text content block
    (skipping thinking blocks) and return its text.
    """
    for block in response.content:
        if block.type == "text":
            return block.text
    raise ValueError("No text content block found in API response.")


def _validate_items(items: list) -> List[str]:
    """
    Validate each item in the parsed JSON list.
    Returns a list of error strings (empty = valid).
    """
    errors = []
    for i, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"Item {i} is not a dict")
            continue
        for field in ("sn", "included", "status", "comment"):
            if field not in item:
                errors.append(f"Item {i} missing field '{field}'")
        if "included" in item and item["included"] not in VALID_INCLUDED:
            errors.append(
                f"Item {i} (sn={item.get('sn')}) has invalid included={item['included']!r}. "
                f"Must be one of {VALID_INCLUDED}"
            )
        if "status" in item and item["status"] not in VALID_STATUS:
            errors.append(
                f"Item {i} (sn={item.get('sn')}) has invalid status={item['status']!r}. "
                f"Must be one of {VALID_STATUS} (not 'Approved as Noted')"
            )
    return errors


def _parse_and_validate(raw: str) -> Tuple[list, List[str]]:
    """
    Strip fences, parse JSON, validate each item.
    Returns (items, errors).
    """
    cleaned = _strip_fences(raw)
    try:
        items = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        return [], [f"JSON parse error: {exc}"]
    if not isinstance(items, list):
        return [], ["Response is not a JSON array"]
    errors = _validate_items(items)
    return items, errors


def _add_pdf_to_content(
    content: list,
    pdf_bytes: bytes,
    label: str,
    max_chars: Optional[int] = None,
) -> None:
    """
    Add a PDF to the content list.

    - If the PDF is within the native-PDF page limit (≤ MAX_PDF_PAGES):
      send as a base64 document block so Claude can use its vision renderer.
    - Otherwise: extract text from *all* pages with pypdf and send as a text
      block, truncating at max_chars if provided to stay within the token limit.
    """
    pages = _pdf_page_count(pdf_bytes)
    if pages <= MAX_PDF_PAGES:
        content.append({
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": _encode_pdf(pdf_bytes),
            },
            "title": label,
        })
        logger.info("%s: %d pages — sent as native PDF", label, pages)
    else:
        extracted = _pdf_to_text(pdf_bytes, label, max_chars=max_chars)
        content.append({
            "type": "text",
            "text": (
                f"=== {label} ({pages} pages total — full text extracted) ===\n\n"
                f"{extracted}"
            ),
        })
        logger.info(
            "%s: %d pages — exceeds native-PDF limit, sent as extracted text "
            "(char budget: %s)",
            label, pages, f"{max_chars:,}" if max_chars else "unlimited",
        )


def _add_iga_to_content(content: list, iga_bytes: bytes) -> None:
    """
    Always send the IGA as extracted text, never as a native PDF.

    Native PDFs cost ~1 500 tokens/page regardless of text density; for a
    typical 60–200 page IGA that easily blows the 200 000-token limit when
    combined with the M&V Plan.  Text extraction costs only what the actual
    text contains, and the character budget keeps us safely under the limit.
    """
    pages = _pdf_page_count(iga_bytes)
    extracted = _pdf_to_text(iga_bytes, "IGA Document", max_chars=MAX_IGA_TEXT_CHARS)
    content.append({
        "type": "text",
        "text": (
            f"=== IGA Document ({pages} pages total — text extracted) ===\n\n"
            f"{extracted}"
        ),
    })
    logger.info(
        "IGA Document: %d pages — always sent as extracted text (char budget: %s)",
        pages, f"{MAX_IGA_TEXT_CHARS:,}",
    )


def _build_user_content(mv_bytes: bytes, iga_bytes: Optional[bytes]) -> list:
    content: list = []
    _add_pdf_to_content(content, mv_bytes, "M&V Plan")
    if iga_bytes:
        _add_iga_to_content(content, iga_bytes)
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


def _call_api(client: anthropic.Anthropic, user_content: list) -> str:
    """Call the Claude API and return the raw text content."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=1,  # required when extended thinking is enabled
        thinking={"type": "enabled", "budget_tokens": THINKING_TOKENS},
        system=REVIEWER_PROMPT,
        messages=[{"role": "user", "content": user_content}],
        timeout=TIMEOUT_SECS,
    )
    return _extract_json_text(response)


def _call_api_retry(client: anthropic.Anthropic, user_content: list) -> str:
    """
    Retry call: appends the previous assistant response and a correction
    instruction as a new user message, requesting a clean JSON-only output.
    """
    # First make the original call to get the bad response text
    first_raw = _call_api(client, user_content)

    retry_content = user_content + [
        {
            "type": "text",
            "text": first_raw,
        }
    ]
    # Build fresh messages for the retry with correction
    retry_messages = [
        {"role": "user", "content": user_content},
        {"role": "assistant", "content": first_raw},
        {
            "role": "user",
            "content": (
                "Your previous response was invalid. Return ONLY a JSON array. "
                "Each item: sn (string), included (Yes/No/Partial), "
                "status (Approved/Not Approved/Incomplete — not Approved as Noted), "
                "comment (string, empty if Approved). Nothing else."
            ),
        },
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


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/")
async def root():
    from fastapi.responses import FileResponse
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/review")
async def review(
    mv_plan:  UploadFile = File(...),
    iga:      UploadFile = File(None),
    ref_no:   str = "",
    esp_name: str = "",
):
    """
    Main review endpoint.
    Accepts M&V Plan PDF (required) and IGA PDF (optional).
    Returns JSON with summary + base64-encoded filled Excel.
    """
    # ── Validate uploads ──────────────────────────────────────────────────────
    if mv_plan.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="M&V Plan must be a PDF file.")

    mv_bytes = await mv_plan.read()
    if not mv_bytes:
        raise HTTPException(status_code=400, detail="M&V Plan PDF is empty.")

    iga_bytes: Optional[bytes] = None
    if iga and iga.filename:
        if iga.content_type != "application/pdf":
            raise HTTPException(status_code=400, detail="IGA document must be a PDF file.")
        iga_bytes = await iga.read()
        if not iga_bytes:
            iga_bytes = None

    # ── Build API content ─────────────────────────────────────────────────────
    user_content = _build_user_content(mv_bytes, iga_bytes)

    # ── Call Claude API ───────────────────────────────────────────────────────
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        logger.error("ANTHROPIC_API_KEY not set")
        raise HTTPException(status_code=500, detail="Server configuration error.")

    client = anthropic.Anthropic(api_key=api_key)

    try:
        logger.info("Calling Claude API (model=%s)...", MODEL)
        raw_text = _call_api(client, user_content)
        items, errors = _parse_and_validate(raw_text)

        if errors:
            logger.warning("Initial response invalid (%s). Retrying...", errors)
            raw_text = _call_api_retry(client, user_content)
            items, errors = _parse_and_validate(raw_text)
            if errors:
                logger.error("Retry response also invalid: %s", errors)
                raise HTTPException(
                    status_code=500,
                    detail=(
                        "The AI returned an invalid response after retry. "
                        "Please re-run the review."
                    ),
                )

    except anthropic.BadRequestError as exc:
        msg = str(exc)
        if "prompt is too long" in msg or "token" in msg.lower():
            logger.error("Prompt too long for Claude API: %s", msg)
            raise HTTPException(
                status_code=413,
                detail=(
                    "The combined documents are too large for the AI to process "
                    f"(exceeded the 200 000-token limit). "
                    "Try uploading a shorter IGA document, or split it into smaller sections."
                ),
            )
        logger.error("Claude API bad request: %s", msg)
        raise HTTPException(status_code=400, detail=f"API request error: {msg}")
    except anthropic.APITimeoutError:
        logger.error("Claude API call timed out after %ds", TIMEOUT_SECS)
        raise HTTPException(
            status_code=504,
            detail=(
                f"The review timed out after {TIMEOUT_SECS} seconds. "
                "Please try again — large documents may take longer."
            ),
        )
    except anthropic.AuthenticationError:
        logger.error("Claude API authentication error (key may be invalid)")
        raise HTTPException(status_code=500, detail="Server configuration error.")
    except HTTPException:
        raise
    except Exception:
        logger.error("Unexpected error during API call:\n%s", traceback.format_exc())
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred. Please try again.",
        )

    # ── Build review dict + coverage check ───────────────────────────────────
    review_by_sn = {str(item["sn"]).strip(): item for item in items}
    returned_sns = set(review_by_sn.keys())
    expected_sns = set(EXPECTED_SNS)
    missing_sns  = sorted(expected_sns - returned_sns)

    if missing_sns:
        logger.warning("Missing SNs not returned by AI: %s", missing_sns)

    # ── Tally results ─────────────────────────────────────────────────────────
    approved_count     = sum(1 for it in items if it.get("status") == "Approved")
    not_approved_count = sum(1 for it in items if it.get("status") == "Not Approved")
    incomplete_count   = sum(1 for it in items if it.get("status") == "Incomplete")
    total_count        = len(items)

    # ── Write Excel ───────────────────────────────────────────────────────────
    try:
        filled_bytes = write_review(TEMPLATE_BYTES, review_by_sn)
    except Exception:
        logger.error("Excel write error:\n%s", traceback.format_exc())
        raise HTTPException(
            status_code=500,
            detail="Failed to write results into the Excel template.",
        )

    # ── Return response ───────────────────────────────────────────────────────
    excel_b64 = base64.standard_b64encode(filled_bytes).decode("utf-8")

    # Build a descriptive output filename using ref_no / esp_name if provided
    base_name = mv_plan.filename.replace(".pdf", "")
    parts = ["MV_Plan_Review"]
    if ref_no:
        parts.append(ref_no.replace(" ", "_"))
    if esp_name:
        parts.append(esp_name.replace(" ", "_"))
    if not (ref_no or esp_name):
        parts.append(base_name)
    output_filename = "_".join(parts) + ".xlsx"

    return JSONResponse(
        content={
            "summary": {
                "total":        total_count,
                "approved":     approved_count,
                "not_approved": not_approved_count,
                "incomplete":   incomplete_count,
                "missing_sns":  missing_sns,
            },
            "excel_b64": excel_b64,
            "filename":  output_filename,
        }
    )

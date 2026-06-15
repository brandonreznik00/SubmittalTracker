import anthropic
import base64
import json
import os
import tempfile
import time

import fitz  # PyMuPDF

_client = None

def client():
    global _client
    if _client is None:
        # hard 2-minute cap per request — a wedged HTTP connection should
        # fail fast and retry, not hang the watcher for the SDK's 10-min default
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"],
                                      timeout=120.0, max_retries=1)
    return _client

MODEL = "claude-opus-4-8"
MAX_PAGES = 8          # metadata lives on the transmittal/cover pages
MAX_BYTES = 25_000_000
# Stay under the org's input-tokens-per-minute rate limit (10k on the
# starter tier — raise via env var if the account tier grows).
TPM_BUDGET = int(os.environ.get("SUBMITTAL_TPM_BUDGET", "9000"))

_window_start = 0.0
_window_tokens = 0


def _pace(tokens: int):
    """Block until sending `tokens` won't exceed the per-minute budget."""
    global _window_start, _window_tokens
    now = time.time()
    if now - _window_start >= 60:
        _window_start, _window_tokens = now, 0
    if _window_tokens + tokens > TPM_BUDGET:
        wait = 60 - (now - _window_start)
        if wait > 0:
            print(f"    (pacing {wait:.0f}s to stay under the API rate limit)",
                  flush=True)
            time.sleep(wait)
        _window_start, _window_tokens = time.time(), 0
    _window_tokens += tokens

PROMPT = """This is a construction submittal PDF (transmittal, shop drawing, product data, sample, or similar).
Extract the following fields and return JSON only, no other text. Use null when a field is not present.

- project (string — the project/building name or address this submittal belongs to, e.g. "123 Main Street". IMPORTANT: if the filename names a project/address and the document shows a different one, USE THE FILENAME's project — PMs name files by project; documents often carry a consultant's alternate lot address.)
- trade (string — the construction trade this submittal belongs to, e.g. "Superstructure", "Concrete", "Masonry", "Structural Steel", "Windows", "Roofing", "HVAC", "Plumbing", "Electrical", "Elevator", "Sprinkler", "Carpentry", "Scaffolding", "Lighting")
- spec_section (string — CSI spec section number and name if shown, e.g. "316300 - Auger Cast-In-Place Piles")
- submittal_number (string — the submittal/transmittal number, e.g. "316300-12" or "38")
- package (string or null — submittal package name/number)
- title (string — short title of the submittal)
- type (one of: "Shop Drawing", "Product Information", "Document", "Plans", "Sample", "Other")
- description (string or null — one-line description)
- revision (string or number — revision; use 0 if not shown)
- responsible_contractor (string or null — the subcontractor/vendor company responsible)
- received_from (string or null — person name it was received from / submitted by)
- location (string or null — building location/area the submittal applies to)
- submittal_date (MM/DD/YYYY or null — the date on the submittal/transmittal)
- required_on_site_date (MM/DD/YYYY or null)
- lead_time (string or null — lead time if stated, e.g. "6 weeks")
- approvers (string or null — engineer/architect named as reviewer/approver, "Name (Company)" format)
- approval_status (one of "Approved", "Approved as Noted", "Revise & Resubmit", "Rejected", or null — ONLY from an actual reviewer disposition stamped/marked on the document, e.g. "APPROVED", "NO EXCEPTIONS TAKEN", "APPROVED AS NOTED", "MAKE CORRECTIONS NOTED" (= Approved as Noted), "REVISE AND RESUBMIT", "REJECTED". A filename marker like "Reviewed", "Approved" or "R&R" also counts when the stamp page isn't included. Use null when the submittal has not been reviewed yet — never guess.)
- approval_date (MM/DD/YYYY or null — the date on or next to that reviewer stamp)
- approval_discipline (one of "Structural", "MEP", "Interior", "Consultant", "Architect", or null — which kind of reviewer stamped it, judged from the stamping firm/engineer's role)

If multiple distinct submittals appear in one PDF, extract the PRIMARY one described by the transmittal page."""


_COVER_MARKERS = (
    "transmittal", "we are sending you", "submittal cover",
    "submittal review", "submitted for", "submittal package",
    "shop drawing review", "submitted by", "returned for",
)

def _has_cover_page(pdf_path: str) -> bool:
    """True when page 1 reads like a transmittal/cover sheet — those carry
    all the metadata (and usually the reviewer disposition), so sending just
    that page costs ~1/8 of a full read."""
    try:
        doc = fitz.open(pdf_path)
        text = doc[0].get_text("text").lower()
        doc.close()
    except Exception:
        return False
    # Drawings have sparse title-block text; a real cover sheet is text-rich.
    if len(text.split()) < 40:
        return False
    return any(m in text for m in _COVER_MARKERS)


def _shrink_pdf(pdf_path: str, max_pages: int = MAX_PAGES) -> str:
    """Return a path to a PDF small enough to send: first max_pages pages."""
    doc = fitz.open(pdf_path)
    if len(doc) <= max_pages and os.path.getsize(pdf_path) <= MAX_BYTES:
        doc.close()
        return pdf_path
    out = fitz.open()
    out.insert_pdf(doc, from_page=0, to_page=min(max_pages, len(doc)) - 1)
    tmp = os.path.join(tempfile.gettempdir(), "_submittal_head.pdf")
    out.save(tmp, garbage=3, deflate=True)
    out.close(); doc.close()
    return tmp


def _page_images(pdf_path: str, max_pages: int) -> list:
    """Rasterize the first pages as PNG image blocks — fallback for PDFs the
    API can't parse (Bluebeam Stapler exports trip 'Could not process PDF')."""
    doc = fitz.open(pdf_path)
    blocks = []
    for i in range(min(max_pages, len(doc))):
        page = doc[i]
        zoom = 1568 / max(page.rect.width, page.rect.height)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        blocks.append({"type": "image",
                       "source": {"type": "base64", "media_type": "image/png",
                                  "data": base64.standard_b64encode(
                                      pix.tobytes("png")).decode("utf-8")}})
    doc.close()
    return blocks


def extract_submittal(pdf_path: str) -> dict:
    fname = os.path.basename(pdf_path)
    prompt = (f"Filename of this PDF: \"{fname}\" (the project name and submittal info may "
              f"appear only in the filename — use it when the document itself doesn't say).\n\n" + PROMPT)

    # Shrink to fit the per-minute token budget using an OFFLINE estimate.
    # (count_tokens calls consume the same input-tokens-per-minute allowance
    # as real requests, so pre-counting was starving the actual extraction.)
    # Metadata is on the cover pages, so dropping trailing pages is safe.
    EST_PER_PAGE, EST_OVERHEAD = 2800, 1000
    pages = MAX_PAGES
    if _has_cover_page(pdf_path):
        pages = 1
        print("    (transmittal/cover page detected — sending page 1 only)",
              flush=True)
    while pages > 2 and EST_OVERHEAD + pages * EST_PER_PAGE > TPM_BUDGET:
        pages = max(2, pages // 2)
    if 1 < pages < MAX_PAGES:
        print(f"    (sending first {pages} pages to fit the API rate limit)",
              flush=True)
    send_path = _shrink_pdf(pdf_path, pages)
    with open(send_path, "rb") as f:
        pdf_data = base64.standard_b64encode(f.read()).decode("utf-8")

    tokens = EST_OVERHEAD + pages * EST_PER_PAGE
    _pace(tokens)
    try:
        response = client().messages.create(
            model=MODEL,
            max_tokens=1500,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "document",
                     "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_data}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )
    except anthropic.BadRequestError as e:
        if "could not process pdf" not in str(e).lower():
            raise
        # The API's PDF parser rejected the file (deterministic 400) even
        # though it opens fine locally. Send the same pages as images instead.
        print("    (API can't parse this PDF — retrying as page images)",
              flush=True)
        _pace(tokens)
        response = client().messages.create(
            model=MODEL,
            max_tokens=1500,
            messages=[{
                "role": "user",
                "content": _page_images(send_path, pages)
                           + [{"type": "text", "text": prompt}],
            }],
        )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

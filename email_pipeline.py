"""
email_pipeline.py
-----------------
Complete 6-step email pipeline. Run with:
    python email_pipeline.py

Flow:
  1. Poll Gmail inbox (IMAP) for unseen emails
  2. Download PDF / DOCX / DOC / TXT attachments
  3. Extract plain text (auto-detects Krutidev → Unicode; else English as-is)
     → save as .txt next to attachment
  4. Send .txt to Gemini 2.5 Flash → get all table structures (printed to terminal)
  5. POST email info + extracted text + table details to OpenClaw → get response
  6. Build reply PDF (reportlab) → send SMTP reply with PDF attached

No images / vision API used anywhere — only plain text extraction.
"""

import datetime
import imaplib
import json
import os
import smtplib
import sys
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pyzmail
import requests
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Local modules ──────────────────────────────────────────────────────────────
from extract_text import extract_text as extract_docx_text   # for DOCX/DOC
from krutidev_converter import krutidev_to_unicode, unicode_to_krutidev  # for heuristic check / reverse

# ── Config ─────────────────────────────────────────────────────────────────────
from config import (
    IMAP_HOST, EMAIL, PASSWORD,
    SMTP_SERVER, SMTP_PORT, SMTP_EMAIL, SMTP_PASSWORD,
    OPENCLAW_URL, TOKEN,
    GEMINI_API_KEY, GEMINI_MODEL,
    ATTACHMENT_DIR, PROCESSED_LOG,
    CHECK_INTERVAL,
)

# ── Gemini SDK ─────────────────────────────────────────────────────────────────
try:
    from google import genai as google_genai
except ImportError:
    print("ERROR: google-genai not installed. Run: pip install google-genai")
    sys.exit(1)

# ── pypdf for PDF text extraction ──────────────────────────────────────────────
try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader      # fallback
    except ImportError:
        print("ERROR: pypdf not installed. Run: pip install pypdf")
        sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# Logging helper
# ─────────────────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# ─────────────────────────────────────────────────────────────────────────────
# Processed-file tracking
# ─────────────────────────────────────────────────────────────────────────────

def load_processed() -> set:
    if not os.path.exists(PROCESSED_LOG):
        return set()
    with open(PROCESSED_LOG, "r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def mark_processed(filename: str) -> None:
    with open(PROCESSED_LOG, "a", encoding="utf-8") as f:
        f.write(filename + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — Download attachments
# ─────────────────────────────────────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".doc", ".txt"}


def save_attachment(filename: str, payload: bytes) -> str:
    """Save bytes to ATTACHMENT_DIR — returns saved path."""
    os.makedirs(ATTACHMENT_DIR, exist_ok=True)
    base, ext = os.path.splitext(filename)
    dest = os.path.join(ATTACHMENT_DIR, filename)
    counter = 1
    while os.path.exists(dest):
        dest = os.path.join(ATTACHMENT_DIR, f"{base}_{counter}{ext}")
        counter += 1
    with open(dest, "wb") as f:
        f.write(payload)
    log(f"  [STEP 2] Saved: {dest} ({len(payload):,} bytes)")
    return dest


def extract_attachments(msg: pyzmail.PyzMessage) -> list[str]:
    """Return saved paths for all supported attachments."""
    paths = []
    for part in msg.mailparts:
        filename = part.filename
        if not filename:
            continue
        ext = os.path.splitext(filename)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            continue
        payload = part.get_payload()
        if not isinstance(payload, bytes):
            continue
        path = save_attachment(filename, payload)
        paths.append(path)
    return paths


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — Text extraction (no images)
# ─────────────────────────────────────────────────────────────────────────────

# ── Language detection: Devanagari Unicode range U+0900–U+097F ───────────────
# Krutidev is a font that stores Hindi as *plain ASCII*.
# The only reliable way to detect it is:
#   1. Extract raw text (no conversion).
#   2. Try Krutidev → Unicode conversion.
#   3. If the converted result contains Devanagari chars → it was Krutidev.
#   4. If not → it was English (conversion is a no-op on real English text).

def _contains_devanagari(text: str) -> bool:
    """Return True if *text* contains any Devanagari Unicode character."""
    return any('\u0900' <= c <= '\u097f' for c in text)


def _extract_pdf_text(pdf_path: str) -> str:
    """Extract plain text from PDF using pypdf (no images)."""
    reader = PdfReader(pdf_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _detect_and_convert_text(raw_text: str, file_path: str | None = None,
                              is_docx: bool = False) -> tuple[str, str]:
    """
    Given *raw_text* (extracted without Krutidev conversion):
    - If it already contains Devanagari Unicode → return as-is ("Unicode Hindi").
    - Try Krutidev conversion; if result has Devanagari → return converted ("Krutidev→Unicode").
    - Otherwise → return raw ("English").

    For DOCX files pass *file_path* + *is_docx=True* so we re-extract via
    extract_docx_text(convert_krutidev=True) which handles cell-level conversion.
    """
    # Already proper Unicode Hindi?
    if _contains_devanagari(raw_text):
        log("  [STEP 3] Already Unicode Hindi — no conversion needed.")
        return raw_text, "Unicode Hindi"

    # Try Krutidev conversion
    if is_docx and file_path:
        converted = extract_docx_text(file_path, convert_krutidev=True)
    else:
        converted = krutidev_to_unicode(raw_text)

    if _contains_devanagari(converted):
        log("  [STEP 3] Krutidev detected → converted to Unicode.")
        return converted, "Krutidev→Unicode"

    log("  [STEP 3] English text detected — no conversion needed.")
    return raw_text, "English"


def extract_and_save_text(file_path: str) -> tuple[str, str, str]:
    """
    Extract text from *file_path*, detect language, convert if needed.

    Returns
    -------
    txt_path : str   – path of the saved .txt file
    text     : str   – extracted (and possibly converted) text
    lang     : str   – detected language ("English", "Unicode Hindi", "Krutidev→Unicode")
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in (".docx", ".doc"):
        log(f"  [STEP 3] Extracting DOCX/DOC: {os.path.basename(file_path)}")
        raw_text = extract_docx_text(file_path, convert_krutidev=False)
        text, lang = _detect_and_convert_text(raw_text, file_path, is_docx=True)

    elif ext == ".pdf":
        log(f"  [STEP 3] Extracting PDF: {os.path.basename(file_path)}")
        raw_text = _extract_pdf_text(file_path)
        text, lang = _detect_and_convert_text(raw_text)

    elif ext == ".txt":
        log(f"  [STEP 3] Reading TXT: {os.path.basename(file_path)}")
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            raw_text = f.read()
        text, lang = _detect_and_convert_text(raw_text)

    else:
        log(f"  [STEP 3] Unsupported extension '{ext}' — skipping.")
        return "", "", "English"

    log(f"  [STEP 3] Language: {lang} | Characters extracted: {len(text):,}")

    # Save as .txt beside the attachment
    base = os.path.splitext(file_path)[0]
    txt_path = base + ".txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)
    log(f"  [STEP 3] Saved TXT → {txt_path}")

    return txt_path, text, lang


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — Gemini: find all table structures in extracted text
# ─────────────────────────────────────────────────────────────────────────────

_GEMINI_TABLE_PROMPT = """\
You are a document analysis assistant.

Below is the plain-text content extracted from a document.
Tables in DOCX files appear with cells separated by ' | ' and rows on separate lines.

Your task:
  1. Identify EVERY table present in the text.
  2. For each table:
     - Number it sequentially (Table 1, Table 2, …).
     - Include any nearby title or heading.
     - Show column headers.
     - Show ALL data rows as a clean Markdown table.
  3. After the human-readable analysis, output a JSON block in this exact format:

```json
{{
  "columns_by_table": {{
    "table_1": ["Col A", "Col B", …],
    "table_2": ["Col X", "Col Y", …]
  }}
}}
```

  4. Do NOT invent data. Only describe what is actually in the text.

--- DOCUMENT TEXT START ---
{document_text}
--- DOCUMENT TEXT END ---

Now provide the analysis followed by the JSON block.
"""


def get_columns_from_gemini(text: str) -> dict[str, list[str]]:
    """
    Send plain text to Gemini.
    - Prints the full human-readable table analysis to terminal.
    - Returns columns_by_page dict: {"table_1": ["Col A", ...], ...}
    """
    client = google_genai.Client(api_key=GEMINI_API_KEY)
    prompt = _GEMINI_TABLE_PROMPT.format(document_text=text)

    log(f"  [STEP 4] Sending text to Gemini ({GEMINI_MODEL}) …")
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
    )
    full_response = response.text

    # Print the full analysis to terminal
    print("\n" + "=" * 70)
    print("  GEMINI TABLE ANALYSIS:")
    print("=" * 70)
    print(full_response)
    print("=" * 70 + "\n")

    # Extract structured JSON from the response
    columns_by_page: dict[str, list[str]] = {}
    try:
        start = full_response.find("```json")
        end   = full_response.rfind("```")
        if start != -1 and end != -1 and end > start:
            json_str = full_response[start + 7 : end].strip()
            parsed = json.loads(json_str)
            columns_by_page = parsed.get("columns_by_table", {})
            log(f"  [STEP 4] Columns extracted: {columns_by_page}")
        else:
            log("  [STEP 4] No JSON block found in Gemini response.")
    except Exception as e:
        log(f"  [STEP 4] Could not parse Gemini JSON: {e}")

    return columns_by_page


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — OpenClaw: enrich with email context + table details
# ─────────────────────────────────────────────────────────────────────────────

def call_openclaw(subject: str, sender: str, body: str,
                  columns_by_page: dict[str, list[str]],
                  language: str = "English") -> dict | None:
    """
    Sends email context + detected column headers to OpenClaw using the
    email-context-responder skill.

    The skill returns JSON in the format:
    {
        "category": "...",
        "priority": "...",
        "requires_reply": true,
        "language_of_reply": "...",
        "matched_documents": [...],
        "data_used_in_reply": {
            "facts": [...],
            "source_files": [...]
        },
        "suggested_reply": "...",
        "reply_note": "..."
    }

    We additionally expect the skill to populate table_data so we can build
    a structured PDF. The skill is prompted to include it.
    """
    # Determine human-readable language label for the prompt (to help OpenClaw understand the input)
    is_hindi = language in ("Unicode Hindi", "Krutidev→Unicode")
    input_lang_label = "Hindi" if is_hindi else "English"

    prompt = f"""Use the email-context-responder skill.

IMPORTANT: The email and its attachments are written in {input_lang_label}.
However, your entire response — including the 'suggested_reply' field and all table data —
MUST be written in English only. Do NOT use any Hindi or Devanagari script.

AVAILABLE TABLES & COLUMNS (extracted from attachments):
{json.dumps(columns_by_page, indent=2, ensure_ascii=False)}

Your response MUST include a 'tables' field which is a LIST of objects.
Each object in the 'tables' list should have:
  - 'title': A short descriptive title for the table.
  - 'headers': The list of column headers.
  - 'rows': A list of lists representing the data rows.

EMAIL:
Subject : {json.dumps(subject)}
From    : {json.dumps(sender)}
Body    :
{body}
"""

    payload = {
        "model": "openclaw",
        "messages": [{"role": "user", "content": prompt}],
    }
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json",
        "x-openclaw-agent-id": "main",
    }

    log("  [STEP 5] Calling OpenClaw …")
    try:
        resp = requests.post(OPENCLAW_URL, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]
        log("--- OpenClaw response (first 25000 chars) ---")
        log(raw[:25000])
        log("-------------------------------------------")
        return _parse_json(raw)
    except Exception as e:
        log(f"  [STEP 5] OpenClaw API error: {e}")
        return None


def _parse_json(raw: str) -> dict | None:
    try:
        cleaned = raw.strip().replace("```json", "").replace("```", "").strip()
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start == -1 or end == -1:
            raise ValueError("No JSON object found.")
        return json.loads(cleaned[start : end + 1])
    except Exception as e:
        log(f"  [STEP 5] JSON parse error: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — Build reply PDF
# ─────────────────────────────────────────────────────────────────────────────

# ── Krutidev font registration helper ─────────────────────────────────────────

_KRUTIDEV_FONT_REGISTERED = False
_KRUTIDEV_FONT_NAME       = "KrutiDev010"

_KRUTIDEV_SEARCH_PATHS = [
    r"C:\Windows\Fonts\KRDEV010.TTF",
    r"C:\Windows\Fonts\KrutiDev010.ttf",
    r"C:\Windows\Fonts\Kruti Dev 010.ttf",
    os.path.join(os.path.dirname(__file__), "KRDEV010.TTF"),
    os.path.join(os.path.dirname(__file__), "KrutiDev010.ttf"),
    os.path.join(os.path.dirname(__file__), "Kruti Dev 010.ttf"),
]


def _ensure_krutidev_font() -> tuple[str, str]:
    """
    Register Kruti Dev 010 with ReportLab (once).
    Returns (regular_font_name, bold_font_name).
    Falls back to NotoSansDevanagari (variable font) if Kruti Dev is absent.
    """
    global _KRUTIDEV_FONT_REGISTERED
    if _KRUTIDEV_FONT_REGISTERED:
        return _KRUTIDEV_FONT_NAME, _KRUTIDEV_FONT_NAME

    path = next((p for p in _KRUTIDEV_SEARCH_PATHS if os.path.exists(p)), None)
    if not path:
        log("  [STEP 6] WARNING: Kruti Dev 010 font not found. "
            "Falling back to NotoSansDevanagari for Krutidev PDF. "
            "Install KRDEV010.TTF / KrutiDev010.ttf in C:\\Windows\\Fonts\\ "
            "or place it next to email_pipeline.py.")
        # Fall back to NotoSansDevanagari so at least Unicode Hindi renders
        return _ensure_hindi_font()

    try:
        pdfmetrics.registerFont(TTFont(_KRUTIDEV_FONT_NAME, path))
        _KRUTIDEV_FONT_REGISTERED = True
        log(f"  [STEP 6] Kruti Dev font registered: {path}")
        return _KRUTIDEV_FONT_NAME, _KRUTIDEV_FONT_NAME
    except Exception as e:
        log(f"  [STEP 6] Kruti Dev font registration failed: {e} — falling back.")
        return _ensure_hindi_font()


# ── Hindi font registration helper ────────────────────────────────────────────

_HINDI_FONT_REGISTERED = False
_HINDI_FONT_NAME = "NotoSansDevanagari"
_HINDI_FONT_BOLD_NAME = "NotoSansDevanagari-Bold"

# Common install paths for Noto Sans Devanagari on Windows
_NOTO_SEARCH_PATHS = [
    # Variable font (placed next to email_pipeline.py)
    os.path.join(os.path.dirname(__file__), "NotoSansDevanagari-VariableFont_wdth,wght.ttf"),
    # Standard static font names (Windows Fonts or project folder)
    r"C:\Windows\Fonts\NotoSansDevanagari-Regular.ttf",
    r"C:\Windows\Fonts\NotoSansDevanagari.ttf",
    os.path.join(os.path.dirname(__file__), "NotoSansDevanagari-Regular.ttf"),
    os.path.join(os.path.dirname(__file__), "NotoSansDevanagari.ttf"),
]
_NOTO_BOLD_SEARCH_PATHS = [
    r"C:\Windows\Fonts\NotoSansDevanagari-Bold.ttf",
    os.path.join(os.path.dirname(__file__), "NotoSansDevanagari-Bold.ttf"),
    # Fallback: reuse the variable font as bold
    os.path.join(os.path.dirname(__file__), "NotoSansDevanagari-VariableFont_wdth,wght.ttf"),
]


def _ensure_hindi_font() -> tuple[str, str]:
    """
    Register Noto Sans Devanagari with ReportLab (once).
    Returns (regular_font_name, bold_font_name).
    Falls back to Helvetica if the font file is not found.
    """
    global _HINDI_FONT_REGISTERED
    if _HINDI_FONT_REGISTERED:
        return _HINDI_FONT_NAME, _HINDI_FONT_BOLD_NAME

    reg_path  = next((p for p in _NOTO_SEARCH_PATHS  if os.path.exists(p)), None)
    bold_path = next((p for p in _NOTO_BOLD_SEARCH_PATHS if os.path.exists(p)), None)

    if not reg_path:
        log("  [STEP 6] WARNING: Noto Sans Devanagari font not found. "
            "Hindi text may not render correctly. "
            "Place NotoSansDevanagari-Regular.ttf next to email_pipeline.py "
            "or install it in C:\\Windows\\Fonts\\.")
        return "Helvetica", "Helvetica-Bold"

    try:
        pdfmetrics.registerFont(TTFont(_HINDI_FONT_NAME, reg_path))
        if bold_path:
            pdfmetrics.registerFont(TTFont(_HINDI_FONT_BOLD_NAME, bold_path))
        else:
            # Re-register regular as bold fallback
            pdfmetrics.registerFont(TTFont(_HINDI_FONT_BOLD_NAME, reg_path))
        _HINDI_FONT_REGISTERED = True
        log(f"  [STEP 6] Hindi font registered: {reg_path}")
        return _HINDI_FONT_NAME, _HINDI_FONT_BOLD_NAME
    except Exception as e:
        log(f"  [STEP 6] Font registration failed: {e} — falling back to Helvetica.")
        return "Helvetica", "Helvetica-Bold"


def create_reply_pdf(tables: list[dict], subject: str, facts: list,
                     output_path: str, language: str = "English") -> str:
    """
    Create a styled PDF with multiple tables + facts in the detected language.
    - "Krutidev→Unicode" : convert text back to Krutidev ASCII, use Kruti Dev font
    - "Unicode Hindi"    : use NotoSansDevanagari font
    - "English"          : use Helvetica
    Returns output_path.
    """
    is_krutidev = language == "Krutidev→Unicode"
    is_hindi    = language in ("Unicode Hindi", "Krutidev→Unicode")

    if is_krutidev:
        body_font, bold_font = _ensure_krutidev_font()
        log(f"  [STEP 6] PDF language: Krutidev — using font '{body_font}'")
    elif is_hindi:
        body_font, bold_font = _ensure_hindi_font()
        log(f"  [STEP 6] PDF language: Unicode Hindi — using font '{body_font}'")
    else:
        body_font, bold_font = "Helvetica", "Helvetica-Bold"
        log("  [STEP 6] PDF language: English — using Helvetica")

    page_size = landscape(A4) # default to landscape for safety with multiple tables
    doc = SimpleDocTemplate(
        output_path,
        pagesize=page_size,
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=2 * cm,    bottomMargin=2 * cm,
    )

    styles = getSampleStyleSheet()
    story  = []

    # ── Title ──────────────────────────────────────────────────────────────
    title_style = ParagraphStyle(
        "Title2", parent=styles["Title"], fontSize=14, spaceAfter=10,
        fontName=bold_font
    )
    story.append(Paragraph(f"Response: {subject}", title_style))
    story.append(Paragraph(
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        ParagraphStyle("GenDate", parent=styles["Normal"], fontName=body_font)
    ))
    story.append(Spacer(1, 0.5 * cm))

    # ── Styles for table cells ─────────────────────────────────────────────
    cell_style = ParagraphStyle(
        "Cell", parent=styles["Normal"], fontSize=8, leading=10,
        fontName=body_font
    )
    hdr_style = ParagraphStyle(
        "Hdr", parent=styles["Normal"], fontSize=8, leading=10,
        textColor=colors.white, fontName=bold_font
    )
    tbl_title_style = ParagraphStyle(
        "TblTitle", parent=styles["Heading3"], fontSize=10, spaceAfter=5,
        fontName=bold_font
    )

    if tables:
        for i, table_data in enumerate(tables, start=1):
            title = table_data.get("title", f"Table {i}")
            headers = table_data.get("headers", [])
            rows = table_data.get("rows", [])

            if not headers:
                continue

            story.append(Paragraph(title, tbl_title_style))
            
            table_body = [[Paragraph(str(h), hdr_style) for h in headers]]
            for row in rows:
                padded = (list(row) + [""] * len(headers))[: len(headers)]
                table_body.append([Paragraph(str(c), cell_style) for c in padded])

            usable_w  = page_size[0] - 3 * cm
            col_w     = usable_w / len(headers)
            tbl = Table(table_body, colWidths=[col_w] * len(headers), repeatRows=1)
            tbl.setStyle(TableStyle([
                ("BACKGROUND",     (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
                ("TEXTCOLOR",      (0, 0), (-1, 0), colors.white),
                ("FONTNAME",       (0, 0), (-1, 0), bold_font),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                    [colors.white, colors.HexColor("#F2F2F2")]),
                ("GRID",           (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
                ("VALIGN",         (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING",     (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
                ("LEFTPADDING",    (0, 0), (-1, -1), 4),
                ("RIGHTPADDING",   (0, 0), (-1, -1), 4),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 0.5 * cm))
    else:
        story.append(Paragraph(
            "No structured table data was returned by OpenClaw.",
            ParagraphStyle("Fallback", parent=styles["Normal"], fontName=body_font)
        ))

    # ── Source facts ───────────────────────────────────────────────────────
    if facts:
        story.append(Spacer(1, 0.8 * cm))
        story.append(Paragraph(
            "Source Facts" if not is_hindi else "स्रोत तथ्य",
            ParagraphStyle("H2", parent=styles["Heading2"], fontName=bold_font)
        ))
        story.append(Spacer(1, 0.2 * cm))
        fact_style = ParagraphStyle(
            "Fact", parent=styles["Normal"], fontSize=8, leading=12,
            leftIndent=10, fontName=body_font
        )
        for i, fact in enumerate(facts, start=1):
            story.append(Paragraph(f"{i}. {fact}", fact_style))

    doc.build(story)
    log(f"  [STEP 6] Reply PDF created: {output_path}")
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — Send reply via SMTP
# ─────────────────────────────────────────────────────────────────────────────

def send_reply(to_email: str, subject: str, body: str,
               pdf_path: str | None = None) -> None:
    """Send email reply, optionally with PDF attachment."""
    msg = MIMEMultipart()
    msg["From"]    = SMTP_EMAIL
    msg["To"]      = to_email
    msg["Subject"] = f"Re: {subject}"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(pdf_path)}"',
        )
        msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        log(f"  [STEP 6] Reply sent to {to_email}" +
            (f" with PDF: {os.path.basename(pdf_path)}" if pdf_path else ""))
    except Exception as e:
        log(f"  [STEP 6] SMTP error: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# Main email handler — orchestrates all 6 steps for one email
# ─────────────────────────────────────────────────────────────────────────────

def handle_email(subject: str, sender: str, body: str,
                 attachment_paths: list[str]) -> None:
    log("━" * 60)
    log(f"  From    : {sender}")
    log(f"  Subject : {subject}")
    log(f"  Files   : {[os.path.basename(p) for p in attachment_paths]}")
    log("━" * 60)

    processed = load_processed()
    all_extracted_text   = []
    all_columns_by_page: dict[str, list[str]] = {}
    detected_language    = "English"   # updated as files are processed

    for file_path in attachment_paths:
        filename = os.path.basename(file_path)

        # ── STEP 3: Extract text ───────────────────────────────────────────
        if filename in processed:
            log(f"  [STEP 3] '{filename}' already processed — skipping extraction.")
            continue

        txt_path, text, lang = extract_and_save_text(file_path)
        if not text.strip():
            log(f"  [STEP 3] No text extracted from '{filename}' — skipping.")
            continue

        all_extracted_text.append(f"=== {filename} ===\n{text}")
        mark_processed(filename)

        # Track the dominant language (Hindi takes priority over English)
        if lang in ("Unicode Hindi", "Krutidev→Unicode"):
            detected_language = lang
        elif detected_language == "English":
            detected_language = lang

        # ── STEP 4: Gemini — extract column headers (+ print table analysis) ──
        columns_by_page = get_columns_from_gemini(text)
        all_columns_by_page.update({
            f"{filename}_{k}": v for k, v in columns_by_page.items()
        })

    if not all_extracted_text:
        log("  No text extracted from any attachment — skipping reply.")
        return

    log(f"  [PIPELINE] Detected language for reply: {detected_language}")

    # ── STEP 5: OpenClaw ─────────────────────────────────────────────────────
    result = call_openclaw(subject, sender, body, all_columns_by_page,
                           language=detected_language)

    if not result:
        log("  [STEP 5] OpenClaw returned no result — sending text-only fallback reply.")
        send_reply(sender, subject,
                   "Thank you for your email. We have received your document.\n\nRegards,\nAuto-Reply System")
        return

    log(f"  [STEP 5] Category : {result.get('category', 'unknown')}")
    log(f"  [STEP 5] Priority : {result.get('priority', 'unknown')}")
    log(f"  [STEP 5] Requires reply: {result.get('requires_reply', False)}")

    if not result.get("requires_reply", False):
        log("  [STEP 5] No reply required.")
        return

    suggested_reply = result.get("suggested_reply", "").strip() or (
        "Please find the extracted data attached.\n\nRegards,\nAuto-Reply System"
    )
    
    # Extract multiple tables if present, otherwise fallback to single table_data
    tables = result.get("tables", [])
    if not tables and result.get("table_data"):
        tables = [result["table_data"]]

    facts = result.get("data_used_in_reply", {}).get("facts", []) or \
            result.get("facts", [])

    # ── STEP 6: Create PDF (Always English) ───────────────────────────────
    os.makedirs(ATTACHMENT_DIR, exist_ok=True)
    ts      = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_out = os.path.join(ATTACHMENT_DIR, f"reply_{ts}.pdf")

    pdf_path = None
    if tables:
        try:
            # We now force 'English' for the PDF to ensure English fonts and labels
            pdf_path = create_reply_pdf(tables, subject, facts, pdf_out,
                                        language="English")
        except Exception as e:
            log(f"  [STEP 6] PDF creation failed: {e}")
    else:
        log("  [STEP 6] No table data from OpenClaw — sending text-only reply.")

    # ── STEP 6: Send reply ────────────────────────────────────────────────
    send_reply(sender, subject, suggested_reply, pdf_path)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — IMAP polling loop
# ─────────────────────────────────────────────────────────────────────────────

def check_inbox() -> None:
    """Connect to Gmail, fetch unseen emails, process each one."""
    log("[STEP 1] Connecting to Gmail IMAP …")
    with imaplib.IMAP4_SSL(IMAP_HOST) as server:
        server.login(EMAIL, PASSWORD)
        server.select("INBOX")

        today = datetime.date.today().strftime("%d-%b-%Y")
        _, uids = server.search(None, f'(UNSEEN SINCE "{today}")')

        if not uids or not uids[0]:
            log("[STEP 1] No new emails.")
            return

        uid_list = uids[0].split()
        log(f"[STEP 1] Found {len(uid_list)} new email(s).")

        for uid in uid_list:
            _, raw_data = server.fetch(uid, "(RFC822)")
            if not raw_data or not raw_data[0]:
                continue

            msg = pyzmail.PyzMessage.factory(raw_data[0][1])

            subject   = msg.get_subject() or "(No Subject)"
            addresses = msg.get_addresses("from")
            sender    = addresses[0][1] if addresses else "unknown"

            if msg.text_part:
                charset = msg.text_part.charset or "utf-8"
                body = msg.text_part.get_payload().decode(charset, errors="ignore")
            else:
                body = "(No text body)"

            # ── STEP 2: Download attachments ──────────────────────────────
            log("[STEP 2] Checking for attachments …")
            attachment_paths = extract_attachments(msg)

            if not attachment_paths:
                log(f"  [STEP 2] No supported attachments from {sender} — skipping.")
                server.store(uid, "+FLAGS", "\\Seen")
                continue

            log(f"  [STEP 2] {len(attachment_paths)} attachment(s) saved.")

            # ── Steps 3-6 ─────────────────────────────────────────────────
            handle_email(subject, sender, body, attachment_paths)

            # Mark as read
            server.store(uid, "+FLAGS", "\\Seen")


# ─────────────────────────────────────────────────────────────────────────────
# Main loop
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    log("=" * 60)
    log("  Email Pipeline Starting")
    log(f"  Attachment dir : {ATTACHMENT_DIR}/")
    log(f"  Poll interval  : {CHECK_INTERVAL}s")
    log("=" * 60)

    while True:
        try:
            check_inbox()
        except Exception as e:
            log(f"Error in polling loop: {e}")
        log(f"Sleeping {CHECK_INTERVAL}s …\n")
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()

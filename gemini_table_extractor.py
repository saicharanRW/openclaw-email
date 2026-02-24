"""
gemini_table_extractor.py

Pipeline:
  1. Convert the input DOCX → plain TXT (using existing extract_text logic,
     with Krutidev → Unicode conversion).
  2. Save the TXT file next to the DOCX with a .txt extension.
  3. Send the TXT content to Google Gemini and ask it to identify every table
     structure present in the document.
  4. Print Gemini's response (the table structures) to the terminal.

Usage:
    python gemini_table_extractor.py <path_to_docx>
"""

import sys
import os
import argparse

# ── Hardcoded API key & model ────────────────────────────────────────────────
GEMINI_API_KEY = "AIzaSyBpWGX8oBBwSKj821UE8apI7bTAKFxqE6o"
GEMINI_MODEL   = "gemini-2.5-flash"

# ── 1. Gemini SDK ──────────────────────────────────────────────────────────────
try:
    from google import genai
except ImportError:
    print(
        "ERROR: google-genai package not found.\n"
        "Install it with:  pip install google-genai",
        file=sys.stderr,
    )
    sys.exit(1)

# ── 2. Local modules ───────────────────────────────────────────────────────────
from extract_text import extract_text          # existing extractor


# ─────────────────────────────────────────────────────────────────────────────
# Helper: convert DOCX → TXT and write the TXT file
# ─────────────────────────────────────────────────────────────────────────────
def docx_to_txt(docx_path: str) -> tuple[str, str]:
    """
    Convert *docx_path* to plain text (Krutidev → Unicode included).

    Returns
    -------
    txt_path : str
        Path of the saved .txt file.
    text : str
        The extracted text content.
    """
    if not os.path.isfile(docx_path):
        print(f"ERROR: File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    print(f"[1/3] Extracting text from:  {docx_path}")
    text = extract_text(docx_path, convert_krutidev=True)

    # Save next to the original DOCX, same name but .txt
    base = os.path.splitext(docx_path)[0]
    txt_path = base + ".txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"      TXT file saved to:      {txt_path}")
    return txt_path, text


# ─────────────────────────────────────────────────────────────────────────────
# Helper: send text to Gemini and ask for table structures
# ─────────────────────────────────────────────────────────────────────────────
GEMINI_PROMPT_TEMPLATE = """
You are a document analysis assistant.

Below is the plain-text content extracted from a Word document (DOCX).
The tables in the document were serialised with cells separated by ' | ' and
rows on separate lines.

Your task:
  1. Identify EVERY table present in the text.
  2. For each table:
     - Give it a sequential number (Table 1, Table 2, …).
     - Show its column headers (if detectable).
     - Show all data rows, formatted as a clean Markdown table.
     - If the table has a title or heading nearby, include it.

Only describe what you actually find in the text. Do not invent data.

--- DOCUMENT TEXT START ---
{document_text}
--- DOCUMENT TEXT END ---

Now list all table structures found in the document.
"""


def ask_gemini_for_tables(text: str, api_key: str) -> str:
    """
    Send *text* to Gemini and return its analysis of all table structures.
    """
    client = genai.Client(api_key=GEMINI_API_KEY)

    prompt = GEMINI_PROMPT_TEMPLATE.format(document_text=text)

    print(f"[2/3] Sending TXT content to Gemini ({GEMINI_MODEL}) …")
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
    )
    return response.text


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description=(
            "Convert a DOCX to TXT, send it to Gemini, and print all "
            "table structures found in the document."
        )
    )
    parser.add_argument("docx_file", help="Path to the input DOCX file")
    args = parser.parse_args()

    # Step 1 – DOCX → TXT
    txt_path, text = docx_to_txt(args.docx_file)

    if not text.strip():
        print("WARNING: Extracted text is empty. Nothing to send to Gemini.")
        sys.exit(0)

    # Step 2 – Send to Gemini
    result = ask_gemini_for_tables(text, GEMINI_API_KEY)

    # Step 3 – Print result
    print("\n[3/3] Gemini's table analysis:\n")
    print("=" * 70)
    print(result)
    print("=" * 70)


if __name__ == "__main__":
    main()

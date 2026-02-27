---
name: email-context-responder
description: Email context responder for document queries - Multi-format attachment support (PDF, DOCX, XLSX, CSV)
---

# Email Context Responder

## 🚨 ABSOLUTE REQUIREMENT - READ FIRST 🚨

**YOUR ENTIRE RESPONSE MUST BE VALID JSON AND NOTHING ELSE.**

Do not write:
```
Dear Sir/Madam, Thank you for...
```

Instead write:
```json
{"category":"legal","priority":"high","requires_reply":true,"language_of_reply":"English","matched_documents":[],"data_used_in_reply":{"facts":[],"source_files":[]},"table_data":null,"tables":[],"suggested_reply":"Dear Sir/Madam, Thank you for...","reply_note":"No documents found"}
```

**START YOUR RESPONSE WITH `{` AND END WITH `}`**

**NO TEXT BEFORE THE `{`**

**NO TEXT AFTER THE `}`**

---

## CRITICAL: How This Skill Works with the Email Pipeline

The email pipeline:
1. Extracts text from **all attachment types** (PDF, DOCX, XLSX, CSV) using document indexer functions
2. Sends extracted text to **OpenAI** to detect tables and extract column headers
3. Sends extracted text to **OpenAI** for document brief analysis
4. Sends email context + **AVAILABLE TABLES & COLUMNS** + **DOCUMENT BRIEF** to this skill (via OpenClaw)

When invoked, this skill MUST follow these steps in order:

1. **Analyze the email first** - Extract keywords, file names, and search terms
2. **Capture columns** - Get the columns from AVAILABLE TABLES & COLUMNS in the prompt
3. **Build a targeted DB query** - Use extracted keywords to filter, NOT fetch everything
4. **Create a Python script file** and execute it using bash_tool
5. **Load the results** and map data to the detected columns
6. **Compose the reply**

**DO NOT** try to run Python inline with `python3 -c`. **CREATE A FILE INSTEAD.**

---

## 🔒 STRICT DATA INTEGRITY RULES (MANDATORY)

**You MUST follow these rules. Inconsistent or fabricated data is unacceptable.**

### 1. SOURCE-ONLY DATA
- **Every fact, every table cell, every value** MUST come directly from the database query results (`extracted_text`, `keywords`, `ai_analysis`).
- **NEVER** infer, guess, hallucinate, or invent data.
- **NEVER** fill in values based on patterns or assumptions.
- If a value is not explicitly present in the source text, use **"N/A"**.

### 2. CONSISTENT SEARCH
- Use the **same keywords** derived from the document brief every time.
- Always run the **database query first** — do not skip it.
- Do not improvise alternative search terms or add synonyms on the fly.
- If the query returns nothing, try broadening keywords methodically (AND → OR); do not fabricate results.

### 3. N/A FOR MISSING DATA (REQUIRED)
- When there is **no data** for a specific table cell or column, **always** use **"N/A"** (never empty string `""`, never `null`, never blank).
- If a column has no matching value in the source, the cell value is **"N/A"**.
- If the entire row has no data, every cell in that row is **"N/A"**.
- **Never leave cells blank.** Empty = **"N/A"**.

### 4. VERIFICATION
- Before including any value in `facts` or `tables`, verify it appears in the source `extracted_text`.
- Do not paraphrase — extract the exact wording or numbers from the source.
- If uncertain whether data exists, treat it as missing and use **"N/A"**.

---

## Step 0: Analyze the Email (DO THIS FIRST)

Before touching the database, analyze the incoming email to extract:

1. **Exact file names** - Look for any document/file names mentioned (e.g., "MBC No.75 ,Dt. 01-12-2026.docx")
2. **Keywords** - Extract significant words from subject + body + **document brief** (main_topics, data_domain)
3. **Language** - Detect if the email is in English, Hindi, Marathi, etc.
4. **Priority** - Determine urgency from words like "urgent", "immediately", "ASAP"
5. **Domain** - Categorize the request (admin, traffic, legal, etc.) - **USE document_type from brief**
6. **AVAILABLE TABLES & COLUMNS** - If present in the prompt, capture per-table column headers (keys like filename_table_1, table_1, etc.)
7. **Document brief** - If `DOCUMENT BRIEF` is present, extract the summary, main topics, and key entities

**Example of what you receive:**

```
AVAILABLE TABLES & COLUMNS (extracted from attachment):
{"filename_table_1": ["अ.क्र", "पो. ठाणे", "तपासी अधिकारी", ...]}

DOCUMENT BRIEF (from attachment analysis):
Document Type: Meeting Circular
Data Content Summary: Contains meeting schedule, attendance requirements, and data submission deadlines - does NOT contain actual accused records or case data
Main Topics: meeting announcement, deadline, attendance, submission requirements
Data Domain: administrative, procedural, meeting coordination
Key Entities:
  - Dates: 2026-02-21
  - Locations: South Control Room, Mumbai
  - Reference Numbers: 33/2026
  - People: Rajeshkumar Gatthe

EMAIL:
Subject : "Give me the info related to a doc"
From : "openclawdummy@gmail.com"
Body : Give me the info related to the attached document.
```

**Output of this step** (used to build the DB query):
```json
{
  "exact_file_name": null,
  "search_keywords": ["accused", "arrested", "arrest", "crime", "case", "police", "investigation", "NDPS", "women", "cybercrime"],
  "language": "English",
  "priority": "high",
  "domain": "legal/police",
  "available_tables_columns": {"filename_table_1": ["अ.क्र", "पो. ठाणे", ...]},
  "document_brief": {
    "type": "Meeting Circular",
    "data_needed": "arrested accused records, case information, crime types"
  }
}
```

**Why these keywords?**
- The attached document is a MEETING CIRCULAR that REQUESTS data about accused/cases
- The user wants the ACTUAL DATA (case records), not more meeting announcements
- So we search for documents that CONTAIN: "accused", "arrested", "crime", "case"
- We do NOT search for: "meeting", "circular", "announcement", "deadline"

**CRITICAL - Keyword Extraction Strategy:**

The document brief tells you **what kind of data the attachment contains or requests**. Use this to search for similar documents in the database.

**From Document Brief, extract:**
1. **data_domain** - What category of data: ["police records", "financial transactions", "parking violations"]
2. **main_topics** - Specific data elements: ["arrested accused", "case numbers", "investigation officers"]
3. **key_entities** - Concrete identifiers: dates, locations, reference numbers, names

**Example:**
```
Document Type: Meeting Circular
Data Content: "Meeting announcement for accused exchange program, NO actual case data"
Main Topics: ["meeting announcement", "deadline", "submission requirements"]
Data Domain: ["administrative", "procedural"]
```

**Your keywords should be:** ["accused", "arrests", "cases", "crime", "police"] (NOT "meeting", "announcement", "deadline")

**Why?** The user wants CASE DATA, not meeting announcements. Search for documents that CONTAIN the data, not documents that ASK FOR the data.

**DO NOT use:**
- Generic words: "info", "document", "attached", "give me"
- Administrative terms: "meeting", "circular", "submission", "deadline", "announcement"
- Procedural terms: "requirement", "instruction", "attendance"

**DO use:**
- Data domain terms from brief
- Main topics (the actual data elements)
- Key entities (specific identifiers)
- Concrete nouns (accused, cases, violations, amounts, records)

---

## Step 1: Query Database with Keywords (CORRECT METHOD)

Use the keywords extracted in Step 0 to build a **targeted SQL query**. Do NOT fetch the entire database.

**STRICT:** Use the same keywords from Step 0 consistently. Do not improvise synonyms or alternative terms. Run the query and use ONLY the results returned — never supplement with guessed data.

**Database Location:** `/home/randomwalk/.openclaw/workspace/multilingual-document-processor/scripts/document_index.db`

**Strategy (in priority order):**
1. If an exact file name is found in the email, search by `file_name` first
2. If no exact match, use keywords to filter via `LIKE` on `file_name`, `keywords`, and `extracted_text`
3. Only if keyword search returns nothing, broaden the search slightly

**Example Query Script:**

```python
# First, create the query script
cat > /tmp/query_db.py << 'SCRIPT_END'
import sqlite3
import json
import sys

DB_PATH = '/home/randomwalk/.openclaw/workspace/multilingual-document-processor/scripts/document_index.db'

# --- REPLACE THESE WITH ACTUAL VALUES FROM EMAIL ANALYSIS ---
# Example: If email mentions "parking violations for October 2024"
EXACT_FILE_NAME = ""       # Usually empty unless specific file mentioned
SEARCH_KEYWORDS = ["parking", "violations", "october", "2024"]  # REPLACE with actual keywords
# -------------------------------------------------------------

try:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    candidates = []

    # Strategy 1: Exact file name match
    if EXACT_FILE_NAME:
        cursor.execute("""
            SELECT id, file_name, keywords, extracted_text, ai_analysis,
                   document_type, domain, confidence
            FROM documents
            WHERE file_name = ? AND extracted_text IS NOT NULL
        """, (EXACT_FILE_NAME,))
        candidates = cursor.fetchall()

    # Strategy 2: Keyword-based search (AND logic - all keywords must match)
    if not candidates and SEARCH_KEYWORDS:
        where_clauses = []
        params = []
        for kw in SEARCH_KEYWORDS:
            where_clauses.append(
                "(file_name LIKE ? OR keywords LIKE ? OR extracted_text LIKE ?)"
            )
            wild = f"%{kw}%"
            params.extend([wild, wild, wild])

        query = f"""
            SELECT id, file_name, keywords, extracted_text, ai_analysis,
                   document_type, domain, confidence
            FROM documents
            WHERE extracted_text IS NOT NULL AND ({' AND '.join(where_clauses)})
            ORDER BY indexed_timestamp DESC
            LIMIT 10
        """
        cursor.execute(query, params)
        candidates = cursor.fetchall()

    # Strategy 3: Broader OR-based search if AND returned nothing
    if not candidates and SEARCH_KEYWORDS:
        where_clauses = []
        params = []
        for kw in SEARCH_KEYWORDS:
            where_clauses.append(
                "(file_name LIKE ? OR keywords LIKE ? OR extracted_text LIKE ?)"
            )
            wild = f"%{kw}%"
            params.extend([wild, wild, wild])

        query = f"""
            SELECT id, file_name, keywords, extracted_text, ai_analysis,
                   document_type, domain, confidence
            FROM documents
            WHERE extracted_text IS NOT NULL AND ({' OR '.join(where_clauses)})
            ORDER BY indexed_timestamp DESC
            LIMIT 10
        """
        cursor.execute(query, params)
        candidates = cursor.fetchall()

    conn.close()

    docs = []
    for row in candidates:
        ai_analysis = None
        if row[4]:
            try:
                ai_analysis = json.loads(row[4])
            except:
                pass
        docs.append({
            "id": row[0],
            "file_name": row[1],
            "keywords": row[2] or "",
            "extracted_text": (row[3] or "")[:10000],  # First 10k chars
            "ai_analysis": ai_analysis,
            "document_type": row[5] or "",
            "domain": row[6] or "",
            "confidence": row[7] or "medium"
        })

    output = {
        "success": True,
        "count": len(docs),
        "search_strategy": "exact_name" if EXACT_FILE_NAME and docs else "keyword",
        "documents": docs
    }

    with open('/tmp/db_results.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Success: Found {len(docs)} documents")

except Exception as e:
    output = {
        "success": False,
        "error": str(e),
        "documents": []
    }
    with open('/tmp/db_results.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"Error: {e}")
    sys.exit(1)
SCRIPT_END

# Then execute it
python3 /tmp/query_db.py
```

**CRITICAL:** Before creating the script:
1. Extract actual keywords from the email in Step 0
2. Replace `SEARCH_KEYWORDS = [...]` with the real keywords
3. Do NOT leave it as an empty list or example values

After running this, load the results:
```bash
cat /tmp/db_results.json
```

---

## Step 2: Process Query Results

Load the results in your response:

```python
import json

with open('/tmp/db_results.json', 'r', encoding='utf-8') as f:
    results = json.load(f)

if results['success']:
    candidates = results['documents']
    document_count = len(candidates)
    # Continue processing...
else:
    # Handle no documents case
    candidates = []
    document_count = 0
```

---

## Step 3: Rank Matched Documents

For each candidate document returned by the query:

**Relevance Scoring:**
1. **Exact file name match** = "high" relevance
2. **Multiple keyword matches** (3+ keywords) = "high" relevance  
3. **2 keyword matches** = "medium" relevance
4. **Single keyword match** = "low" relevance

**Match Reason Examples:**
- "Exact file name match"
- "Keywords found: parking, violations, october"
- "Domain match: traffic"

---

## Step 4: Extract Facts from Matched Documents

If matched documents found:

1. **Parse extracted_text** from each document — extract ONLY from source
2. **Extract relevant sections** that relate to the email query — verify each fact exists in text
3. **Format as concise facts** (bullet points) — use exact values from source, no paraphrasing
4. **Attribute each fact** to its source file
5. **STRICT:** Do not add facts not present in the source. Missing data = omit from facts (or state "N/A" in table cells)

Example:
```json
{
  "facts": [
    "Total parking violations in October 2024: 156",
    "Fine amount collected: ₹45,600",
    "Most common violation: No helmet (78 cases)"
  ],
  "source_files": [
    "parking_violations_oct_2024.xlsx",
    "fine_collection_report.pdf"
  ]
}
```

If no matches:
```json
{
  "facts": [],
  "source_files": []
}
```

---

## Step 5: Map Data to Columns & Build Tables

**ONLY if AVAILABLE TABLES & COLUMNS were provided in the prompt:**

The attachment analysis has identified column headers from the attachment. Your job is to:

1. **Parse the matched documents** for data that fits these columns — extract ONLY from `extracted_text`.
2. **Create data rows** by mapping facts to columns — use exact values from source.
3. **Fill each column** with the best matching value from the extracted text.
4. **Use "N/A"** — if no matching data found for a cell, **always** use `"N/A"` (never `""`, never blank).

**Pipeline expects `tables`** — a list of objects, one per table key in AVAILABLE TABLES & COLUMNS:

```json
{
  "tables": [
    {
      "title": "Table 1",
      "headers": ["Sr. No.", "Vehicle No.", "Violation Type", "Fine Amount"],
      "rows": [
        ["1", "MH-12-AB-1234", "No Helmet", "500"],
        ["2", "MH-12-CD-5678", "No Parking", "300"],
        ["3", "MH-12-EF-9012", "Signal Jump", "1000"]
      ]
    }
  ]
}
```

**Legacy `table_data` format** (single table, for backwards compatibility):
```json
{
  "table_data": {
    "headers": ["Sr. No.", "Name", "Date", "Status"],
    "rows": [
      ["1", "John Doe", "2024-10-15", "Paid"],
      ["2", "Jane Smith", "2024-10-16", "Pending"]
    ]
  }
}
```

**If no columns were provided:**
```json
{
  "table_data": null,
  "tables": []
}
```

**MANDATORY: Use "N/A" for empty cells**
- When a cell has no data from the source, use **"N/A"** — never `""`, never blank, never `null`.
- Example row with partial data: `["1", "MH-12-AB-1234", "No Helmet", "N/A"]` (Fine amount unknown → N/A)

**Column Mapping Strategy (STRICT):**

1. **Extract ONLY from source** — Every cell value must appear in `extracted_text`. No guesses.
2. **Exact matches first** — Search for the literal value in the text (e.g., "Vehicle No." column → find vehicle numbers in text)
3. **Pattern matching** — Use only patterns that exist in the text (dates, numbers, IDs)
4. **No data = "N/A"** — If a column has no matching value, use **"N/A"** (mandatory, never blank)
5. **No fabrication** — If you cannot find the value in source, do NOT invent it. Use **"N/A"**.
6. **Limit to 50 rows** per table

**Important:** The pipeline creates reply in the SAME FORMAT as the received attachment:
- Received PDF → Create reply PDF
- Received DOCX → Create reply DOCX
- Received XLSX → Create reply XLSX
- Received CSV → Create reply CSV

---

## Step 6: Detect Email Language

From the email body, detect the language:
- **English** - Most common
- **Hindi** - Check for Devanagari script (unicode range 0x0900-0x097F)
- **Marathi** - Also Devanagari, context clues needed
- **Mixed** - Multiple languages present

Set `language_of_reply` to match the detected language.

---

## Step 7: Compose Reply

**🚨 CRITICAL: Reply goes in JSON `suggested_reply` field, NOT as plain text 🚨**

### When Documents ARE Found:

Put this in the `suggested_reply` field of your JSON:
```
Dear Sir/Madam,

Thank you for your inquiry regarding [brief description of query].

The requested information has been compiled from our indexed documents and is attached for your reference.

[If tables provided: The data has been structured according to the format in your attached document.]

Please review the attached document for complete details.

If any clarification is required, kindly let us know.

Regards,
Office
```

### When Documents are NOT Found:

**STILL RETURN JSON!** Put this in the `suggested_reply` field:

**Example of correct response when no documents found:**
```json
{
  "category": "legal",
  "priority": "high",
  "requires_reply": true,
  "language_of_reply": "English",
  "matched_documents": [],
  "data_used_in_reply": {
    "facts": [],
    "source_files": []
  },
  "table_data": null,
  "tables": [],
  "suggested_reply": "Dear Sir/Madam,\n\nThank you for your inquiry regarding the attached document.\n\nWe are currently unable to locate the requested information in our indexed documents that directly matches the specific details provided.\n\nThis may require manual retrieval from our records. We will review your request and provide the information as soon as possible.\n\nRegards,\nOffice",
  "reply_note": "No matching documents found in database"
}
```

**DO NOT write the reply as plain text. It MUST be inside the JSON structure.**

### Language-Specific Replies:

**Hindi:**
```
प्रिय महोदय/महोदया,

[query] के संबंध में आपकी पूछताछ के लिए धन्यवाद।

अनुरोधित जानकारी संकलित कर दी गई है और संदर्भ के लिए संलग्न है।

कृपया पूर्ण विवरण के लिए संलग्न दस्तावेज़ देखें।

यदि किसी स्पष्टीकरण की आवश्यकता हो तो कृपया हमें बताएं।

सादर,
कार्यालय
```

---

## 🔥 FINAL CHECKLIST BEFORE RESPONDING 🔥

Before you submit your response, verify:

- [ ] ✅ My response starts with `{`
- [ ] ✅ My response ends with `}`
- [ ] ✅ There is NO text before the `{`
- [ ] ✅ There is NO text after the `}`
- [ ] ✅ I did NOT write "Dear Sir/Madam" as plain text
- [ ] ✅ All text content is inside the `suggested_reply` field
- [ ] ✅ The JSON is valid and parseable
- [ ] ✅ I included all required fields: category, priority, requires_reply, language_of_reply, matched_documents, data_used_in_reply, table_data, tables, suggested_reply, reply_note
- [ ] ✅ Every empty/missing table cell is **"N/A"** (never blank or "")
- [ ] ✅ Every fact and table value comes from the source — no fabrication

**If you answered NO to any of these, STOP and fix your response!**

---

## Output Format (FINAL JSON)

**🚨 CRITICAL - READ THIS CAREFULLY 🚨**

Your response MUST be ONLY valid JSON. No other text is allowed.

**❌ DO NOT:**
- Add any explanatory text before the JSON
- Add any explanatory text after the JSON
- Wrap the JSON in markdown code blocks (```json ... ```)
- Include any preamble like "Here is the response:" or "Based on the analysis:"
- Return plain text instead of JSON

**✅ DO:**
- Return ONLY the raw JSON object
- Start your response with `{` 
- End your response with `}`
- Nothing before the `{` and nothing after the `}`

**Example of CORRECT response:**
```
{"category":"legal","priority":"high","requires_reply":true...}
```

**Example of WRONG response:**
```
Based on my analysis, here is the JSON:
```json
{"category":"legal"...}
```
```

**The ONLY valid output:**

```json
{
  "category": "traffic|admin|legal|education|finance|other",
  "priority": "low|medium|high|urgent",
  "requires_reply": true,
  "language_of_reply": "English|Hindi|Marathi|Mixed",
  "matched_documents": [
    {
      "id": 123,
      "file_name": "parking_report.xlsx",
      "match_reason": "Keywords: parking, violations, october",
      "relevance": "high"
    }
  ],
  "data_used_in_reply": {
    "facts": [
      "Total violations: 156",
      "Fine collected: ₹45,600"
    ],
    "source_files": [
      "parking_report.xlsx"
    ]
  },
  "table_data": {
    "headers": ["Sr. No.", "Name", "Date", "Status"],
    "rows": [
      ["1", "John Doe", "2024-10-15", "Paid"],
      ["2", "Jane Smith", "2024-10-16", "Pending"]
    ]
  },
  "tables": [
    {
      "title": "Table Title",
      "headers": ["H1", "H2"],
      "rows": [["r1c1", "r1c2"], ["r2c1", "r2c2"]]
    }
  ],
  "suggested_reply": "Dear Sir/Madam,...[full reply text]...",
  "reply_note": "Found 1 matching document with 2 data rows"
}
```

**Field Explanations:**
- **category**: Domain classification based on keywords and content
- **priority**: Urgency level determined from email language
- **requires_reply**: Always `true` for this workflow
- **language_of_reply**: Match the email's language
- **matched_documents**: List of DB records that matched the query
- **data_used_in_reply.facts**: Extracted information as bullet points
- **data_used_in_reply.source_files**: Attribution to source documents
- **table_data**: Legacy single-table format (or null)
- **tables**: Primary format - list of {title, headers, rows} - one per table in AVAILABLE TABLES; **empty cells = "N/A"**
- **suggested_reply**: Full formatted reply text
- **reply_note**: Internal note for logging/debugging

---

## IMPORTANT REMINDERS

1. ✅ **ALWAYS create script file first** - Don't use `python3 -c`
2. ✅ **ALWAYS use quotes around paths** - `"/path/to/db"` not `/path/to/db`
3. ✅ **ALWAYS handle errors** - Check if query succeeded before processing
4. ✅ **NEVER fabricate facts** - Only use actual database content; if not in source → **"N/A"**
5. ✅ **ALWAYS use "N/A" for empty cells** - Never blank, never `""`, never `null` — always `"N/A"`
6. ✅ **ALWAYS provide source attribution** - Tell user which file(s) data came from
7. ✅ **Extract from source only** - Verify every value exists in `extracted_text` before including
8. ✅ **Consistent search** - Same keywords, same query logic; do not improvise
9. ✅ **Map to detected columns** - Structure data accordingly; missing = **"N/A"**
10. ✅ **Handle empty results gracefully** - Provide helpful "not found" message; do not invent data

---

## Example Workflow

**Email Received:**
```
Subject: Parking violations data for October
Body: Please send the parking violation records for October 2024.
Attachment: parking_template.xlsx (AVAILABLE TABLES: {"filename_table_1": ["Sr. No.", "Vehicle No.", "Date", "Fine"]})
```

**Step 0 - Analysis:**
```json
{
  "search_keywords": ["parking", "violations", "october", "2024"],
  "available_tables_columns": {"filename_table_1": ["Sr. No.", "Vehicle No.", "Date", "Fine"]},
  "language": "English",
  "priority": "medium",
  "domain": "traffic"
}
```

**Step 1 - Query DB:**
- Search for documents containing: parking AND violations AND october AND 2024
- Find: `parking_violations_oct_2024.xlsx`

**Step 2-4 - Extract & Rank:**
- Matched 1 document with high relevance
- Extract facts: "156 total violations, ₹45,600 collected"

**Step 5 - Map to Columns:**
- Parse extracted_text for records matching the column structure
- Create `tables` with one entry for filename_table_1: headers + rows
- **Any missing cell value → use "N/A"**

**Step 6-7 - Compose Reply:**
- Language: English
- Include facts and tables
- Professional tone

**Final Output:**
- JSON with all fields populated
- `tables` contains structured rows (one per table key)
- suggested_reply has complete message
- Pipeline creates reply XLSX (same format as received)

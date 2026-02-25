"""
config.py
---------
Single source of truth for all configuration.
"""

# ── OpenClaw (email_bridge / email_pipeline HTTP calls) ───────────────────────
OPENCLAW_URL = "http://localhost:18790/v1/chat/completions"
TOKEN        = "e2d7d414ae5341285a8a9ba28fce3b169370efd51e370b4a"

# ── IMAP (incoming mail) ──────────────────────────────────────────────────────
IMAP_HOST = "imap.gmail.com"
EMAIL     = "techonomicshub@gmail.com"
PASSWORD  = "digt rpxm gbij wqst"

# ── SMTP (outgoing mail) ──────────────────────────────────────────────────────
SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_EMAIL    = EMAIL
SMTP_PASSWORD = PASSWORD

# ── Gemini ─────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = "AIzaSyDdUlTDsIIwpT-LPuMED52aqhWaZSSolH4"
GEMINI_MODEL   = "gemini-2.5-flash"

# ── File paths ─────────────────────────────────────────────────────────────────
ATTACHMENT_DIR = "attachments"          # folder where attachments are saved
PROCESSED_LOG  = "processed_files.txt"  # tracks already-processed files

# ── Polling ────────────────────────────────────────────────────────────────────
CHECK_INTERVAL = 30   # seconds between inbox polls

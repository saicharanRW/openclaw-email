"""
config.py
---------
Loads all configuration from the .env file (or environment variables).
Install dependency once: pip install python-dotenv
"""

import os
from dotenv import load_dotenv

# Load .env from the same directory as this file
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))

# ── OpenClaw ──────────────────────────────────────────────────────────────────
OPENCLAW_URL = os.environ["OPENCLAW_URL"]
TOKEN        = os.environ["TOKEN"]

# ── IMAP (incoming mail) ──────────────────────────────────────────────────────
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
EMAIL     = os.environ["EMAIL"]
PASSWORD  = os.environ["PASSWORD"]

# ── SMTP (outgoing mail) ──────────────────────────────────────────────────────
SMTP_SERVER   = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_EMAIL    = os.environ.get("SMTP_EMAIL", EMAIL)
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", PASSWORD)

# ── Gemini ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
GEMINI_MODEL   = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

# ── File paths ────────────────────────────────────────────────────────────────
ATTACHMENT_DIR = os.environ.get("ATTACHMENT_DIR", "attachments")
PROCESSED_LOG  = os.environ.get("PROCESSED_LOG", "processed_files.txt")

# ── Polling ───────────────────────────────────────────────────────────────────
CHECK_INTERVAL = int(os.environ.get("CHECK_INTERVAL", "30"))
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

OPENAI_MODEL = "gpt-4o"
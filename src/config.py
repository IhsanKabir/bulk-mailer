"""Configuration constants and on-disk paths.

Everything the app persists lives under a single per-user folder so an
uninstall is just deleting that directory. Nothing is written next to the
.exe (which may sit in a read-only location).
"""

from pathlib import Path

# Per-user state directory (mirrors the %LOCALAPPDATA% convention).
APP_DIR = Path.home() / "AppData" / "Local" / "BulkMailer"

# Resume-safe send-log: (campaign, email, subject) → SENT/DRAFTED/FAILED.
MAILER_LOG_DB = APP_DIR / "mailer_log.sqlite"

# Last-used Outlook "send from" account, so the picker pre-selects it.
OUTLOOK_ACCOUNT_FILE = APP_DIR / "outlook_account.txt"

# Last-used transport ("outlook"/"graph"/"smtp") + mode ("draft"/"send"), so the
# app reopens exactly as last used. Added after a real incident: the defaults
# (Outlook + Create drafts) made a user believe mails were SENT when only
# drafts were created on a machine where desktop Outlook can't even sign in.
UI_STATE_FILE = APP_DIR / "ui_state.json"

# Microsoft Graph device-code token cache (see graph_mailer._token_cache_path).
# graph_mailer reads `config.APP_DIR`, so the token lands beside the rest.

# Rotating log.
LOG_FILE = APP_DIR / "bulk_mailer.log"

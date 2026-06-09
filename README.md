# Bulk Mailer

Send **one personalised email per recipient** — each with its own
attachment — from a single Excel list. Windows desktop app, single `.exe`,
no Python or admin required.

[Download the latest release →](../../releases)

---

## What it does

Point it at a mapping sheet and an attachments folder, write the message
once, and it composes a separate email for every row — right recipient,
right file, name filled in. **Draft** them all for review first, or
**Send** straight away.

No recipient cap. 40 or 1,000+ in one pass — the only ceiling is your mail
provider's daily send limit.

## Quick start

1. Download `BulkMailer.exe` from [Releases](../../releases).
2. Double-click. Click *More info → Run anyway* on the SmartScreen warning.
3. **Mapping Excel** → pick your list. **Attachments folder** → pick the folder.
4. **Load + preview** — every row is checked (valid email, file found).
5. Write the subject + body, then **Create drafts** (review first) or **Send now**.

## The mapping sheet

One row per recipient. Header names are matched case-insensitively; only
**Email** and **File** are required.

| Email | Name | File | CC | BCC |
| --- | --- | --- | --- | --- |
| ops@acme.com | Karim | acme-june.xlsx | mgr@acme.com | |
| rm@globex.com | Nila | globex-june.xlsx | | audit@you.com |

- **File** is resolved against the attachments folder you pick. Separate
  several files with `;` or `|` for multi-attachment rows.
- The body is a template — `{name}` (or any column, e.g. `{Route}`) is
  filled per row. Unknown placeholders are left as-is, never crash a run.

## Send via (pick one)

| Transport | Use when | Notes |
| --- | --- | --- |
| **Outlook desktop** | Outlook is installed and signed in | Sends from any account already added to Outlook. No password needed. |
| **Microsoft 365 (Graph)** | You have an M365 mailbox, no desktop Outlook | One-time browser sign-in (device code, MFA supported). Some tenants require admin approval. |
| **SMTP** | Gmail, Workspace, Yahoo, or any host | Needs an app password (not your login password). Has an MX auto-detect to fill the host. |

## Why it's safe for big runs

- **Preview** validates every row before anything sends — bad emails and
  missing files show red.
- **Draft mode** is the default: review in your Drafts (or as `.eml`
  files for SMTP) before a single message leaves.
- **Skip already-sent** + a resume-safe send-log: if a run dies at row 25
  of 40, re-running skips the 24 already sent — no one gets two copies.
- **Delay slider** spaces out sends so you don't trip provider throttling.
- **Stop** halts cleanly between rows.

## Test before a big blast

Use **Send test to myself** to draft/send the first valid row to an
address you type — confirm the formatting and attachment look right, then
run the full list.

---

<details>
<summary><strong>Persistent state</strong> (in <code>%LOCALAPPDATA%\BulkMailer\</code>)</summary>

| File | Contents |
| --- | --- |
| `mailer_log.sqlite` | Resume-safe send-log (campaign, email, subject → status) |
| `outlook_account.txt` | Last-used Outlook "send from" account |
| `graph_token.bin` | Microsoft Graph token cache (device-code sign-in) |
| `bulk_mailer.log` | Rotating log (5 MB × 2) |

SMTP passwords live in the Windows Credential Manager (via `keyring`),
never on disk.

</details>

<details>
<summary><strong>Build from source</strong></summary>

Requires Python 3.13+ on Windows.

```cmd
pip install -r requirements.txt
build_exe.bat
```

Output: `dist\BulkMailer.exe` (~30-40 MB).

Run from source:

```cmd
python run_app.py
```

Tests: `python -m pytest tests/ -v`.

</details>

<details>
<summary><strong>Architecture</strong></summary>

```text
src/main.py          ── entry point
src/gui.py           ── single-window Tkinter UI (sv_ttk theme)
src/config.py        ── paths
src/mailer_io.py     ── mapping reader + {name} templating + row validation
src/mailer_client.py ── Outlook COM + SMTP backends, MX auto-detect
src/graph_mailer.py  ── Microsoft Graph backend (device-code sign-in)
src/mailer_log.py    ── resume-safe send-log
```

</details>

---

**License:** © 2025-2026 A K M Ihsan Kabir. All Rights Reserved. See [LICENSE](LICENSE).

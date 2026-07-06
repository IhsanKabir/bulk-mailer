"""Single-window Tkinter GUI for the Bulk Mailer.

One screen: pick a mapping Excel + an attachments folder, write the message
once (with {name}/{column} placeholders), preview every row, then DRAFT for
review or SEND. Three transports — Outlook desktop, Microsoft 365 (Graph),
or any SMTP host. Long-running work happens on a worker thread; the GUI is
fed via a queue polled on the Tk main loop.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import (
    __version__, config, graph_mailer, guide_view, mailer_client, mailer_io,
    mailer_split, updater,
)
from .mailer_log import MailerLog
from .health_gui import HealthMixin
from .whatsapp_gui import WhatsAppMixin

log = logging.getLogger(__name__)

# Worker → GUI message kinds.
MSG_MAIL_PROGRESS = "mail_progress"   # (i, total, to, status, row_index)
MSG_MAIL_DONE = "mail_done"           # payload: dict
MSG_MAIL_ERROR = "mail_error"


class MailerApp(WhatsAppMixin, HealthMixin):
    """The whole application: one window, one mailer panel + WhatsApp blast."""

    # Semantic palette — single source of truth for coloured widgets.
    _COLOR_PRIMARY = "#0078D4"
    _COLOR_DANGER = "#C42B1C"
    _COLOR_SUCCESS = "#107C10"
    _COLOR_WARNING = "#9D5D00"
    _COLOR_MUTED = "#64748b"
    _COLOR_SECTION = "#0F6CBD"
    _COLOR_ROW_BAD = "#FDE7E9"
    _COLOR_ROW_GOOD = "#DFF6DD"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(f"Bulk Email Sending  v{__version__}")
        root.geometry("960x820")
        root.minsize(820, 640)

        # ----- Mailer state -----
        # Restore the last-used transport + mode so the app reopens as it was
        # left (a fresh install still defaults to Outlook + drafts — the safe pair).
        ui_state = self._load_ui_state()
        self.mail_mapping_path = tk.StringVar()
        self.mail_attach_dir = tk.StringVar()
        self.mail_subject = tk.StringVar()
        self.mail_mode = tk.StringVar(value=ui_state.get("mode", "draft"))   # "draft" | "send"
        self.mail_delay_s = tk.DoubleVar(value=1.0)
        self.mail_skip_sent = tk.BooleanVar(value=True)
        # Transport default: Outlook desktop — sends from whatever account is
        # already added to Outlook, needing neither SMTP basic-auth nor admin
        # Graph consent (both commonly blocked by locked-down M365 tenants).
        self.mail_transport = tk.StringVar(value=ui_state.get("transport", "outlook"))
        self.mail_smtp_preset = tk.StringVar(value="Gmail / Google Workspace")
        self.mail_smtp_host = tk.StringVar(value="smtp.gmail.com")
        self.mail_smtp_port = tk.IntVar(value=587)
        self.mail_smtp_sender = tk.StringVar()
        self.mail_smtp_password = tk.StringVar()
        self.mail_smtp_remember = tk.BooleanVar(value=True)
        self.mail_outlook_account = tk.StringVar()
        self._graph_session = None
        self.mail_graph_status = tk.StringVar(value="Not signed in")
        self._mail_rows: list = []
        self._mail_worker: threading.Thread | None = None
        self._mail_stop_flag = threading.Event()
        self._mail_log = MailerLog(config.MAILER_LOG_DB)
        # Split & Send: one main sheet, filtered per email address in a column.
        self.mail_split_input = tk.StringVar()
        self.mail_split_sheet = tk.StringVar()
        self.mail_split_column = tk.StringVar()
        self.mail_split_outdir = tk.StringVar()
        self.mail_split_cc = tk.StringVar()            # applied to EVERY message
        self.mail_split_bcc = tk.StringVar()
        self._mail_campaign = ""                       # sent-log campaign label

        self._msg_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._theme_is_dark = False

        self._setup_styles()
        self._build_ui()
        self.root.after(100, self._poll_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

    def _setup_styles(self) -> None:
        """Apply the Sun Valley (Windows-11) theme + semantic styles."""
        style = ttk.Style()
        try:
            import sv_ttk
            sv_ttk.set_theme("dark" if self._theme_is_dark else "light")
        except Exception:  # noqa: BLE001 — degrade gracefully
            log.warning("sv_ttk unavailable; using fallback ttk theme")
            for name in ("vista", "winnative", "clam"):
                if name in style.theme_names():
                    style.theme_use(name)
                    break

        style.configure(
            "Section.TLabel",
            font=("Segoe UI Semibold", 11), foreground=self._COLOR_SECTION,
        )
        style.configure(
            "SectionLg.TLabel",
            font=("Segoe UI Semibold", 15), foreground=self._COLOR_SECTION,
        )
        style.configure(
            "Hint.TLabel", foreground=self._COLOR_MUTED, font=("Segoe UI", 9),
        )

        # Primary action — REAL accent blue. sv_ttk renders blue via a custom
        # layout element tied to "Accent.TButton"; cloning that layout onto our
        # "Primary.TButton" name makes the button actually blue.
        try:
            style.layout("Primary.TButton", style.layout("Accent.TButton"))
            style.configure("Primary.TButton", font=("Segoe UI Semibold", 10))
        except tk.TclError:
            style.configure(
                "Primary.TButton", font=("Segoe UI Semibold", 10),
                foreground="white", background=self._COLOR_PRIMARY,
            )

        # Danger — bold red text (sv_ttk ignores TButton background fills).
        style.configure(
            "Danger.TButton", font=("Segoe UI Semibold", 10),
            foreground=self._COLOR_DANGER,
        )
        try:
            style.map("Danger.TButton",
                      foreground=[("active", "#A82A1C"), ("disabled", "#C9A8A4")])
        except tk.TclError:
            pass

        style.configure(
            "Success.TLabel", foreground=self._COLOR_SUCCESS,
            font=("Segoe UI Semibold", 10),
        )

    # ------------------------------------------------------------------
    # Layout helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _make_scrollable(parent: tk.Widget) -> ttk.Frame:
        """Wrap `parent` in a vertical scrollable Canvas; return inner Frame."""
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(window_id, width=e.width),
        )

        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
        inner.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_wheel))
        inner.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
        return inner

    def _section(self, parent: tk.Widget, title: str) -> ttk.Frame:
        """Section heading + separator + inner content frame."""
        wrapper = ttk.Frame(parent)
        wrapper.pack(fill="x", pady=(8, 4), padx=4)
        ttk.Label(wrapper, text=title, style="Section.TLabel").pack(anchor="w")
        ttk.Separator(wrapper, orient="horizontal").pack(fill="x", pady=(2, 6))
        body = ttk.Frame(wrapper)
        body.pack(fill="x")
        return body

    @staticmethod
    def _form_row(
        parent: ttk.Frame, row: int, label: str, widget: tk.Widget,
        *, suffix: tk.Widget | None = None, label_width: int = 16,
    ) -> None:
        ttk.Label(parent, text=label, width=label_width, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(2, 8), pady=4,
        )
        widget.grid(row=row, column=1, sticky="ew", padx=(0, 4), pady=4)
        if suffix is not None:
            suffix.grid(row=row, column=2, padx=(4, 2), pady=4)
        parent.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # Status bar at the bottom.
        status_frm = ttk.Frame(self.root, relief="sunken", borderwidth=1)
        status_frm.pack(side="bottom", fill="x")
        ttk.Label(
            status_frm, text=f"v{__version__}", foreground="#475569",
        ).pack(side="right", padx=(0, 8), pady=2)
        self.btn_theme = ttk.Button(
            status_frm, text="☾  Dark", command=self._toggle_theme, width=10,
        )
        self.btn_theme.pack(side="right", padx=4, pady=2)
        self.btn_check_updates = ttk.Button(
            status_frm, text="Check for updates", command=self._check_for_updates,
        )
        self.btn_check_updates.pack(side="right", padx=4, pady=2)

        # Header.
        header = ttk.Frame(self.root)
        header.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Label(header, text="Bulk Email Sending", style="SectionLg.TLabel").pack(anchor="w")
        ttk.Label(
            header, style="Hint.TLabel", justify="left", wraplength=900,
            text=(
                "One personalised email per recipient from a mapping Excel "
                "(Email · Name · File · optional CC · BCC). Files attach from a "
                "folder you pick. Draft for review, or send."
            ),
        ).pack(anchor="w", pady=(2, 0))

        self._build_mailer_panel(ttk.Frame(self.root))

    def _build_mailer_panel(self, _unused: ttk.Frame) -> None:
        # Email + WhatsApp as sub-tabs so the WhatsApp option is discoverable
        # instead of buried at the bottom of the email form.
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)
        self._nb = nb
        email_tab = ttk.Frame(nb)
        wa_tab = ttk.Frame(nb)
        guide_tab = ttk.Frame(nb)
        nb.add(email_tab, text="Email")
        nb.add(wa_tab, text="WhatsApp")
        nb.add(guide_tab, text="❔ Guide")
        self._build_email_panel(email_tab)
        self._build_whatsapp_section(self._make_scrollable(wa_tab))
        self._guide_tab = guide_tab
        guide_view.build_guide(guide_tab, self, "mailer")

    def _build_email_panel(self, container: ttk.Frame) -> None:
        parent = self._make_scrollable(container)

        # ----- Inputs -----
        io_body = self._section(parent, "Mapping + attachments")
        map_entry = ttk.Entry(io_body, textvariable=self.mail_mapping_path)
        self._form_row(
            io_body, 0, "Mapping Excel:", map_entry,
            suffix=ttk.Button(io_body, text="Browse...", command=self._mail_pick_mapping),
        )
        attach_entry = ttk.Entry(io_body, textvariable=self.mail_attach_dir)
        self._form_row(
            io_body, 1, "Attachments folder:", attach_entry,
            suffix=ttk.Button(io_body, text="Browse...", command=self._mail_pick_attach),
        )

        # ----- Split & Send (no mapping / no separate files needed) -----
        sp_body = self._section(parent, "Split & Send by email column")
        ttk.Label(
            sp_body, style="Hint.TLabel", justify="left", wraplength=880,
            text=(
                "One main sheet only: put each recipient's email address on their "
                "rows. The app detects every address in the column you pick, filters "
                "the sheet, and writes ONE Excel per address into the folder below — "
                "then each recipient is emailed ONLY their own rows (as an attached "
                "workbook). Rows with a blank/invalid address are parked in "
                "_UNMATCHED_ROWS.xlsx, never sent. A cell may hold several addresses "
                "(; , |) — the row goes to each. CC/BCC below apply to every message. "
                "Body/subject placeholders: {email} {name} {rows} {file}."
            ),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=2, pady=(0, 4))
        split_entry = ttk.Entry(sp_body, textvariable=self.mail_split_input)
        self._form_row(
            sp_body, 1, "Main Excel:", split_entry,
            suffix=ttk.Button(sp_body, text="Browse...",
                              command=self._mail_split_pick_input),
        )
        pick_row = ttk.Frame(sp_body)
        self.mail_split_sheet_cb = ttk.Combobox(
            pick_row, textvariable=self.mail_split_sheet, state="readonly", width=26)
        self.mail_split_sheet_cb.pack(side="left")
        self.mail_split_sheet_cb.bind(
            "<<ComboboxSelected>>", lambda _e: self._mail_split_fill_columns())
        ttk.Label(pick_row, text="  Email column:").pack(side="left")
        self.mail_split_column_cb = ttk.Combobox(
            pick_row, textvariable=self.mail_split_column, state="readonly", width=26)
        self.mail_split_column_cb.pack(side="left", padx=(4, 0))
        self._form_row(sp_body, 2, "Sheet:", pick_row)
        out_entry = ttk.Entry(sp_body, textvariable=self.mail_split_outdir)
        self._form_row(
            sp_body, 3, "Split files folder:", out_entry,
            suffix=ttk.Button(sp_body, text="Browse...",
                              command=self._mail_split_pick_outdir),
        )
        ccrow = ttk.Frame(sp_body)
        ttk.Entry(ccrow, textvariable=self.mail_split_cc, width=32).pack(side="left")
        ttk.Label(ccrow, text="  BCC (every message):").pack(side="left")
        ttk.Entry(ccrow, textvariable=self.mail_split_bcc, width=32).pack(
            side="left", padx=(4, 0))
        self._form_row(sp_body, 4, "CC (every message):", ccrow)
        sp_btns = ttk.Frame(sp_body)
        sp_btns.grid(row=5, column=0, columnspan=3, sticky="w", padx=2, pady=(6, 2))
        ttk.Button(sp_btns, text="Create split files",
                   command=lambda: self._mail_split_run(load_into_mailer=False),
                   ).pack(side="left")
        ttk.Button(sp_btns, text="Split + load into mailer", style="Primary.TButton",
                   command=lambda: self._mail_split_run(load_into_mailer=True),
                   ).pack(side="left", padx=(8, 0))
        ttk.Label(
            sp_btns, style="Hint.TLabel",
            text="then use Preview grid + Test / Run below as usual",
        ).pack(side="left", padx=(10, 0))

        # ----- Message -----
        msg_body = self._section(parent, "Message")
        subj_entry = ttk.Entry(msg_body, textvariable=self.mail_subject)
        self._form_row(msg_body, 0, "Subject:", subj_entry)
        ttk.Label(
            msg_body, text="Body (use {name} or any column as a placeholder):",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=2, pady=(6, 2))
        self.mail_body_text = tk.Text(msg_body, height=8, wrap="word")
        self.mail_body_text.grid(row=2, column=0, columnspan=3, sticky="ew", padx=2)
        self.mail_body_text.insert("1.0", (
            "Dear {name},\n\n"
            "Please find attached your report.\n\n"
            "Best regards,\n"
        ))
        msg_body.columnconfigure(0, weight=1)

        # ----- Transport -----
        tx = self._section(parent, "Send via")
        trow = ttk.Frame(tx)
        trow.pack(fill="x", padx=2, pady=(0, 4))
        ttk.Radiobutton(
            trow, text="Outlook desktop", value="outlook",
            variable=self.mail_transport, command=self._mail_sync_transport,
        ).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            trow, text="Microsoft 365 sign-in (Graph)", value="graph",
            variable=self.mail_transport, command=self._mail_sync_transport,
        ).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            trow, text="SMTP (Gmail / Workspace / any)", value="smtp",
            variable=self.mail_transport, command=self._mail_sync_transport,
        ).pack(side="left")

        # Graph sign-in block.
        self.mail_graph_frame = ttk.Frame(tx)
        gr = self.mail_graph_frame
        ttk.Label(
            gr, style="Hint.TLabel", justify="left", wraplength=820,
            text=(
                "Sends as your Microsoft 365 address over HTTPS — no SMTP, no "
                "desktop Outlook. Click Sign in, then enter the code in your "
                "browser (one-time; MFA supported). Some tenants require admin "
                "approval for this."
            ),
        ).pack(anchor="w", pady=(0, 4))
        grow = ttk.Frame(gr)
        grow.pack(fill="x")
        self.btn_mail_graph_signin = ttk.Button(
            grow, text="Sign in to Microsoft 365", command=self._mail_graph_signin,
        )
        self.btn_mail_graph_signin.pack(side="left")
        ttk.Label(grow, textvariable=self.mail_graph_status, style="Hint.TLabel").pack(
            side="left", padx=(10, 0),
        )
        ttk.Button(grow, text="Sign out", command=self._mail_graph_signout).pack(
            side="left", padx=(10, 0),
        )

        # SMTP credential block.
        self.mail_smtp_frame = ttk.Frame(tx)
        sf = self.mail_smtp_frame
        ttk.Label(sf, text="Provider:", width=14, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        preset_cb = ttk.Combobox(
            sf, textvariable=self.mail_smtp_preset, state="readonly",
            values=list(mailer_client.SMTP_PRESETS.keys()), width=30,
        )
        preset_cb.grid(row=0, column=1, sticky="w", pady=3)
        preset_cb.bind("<<ComboboxSelected>>", lambda _e: self._mail_apply_preset())
        ttk.Label(sf, text="Host:", width=14, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(sf, textvariable=self.mail_smtp_host, width=32).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(sf, text="Port:").grid(row=1, column=2, sticky="e", padx=(12, 4))
        ttk.Entry(sf, textvariable=self.mail_smtp_port, width=7).grid(row=1, column=3, sticky="w")
        ttk.Label(sf, text="From (email):", width=14, anchor="w").grid(row=2, column=0, sticky="w", pady=3)
        sender_entry = ttk.Entry(sf, textvariable=self.mail_smtp_sender, width=32)
        sender_entry.grid(row=2, column=1, sticky="w", pady=3)
        sender_entry.bind("<FocusOut>", lambda _e: self._mail_load_saved_password())
        ttk.Button(sf, text="Auto-detect host", command=self._mail_autodetect_host).grid(
            row=2, column=2, columnspan=2, sticky="w", padx=(12, 0),
        )
        ttk.Label(sf, text="Password:", width=14, anchor="w").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(sf, textvariable=self.mail_smtp_password, show="•", width=32).grid(
            row=3, column=1, sticky="w", pady=3,
        )
        ttk.Checkbutton(
            sf, text="Remember (Credential Manager)", variable=self.mail_smtp_remember,
        ).grid(row=3, column=2, columnspan=2, sticky="w", padx=(12, 0))
        ttk.Label(
            sf, style="Hint.TLabel", wraplength=820, justify="left",
            text=(
                "Gmail / Workspace & Office 365 require an APP PASSWORD (enable "
                "2-step verification, then create a 16-char app password). Your "
                "normal login password will be rejected."
            ),
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=(2, 2))

        # Outlook account picker.
        self.mail_outlook_frame = ttk.Frame(tx)
        ttk.Label(
            self.mail_outlook_frame, text="Send from account:", width=16, anchor="w",
        ).pack(side="left")
        self.mail_outlook_combo = ttk.Combobox(
            self.mail_outlook_frame, textvariable=self.mail_outlook_account,
            state="readonly", width=36, values=[],
        )
        self.mail_outlook_combo.pack(side="left", padx=(4, 8))
        self.mail_outlook_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._mail_save_last_outlook_account(self.mail_outlook_account.get()),
        )
        ttk.Button(
            self.mail_outlook_frame, text="Refresh accounts",
            command=self._mail_refresh_outlook_accounts,
        ).pack(side="left")

        # ----- Options -----
        opt = ttk.Frame(parent)
        opt.pack(fill="x", padx=6, pady=(8, 0))
        ttk.Label(opt, text="Mode:").pack(side="left")
        ttk.Radiobutton(
            opt, text="Create drafts (review first)", value="draft",
            variable=self.mail_mode,
        ).pack(side="left", padx=(4, 10))
        ttk.Radiobutton(
            opt, text="Send now", value="send", variable=self.mail_mode,
        ).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(
            opt, text="Skip already-sent rows", variable=self.mail_skip_sent,
        ).pack(side="left", padx=(0, 16))
        ttk.Label(opt, text="Delay (s):").pack(side="left")
        ttk.Spinbox(
            opt, from_=0.0, to=10.0, increment=0.5, width=5,
            textvariable=self.mail_delay_s,
        ).pack(side="left", padx=(4, 0))

        # ----- Actions -----
        act = ttk.Frame(parent)
        act.pack(fill="x", padx=6, pady=(8, 0))
        self.btn_mail_preview = ttk.Button(
            act, text="Load + preview", command=self._mail_preview,
        )
        self.btn_mail_preview.pack(side="left")
        self.btn_mail_test = ttk.Button(
            act, text="Send test to myself", command=self._mail_test, state="disabled",
        )
        self.btn_mail_test.pack(side="left", padx=(8, 0))
        self.btn_mail_run = ttk.Button(
            act, text="Create drafts", style="Primary.TButton",
            command=self._mail_run, state="disabled",
        )
        self.btn_mail_run.pack(side="left", padx=(8, 0))
        self.btn_mail_stop = ttk.Button(
            act, text="Stop", command=self._mail_stop, state="disabled",
            style="Danger.TButton",
        )
        self.btn_mail_stop.pack(side="left", padx=(8, 0))
        ttk.Button(act, text="Health / Diagnostics",
                   command=lambda: self._open_health_dialog("mailer")).pack(
            side="left", padx=(8, 0))
        self.mail_status = ttk.Label(act, text="Idle.", style="Hint.TLabel")
        self.mail_status.pack(side="left", padx=(12, 0))
        self.mail_mode.trace_add("write", lambda *_a: self._mail_sync_run_label())
        # Persist transport + mode whenever they change, so the next launch
        # reopens the app exactly as it was last used.
        self.mail_mode.trace_add("write", lambda *_a: self._save_ui_state())
        self.mail_transport.trace_add("write", lambda *_a: self._save_ui_state())

        # ----- Preview grid -----
        prev = self._section(parent, "Preview")
        grid = ttk.Frame(prev)
        grid.pack(fill="both", expand=True)
        cols = ("row", "email", "name", "files", "cc", "bcc", "status")
        self.mail_tree = ttk.Treeview(grid, columns=cols, show="headings", height=10)
        for cid, txt, w in (
            ("row", "#", 36), ("email", "Email", 200), ("name", "Name", 130),
            ("files", "Attachment(s)", 220), ("cc", "CC", 120),
            ("bcc", "BCC", 120), ("status", "Status", 130),
        ):
            self.mail_tree.heading(cid, text=txt)
            self.mail_tree.column(cid, width=w, anchor="w")
        self.mail_tree.tag_configure("bad", background=self._COLOR_ROW_BAD)
        self.mail_tree.tag_configure("ok", background=self._COLOR_ROW_GOOD)
        self.mail_tree.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
        msb = ttk.Scrollbar(grid, command=self.mail_tree.yview)
        msb.pack(side="left", fill="y")
        self.mail_tree.configure(yscrollcommand=msb.set)

        # ----- Progress -----
        prog = self._section(parent, "Progress")
        self.mail_progress = ttk.Progressbar(prog, mode="determinate", maximum=1)
        self.mail_progress.pack(fill="x", padx=2, pady=2)

        self._mail_sync_transport()

    # ------------------------------------------------------------------
    # Theme
    # ------------------------------------------------------------------

    def _toggle_theme(self) -> None:
        self._theme_is_dark = not self._theme_is_dark
        try:
            import sv_ttk
            sv_ttk.set_theme("dark" if self._theme_is_dark else "light")
        except Exception:  # noqa: BLE001
            pass
        self._setup_styles()
        self.btn_theme.configure(text="☀  Light" if self._theme_is_dark else "☾  Dark")
        self._toggle_theme_guide_rebuild()

    # ------------------------------------------------------------------
    # Check for updates (GitHub-only; opens the download page, no self-swap)
    # ------------------------------------------------------------------

    def _check_for_updates(self) -> None:
        self.btn_check_updates.configure(state="disabled", text="Checking…")

        def worker() -> None:
            info = updater.check_for_update()
            self.root.after(0, lambda: self._show_update_result(info))

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_result(self, info) -> None:
        try:
            self.btn_check_updates.configure(state="normal", text="Check for updates")
        except tk.TclError:
            return
        if info is None:
            messagebox.showwarning(
                "Check for updates",
                "Couldn't reach the update server right now — check your internet "
                f"and try again. You can also visit:\n{updater.RELEASES_PAGE}")
            return
        if not info.is_newer:
            messagebox.showinfo(
                "Check for updates", f"You're on the latest version (v{__version__}).")
            return
        notes = info.notes.strip()
        if len(notes) > 700:
            notes = notes[:700].rstrip() + " …"
        if messagebox.askyesno(
                "Update available",
                f"A newer version is available: v{info.latest_version}\n"
                f"(You have v{__version__}.)\n\n{notes}\n\n"
                "Open the download page in your browser now?"):
            updater.open_download(info)

    def _toggle_theme_guide_rebuild(self) -> None:
        # The Guide's Canvas boxes use theme-fixed colors — rebuild it (cheap,
        # static content) so it re-themes with everything else.
        guide_tab = getattr(self, "_guide_tab", None)
        if guide_tab is not None and guide_tab.winfo_exists():
            for child in guide_tab.winfo_children():
                child.destroy()
            guide_view.build_guide(guide_tab, self, "mailer")

    # ------------------------------------------------------------------
    # UI-state persistence (transport + mode survive restarts)
    # ------------------------------------------------------------------

    @staticmethod
    def _load_ui_state() -> dict:
        """Last-used transport/mode from disk; {} on first run or any error."""
        import json
        try:
            data = json.loads(config.UI_STATE_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def _save_ui_state(self) -> None:
        """Best-effort persist — a failed write must never break the GUI."""
        import json
        try:
            config.APP_DIR.mkdir(parents=True, exist_ok=True)
            config.UI_STATE_FILE.write_text(json.dumps({
                "transport": self.mail_transport.get(),
                "mode": self.mail_mode.get(),
            }), encoding="utf-8")
        except OSError:
            log.warning("Could not save UI state", exc_info=True)

    # ------------------------------------------------------------------
    # Transport switching + credentials
    # ------------------------------------------------------------------

    def _mail_sync_run_label(self) -> None:
        self.btn_mail_run.configure(
            text="Create drafts" if self.mail_mode.get() == "draft" else "Send now",
        )

    def _mail_sync_transport(self) -> None:
        t = self.mail_transport.get()
        self.mail_smtp_frame.pack_forget()
        self.mail_outlook_frame.pack_forget()
        self.mail_graph_frame.pack_forget()
        if t == "smtp":
            self.mail_smtp_frame.pack(fill="x", padx=2)
        elif t == "outlook":
            self.mail_outlook_frame.pack(fill="x", padx=2, pady=(2, 4))
            if not self.mail_outlook_combo.cget("values"):
                self._mail_refresh_outlook_accounts()
        else:  # graph
            self.mail_graph_frame.pack(fill="x", padx=2, pady=(2, 4))
            if self._graph_session is None:
                self._mail_graph_try_silent()

    def _mail_graph_try_silent(self) -> None:
        try:
            sess = graph_mailer.GraphSession.try_silent()
        except Exception:  # noqa: BLE001
            sess = None
        if sess is not None:
            self._graph_session = sess
            self.mail_graph_status.set(f"● Signed in: {sess.account}")

    def _mail_graph_signin(self) -> None:
        self.btn_mail_graph_signin.configure(state="disabled")
        self.mail_graph_status.set("Starting sign-in…")

        def prompt_cb(message: str, code: str) -> None:
            self._post("mail_graph_prompt", (message, code))

        def worker() -> None:
            try:
                sess = graph_mailer.GraphSession.sign_in(prompt_cb=prompt_cb)
                self._post("mail_graph_signed_in", sess)
            except Exception as exc:  # noqa: BLE001
                self._post("mail_graph_signin_failed", str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _mail_graph_signout(self) -> None:
        graph_mailer.GraphSession.sign_out()
        self._graph_session = None
        self.mail_graph_status.set("Not signed in")

    def _mail_load_last_outlook_account(self) -> str:
        try:
            p = config.OUTLOOK_ACCOUNT_FILE
            return p.read_text(encoding="utf-8").strip() if p.exists() else ""
        except Exception:  # noqa: BLE001
            return ""

    def _mail_save_last_outlook_account(self, address: str) -> None:
        try:
            p = config.OUTLOOK_ACCOUNT_FILE
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(address.strip(), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    def _mail_refresh_outlook_accounts(self) -> None:
        accts = mailer_client.list_outlook_accounts()
        self.mail_outlook_combo.configure(values=accts)
        if accts and not self.mail_outlook_account.get():
            last = self._mail_load_last_outlook_account()
            self.mail_outlook_account.set(last if last in accts else accts[0])
        if not accts:
            self.mail_status.configure(
                text="No Outlook accounts found — is Outlook installed/signed in?",
            )

    def _mail_autodetect_host(self) -> None:
        sender = self.mail_smtp_sender.get().strip()
        if "@" not in sender:
            messagebox.showinfo("Auto-detect", "Type the From email address first.")
            return
        self.mail_status.configure(text="Detecting mail host…")
        self.root.update_idletasks()
        info = mailer_client.detect_mail_host(sender)
        if not info:
            self.mail_status.configure(text="Couldn't detect — enter host manually.")
            messagebox.showwarning(
                "Auto-detect",
                "Couldn't resolve the mail host for that domain. "
                "Enter the SMTP host manually.",
            )
            return
        self.mail_smtp_preset.set(info["preset"])
        if info["host"]:
            self.mail_smtp_host.set(info["host"])
        self.mail_smtp_port.set(info["port"])
        self.mail_status.configure(text=f"MX: {info['mx']}")
        if info["note"]:
            messagebox.showinfo("Auto-detect", info["note"])

    def _mail_apply_preset(self) -> None:
        host_port = mailer_client.SMTP_PRESETS.get(self.mail_smtp_preset.get())
        if host_port and host_port[0]:
            self.mail_smtp_host.set(host_port[0])
            self.mail_smtp_port.set(host_port[1])

    def _mail_load_saved_password(self) -> None:
        sender = self.mail_smtp_sender.get().strip()
        if sender and not self.mail_smtp_password.get():
            saved = mailer_client.load_smtp_password(sender)
            if saved:
                self.mail_smtp_password.set(saved)

    def _mail_smtp_settings(self) -> "mailer_client.SMTPSettings | None":
        host = self.mail_smtp_host.get().strip()
        sender = self.mail_smtp_sender.get().strip()
        pwd = self.mail_smtp_password.get()
        if not host or not sender:
            messagebox.showerror(
                "Bulk Mailer", "Enter the SMTP host and the From email address.",
            )
            return None
        if not pwd:
            messagebox.showerror(
                "Bulk Mailer",
                "Enter the SMTP password (an app password for Gmail/O365).",
            )
            return None
        if self.mail_smtp_remember.get():
            try:
                mailer_client.save_smtp_password(sender, pwd)
            except Exception:  # noqa: BLE001
                log.warning("could not save SMTP password to keyring")
        try:
            port = int(self.mail_smtp_port.get())
        except (tk.TclError, ValueError):
            port = 587
        return mailer_client.SMTPSettings(
            host=host, port=port, sender=sender, password=pwd, use_starttls=True,
        )

    # ------------------------------------------------------------------
    # Pick / preview
    # ------------------------------------------------------------------

    def _mail_pick_mapping(self) -> None:
        f = filedialog.askopenfilename(
            title="Pick the mapping Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if f:
            self.mail_mapping_path.set(f)

    def _mail_pick_attach(self) -> None:
        d = filedialog.askdirectory(title="Pick the attachments folder")
        if d:
            self.mail_attach_dir.set(d)

    # ----- Split & Send by email column -------------------------------

    def _mail_split_pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Pick the main Excel (rows carry an email column)",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if not path:
            return
        self.mail_split_input.set(path)
        if not self.mail_split_outdir.get().strip():
            self.mail_split_outdir.set(str(Path(path).parent / "split_by_email"))
        try:
            sheets = mailer_split.list_sheet_names(path)
        except Exception as exc:  # noqa: BLE001 — bad/locked workbook
            messagebox.showerror("Split & Send", f"Couldn't read workbook: {exc}")
            return
        self.mail_split_sheet_cb.configure(values=sheets)
        if sheets:
            self.mail_split_sheet.set(sheets[0])
        self._mail_split_fill_columns()

    def _mail_split_fill_columns(self) -> None:
        """Populate the email-column picker from the chosen sheet's header row."""
        path = self.mail_split_input.get().strip()
        if not path or not Path(path).is_file():
            return
        try:
            headers = [h for h in mailer_split.read_headers(
                path, self.mail_split_sheet.get().strip() or None) if h]
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Split & Send", f"Couldn't read headers: {exc}")
            return
        self.mail_split_column_cb.configure(values=headers)
        # Preselect the most email-ish header so most runs are two clicks.
        guess = next((h for h in headers if "email" in h.lower() or "mail" in h.lower()), "")
        self.mail_split_column.set(guess or (headers[0] if headers else ""))

    def _mail_split_pick_outdir(self) -> None:
        chosen = filedialog.askdirectory(title="Folder for the per-recipient files")
        if chosen:
            self.mail_split_outdir.set(chosen)

    def _mail_split_run(self, *, load_into_mailer: bool) -> None:
        """Split the main sheet per email address; optionally arm the mailer."""
        path = self.mail_split_input.get().strip()
        column = self.mail_split_column.get().strip()
        out_dir = self.mail_split_outdir.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showerror("Split & Send", "Pick the main Excel first.")
            return
        if not column:
            messagebox.showerror("Split & Send", "Pick the email column.")
            return
        if not out_dir:
            messagebox.showerror("Split & Send", "Pick the split-files folder.")
            return
        try:
            result = mailer_split.split_by_email(
                path, out_dir, email_column=column,
                sheet_name=self.mail_split_sheet.get().strip() or None)
        except ValueError as exc:            # missing column — user-fixable
            messagebox.showerror("Split & Send", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            log.exception("split_by_email failed")
            messagebox.showerror("Split & Send", f"{type(exc).__name__}: {exc}")
            return
        if result.warnings:
            messagebox.showwarning("Split & Send", "\n".join(result.warnings))
        if not result.groups:
            self.mail_status.configure(text="Split produced no recipients.")
            return

        if not load_into_mailer:
            self.mail_status.configure(
                text=f"Split {Path(path).name}: {len(result.groups)} recipient file(s) "
                     f"in {out_dir}")
            try:
                import os
                os.startfile(out_dir)  # noqa: S606 — user asked for the folder
            except Exception:  # noqa: BLE001
                pass
            return

        rows = mailer_split.build_mail_rows(
            result, cc=self.mail_split_cc.get(), bcc=self.mail_split_bcc.get())
        self._mail_rows = rows
        self._mail_campaign = Path(path).name        # sent-log key for skip-sent
        for child in self.mail_tree.get_children():
            self.mail_tree.delete(child)
        valid = 0
        for r in rows:
            status = "OK" if r.is_valid else "; ".join(r.issues)
            self.mail_tree.insert(
                "", "end",
                values=(r.row_index, r.email, r.name,
                        ", ".join(p.name for p in r.attachments), r.cc, r.bcc, status),
                tags=("ok",) if r.is_valid else ("bad",))
            valid += 1 if r.is_valid else 0
        self.mail_status.configure(
            text=f"Split-loaded {valid} recipient(s) from {Path(path).name} — "
                 "set Subject/Body, then Test or Run.")
        ready = "normal" if valid else "disabled"
        self.btn_mail_run.configure(state=ready)
        self.btn_mail_test.configure(state=ready)

    def _mail_preview(self) -> None:
        mapping = self.mail_mapping_path.get().strip()
        attach = self.mail_attach_dir.get().strip()
        if not mapping or not Path(mapping).is_file():
            messagebox.showerror("Bulk Mailer", "Pick a valid mapping Excel first.")
            return
        if not attach or not Path(attach).is_dir():
            messagebox.showerror("Bulk Mailer", "Pick a valid attachments folder first.")
            return
        try:
            rows, warnings = mailer_io.read_mapping(mapping, attach)
        except Exception as exc:  # noqa: BLE001
            log.exception("mapping read failed")
            messagebox.showerror("Bulk Mailer", f"Couldn't read mapping: {exc}")
            return
        self._mail_rows = rows
        self._mail_campaign = ""            # mapping flow: campaign = mapping filename
        for child in self.mail_tree.get_children():
            self.mail_tree.delete(child)
        valid = 0
        for r in rows:
            status = "OK" if r.is_valid else "; ".join(r.issues)
            files = ", ".join(p.name for p in r.attachments) or "(none)"
            self.mail_tree.insert(
                "", "end",
                values=(r.row_index, r.email, r.name, files, r.cc, r.bcc, status),
                tags=("ok",) if r.is_valid else ("bad",),
            )
            if r.is_valid:
                valid += 1
        msg = f"{valid} valid / {len(rows)} rows"
        if warnings:
            msg += "  ·  " + "; ".join(warnings)
        self.mail_status.configure(text=msg)
        ready = "normal" if valid else "disabled"
        self.btn_mail_run.configure(state=ready)
        self.btn_mail_test.configure(state=ready)
        if warnings:
            messagebox.showwarning("Bulk Mailer", "\n".join(warnings))

    # ------------------------------------------------------------------
    # Test / run / stop
    # ------------------------------------------------------------------

    def _mail_test(self) -> None:
        """Send/draft the FIRST valid row to an address you type."""
        me = simpledialog.askstring(
            "Send test", "Send the test email to which address?",
        )
        if not me:
            return
        first = next((r for r in self._mail_rows if r.is_valid), None)
        if first is None:
            messagebox.showinfo("Bulk Mailer", "No valid row to test with.")
            return
        subject = self.mail_subject.get().strip() or "(no subject)"
        body_tmpl = self.mail_body_text.get("1.0", "end").rstrip("\n")
        body, _missing = mailer_io.render_template(body_tmpl, first.fields)
        send = self.mail_mode.get() == "send"
        transport = self.mail_transport.get()
        email = mailer_client.OutgoingEmail(
            to=me, subject=f"[TEST] {subject}", body=body,
            attachments=first.attachments,
        )

        smtp_settings = None
        draft_dir = Path(self.mail_attach_dir.get() or str(Path.home())) / "_mail_drafts"
        if transport == "smtp":
            smtp_settings = self._mail_smtp_settings()
            if smtp_settings is None:
                return
        if transport == "graph" and self._graph_session is None:
            messagebox.showerror(
                "Bulk Mailer", "Sign in to Microsoft 365 first (Send via → Sign in).",
            )
            return
        graph_sess = self._graph_session
        self.mail_status.configure(text=f"Testing to {me}…")

        def worker() -> None:
            try:
                if transport == "outlook":
                    acct = self.mail_outlook_account.get().strip()
                    with mailer_client.OutlookSession() as ol:
                        outcome = ol.create(email, send=send, from_account=acct)
                elif transport == "graph":
                    outcome = graph_sess.send(email) if send else graph_sess.draft(email)
                else:
                    with mailer_client.SMTPMailer(smtp_settings) as sm:
                        outcome = sm.send(email) if send else sm.draft(email, draft_dir)
                self._post(MSG_MAIL_DONE, {
                    "test": True, "outcome_status": outcome.status,
                    "error": outcome.error, "to": me, "entry": outcome.entry_id,
                })
            except Exception as exc:  # noqa: BLE001
                log.exception("test mail failed")
                self._post(MSG_MAIL_ERROR, f"{type(exc).__name__}: {exc}")

        threading.Thread(target=worker, daemon=True).start()

    def _mail_stop(self) -> None:
        self._mail_stop_flag.set()
        self.mail_status.configure(text="Stopping…")

    def _mail_run(self) -> None:
        if self._mail_worker and self._mail_worker.is_alive():
            messagebox.showinfo("Bulk Mailer", "A run is already in progress.")
            return
        valid_rows = [r for r in self._mail_rows if r.is_valid]
        if not valid_rows:
            messagebox.showerror("Bulk Mailer", "Load a mapping and fix invalid rows first.")
            return
        subject = self.mail_subject.get().strip()
        if not subject:
            messagebox.showerror("Bulk Mailer", "Enter a subject.")
            return
        mode = self.mail_mode.get()
        send = mode == "send"
        transport = self.mail_transport.get()
        body_tmpl = self.mail_body_text.get("1.0", "end").rstrip("\n")
        # Split-loaded runs log under the main sheet's name; mapping runs under
        # the mapping file's name — so skip-already-sent tracks the right campaign.
        campaign = self._mail_campaign or Path(self.mail_mapping_path.get()).name
        skip_sent = self.mail_skip_sent.get()
        delay = max(0.0, float(self.mail_delay_s.get()))

        smtp_settings = None
        draft_dir = Path(self.mail_attach_dir.get() or str(Path.home())) / "_mail_drafts"
        if transport == "smtp":
            smtp_settings = self._mail_smtp_settings()
            if smtp_settings is None:
                return
        if transport == "graph" and self._graph_session is None:
            messagebox.showerror(
                "Bulk Mailer", "Sign in to Microsoft 365 first (Send via → Sign in).",
            )
            return
        graph_sess = self._graph_session

        if send:
            if transport == "smtp":
                where = f"from {smtp_settings.sender}"
            elif transport == "graph":
                where = f"from {graph_sess.account}"
            else:
                where = "via Outlook"
            tail = f" — these will actually be sent {where}."
        else:
            if transport == "smtp":
                tail = f" as .eml drafts in {draft_dir}."
            elif transport == "graph":
                tail = f" as drafts in the {graph_sess.account} mailbox."
            else:
                tail = " as drafts in Outlook."
        if not messagebox.askyesno(
            "Bulk Mailer — confirm",
            f"About to {'SEND' if send else 'create drafts for'} "
            f"{len(valid_rows)} email(s){tail}\n\nSubject: {subject}\n\nContinue?",
        ):
            return

        self._mail_stop_flag.clear()
        self.btn_mail_run.configure(state="disabled")
        self.btn_mail_test.configure(state="disabled")
        self.btn_mail_preview.configure(state="disabled")
        self.btn_mail_stop.configure(state="normal")
        self.mail_progress.configure(value=0, maximum=len(valid_rows))

        iid_by_row = {}
        for iid in self.mail_tree.get_children():
            vals = self.mail_tree.item(iid, "values")
            iid_by_row[int(vals[0])] = iid

        def _process(make_one) -> dict:
            import time as _t
            counts = {"DRAFTED": 0, "SENT": 0, "FAILED": 0, "SKIPPED": 0}
            for i, r in enumerate(valid_rows, start=1):
                if self._mail_stop_flag.is_set():
                    break
                if skip_sent and self._mail_log.already_sent(campaign, r.email, subject):
                    counts["SKIPPED"] += 1
                    self._post(MSG_MAIL_PROGRESS,
                               (i, len(valid_rows), r.email, "SKIPPED", r.row_index))
                    continue
                body, _missing = mailer_io.render_template(body_tmpl, r.fields)
                email = mailer_client.OutgoingEmail(
                    to=r.email, subject=subject, body=body,
                    attachments=r.attachments, cc=r.cc, bcc=r.bcc,
                )
                outcome = make_one(email)
                counts[outcome.status] = counts.get(outcome.status, 0) + 1
                if outcome.status == "SENT":
                    self._mail_log.record(campaign, r.email, subject, "SENT")
                elif outcome.status == "FAILED":
                    self._mail_log.record(campaign, r.email, subject, "FAILED", outcome.error)
                self._post(
                    MSG_MAIL_PROGRESS,
                    (i, len(valid_rows), r.email,
                     outcome.status + (f": {outcome.error}" if outcome.error else ""),
                     r.row_index),
                )
                if delay > 0 and i < len(valid_rows):
                    _t.sleep(delay)
            return counts

        def worker() -> None:
            try:
                if transport == "outlook":
                    acct = self.mail_outlook_account.get().strip()
                    with mailer_client.OutlookSession() as ol:
                        log.info("Bulk Mailer via Outlook from %s",
                                 acct or ol.verify_account() or "?")
                        counts = _process(
                            lambda e: ol.create(e, send=send, from_account=acct)
                        )
                elif transport == "graph":
                    log.info("Bulk Mailer via Graph as %s", graph_sess.account)
                    counts = _process(
                        (lambda e: graph_sess.send(e)) if send
                        else (lambda e: graph_sess.draft(e))
                    )
                else:
                    with mailer_client.SMTPMailer(smtp_settings) as sm:
                        log.info("Bulk Mailer via SMTP %s as %s",
                                 smtp_settings.host, smtp_settings.sender)
                        counts = _process(
                            (lambda e: sm.send(e)) if send
                            else (lambda e: sm.draft(e, draft_dir))
                        )
                self._post(MSG_MAIL_DONE, {
                    "test": False, "counts": counts, "mode": mode,
                    "draft_dir": str(draft_dir) if transport == "smtp" and not send else "",
                })
            except (mailer_client.OutlookUnavailableError,
                    mailer_client.SMTPConfigError,
                    mailer_client.SMTPAuthError) as exc:
                self._post(MSG_MAIL_ERROR, str(exc))
            except Exception as exc:  # noqa: BLE001
                log.exception("bulk mail run failed")
                self._post(MSG_MAIL_ERROR, f"{type(exc).__name__}: {exc}")

        self._mail_iid_by_row = iid_by_row
        self._mail_worker = threading.Thread(target=worker, daemon=True)
        self._mail_worker.start()

    # ------------------------------------------------------------------
    # Queue plumbing + message handling
    # ------------------------------------------------------------------

    def _post(self, kind: str, payload: object) -> None:
        self._msg_queue.put((kind, payload))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._msg_queue.get_nowait()
                # A raising handler must never kill the pump (would freeze all
                # background messaging for the rest of the session).
                try:
                    self._handle_msg(kind, payload)
                except Exception:  # noqa: BLE001
                    log.exception("message handler failed for %r", kind)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, kind: str, payload: object) -> None:
        # WhatsApp-blast + Health messages are handled by shared mixins.
        if kind.startswith("wa_") and self._wa_handle_msg(kind, payload):
            return
        if kind.startswith("health_") and self._health_handle_msg(kind, payload):
            return
        if kind == MSG_MAIL_PROGRESS:
            idx, total, to, status, row_index = payload  # type: ignore[misc]
            self.mail_progress.configure(maximum=total, value=idx)
            self.mail_status.configure(text=f"{idx}/{total}  ·  {to}  [{status}]")
            iid = getattr(self, "_mail_iid_by_row", {}).get(row_index)
            if iid is not None:
                vals = list(self.mail_tree.item(iid, "values"))
                vals[6] = status
                tag = "ok" if status in ("DRAFTED", "SENT", "SKIPPED") else "bad"
                self.mail_tree.item(iid, values=vals, tags=(tag,))
        elif kind == MSG_MAIL_DONE:
            info = payload  # type: ignore[assignment]
            self.btn_mail_run.configure(state="normal")
            self.btn_mail_test.configure(state="normal")
            self.btn_mail_preview.configure(state="normal")
            self.btn_mail_stop.configure(state="disabled")
            if info.get("test"):
                st = info.get("outcome_status", "?")
                err = info.get("error", "")
                self.mail_status.configure(text=f"Test {st} → {info.get('to','')}")
                if st == "FAILED":
                    messagebox.showerror("Bulk Mailer — Test", err or "Test failed.")
                else:
                    messagebox.showinfo(
                        "Bulk Mailer — Test",
                        f"Test {st.lower()} to {info.get('to','')}.\n\n"
                        + ("Check your Outlook Drafts." if st == "DRAFTED"
                           else "Check your inbox."),
                    )
            else:
                c = info.get("counts", {})
                verb = "drafted" if info.get("mode") == "draft" else "sent"
                summary = (
                    f"{c.get('DRAFTED',0)} drafted · {c.get('SENT',0)} sent · "
                    f"{c.get('SKIPPED',0)} skipped · {c.get('FAILED',0)} failed"
                )
                self.mail_status.configure(text=summary)
                if info.get("mode") == "draft":
                    dd = info.get("draft_dir")
                    if dd:
                        review = (f"\n\n.eml drafts written to:\n{dd}\n"
                                  "Open any to review; double-click to send.")
                    else:
                        review = ("\n\nReview them in your Drafts folder "
                                  "(Outlook desktop or Outlook web), then send.")
                else:
                    review = ""
                messagebox.showinfo(
                    "Bulk Mailer — Done", f"Run complete ({verb}).\n\n{summary}{review}",
                )
        elif kind == MSG_MAIL_ERROR:
            self.btn_mail_run.configure(state="normal")
            self.btn_mail_test.configure(state="normal")
            self.btn_mail_preview.configure(state="normal")
            self.btn_mail_stop.configure(state="disabled")
            self.mail_status.configure(text=f"⚠ {payload}")
            messagebox.showerror("Bulk Mailer — Error", str(payload))
        elif kind == "mail_graph_prompt":
            msg, code = payload  # type: ignore[misc]
            self.mail_graph_status.set(f"Code: {code}  →  microsoft.com/devicelogin")
            self.btn_mail_graph_signin.configure(state="normal")
            import webbrowser
            webbrowser.open("https://microsoft.com/devicelogin")
            messagebox.showinfo(
                "Microsoft 365 sign-in",
                f"{msg}\n\nThe code is also shown on screen "
                f"({code}) in case you need it again.",
            )
        elif kind == "mail_graph_signed_in":
            self._graph_session = payload
            acct = getattr(payload, "account", "")
            self.mail_graph_status.set(f"● Signed in: {acct}")
            self.btn_mail_graph_signin.configure(state="normal")
            messagebox.showinfo("Microsoft 365", f"Signed in as {acct}.")
        elif kind == "mail_graph_signin_failed":
            self.btn_mail_graph_signin.configure(state="normal")
            if self._graph_session is None:
                self.mail_graph_status.set("Sign-in failed — click Sign in to retry")
                messagebox.showerror("Microsoft 365 sign-in", str(payload))

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def _on_close(self) -> None:
        if self._mail_worker and self._mail_worker.is_alive():
            if not messagebox.askyesno(
                "Bulk Mailer", "A run is in progress. Stop it and quit?",
            ):
                return
            self._mail_stop_flag.set()
        # Stop any WhatsApp blast + close the browser session.
        self._wa_shutdown()
        self.root.destroy()


def run() -> None:
    config.LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    from logging.handlers import RotatingFileHandler
    handler = RotatingFileHandler(
        str(config.LOG_FILE), maxBytes=5_000_000, backupCount=2, encoding="utf-8",
    )
    logging.basicConfig(
        level=logging.INFO, handlers=[handler],
        format="%(asctime)s %(levelname)s %(name)s — %(message)s",
    )
    root = tk.Tk()
    try:
        from tkinter import font as tkfont
        tkfont.nametofont("TkDefaultFont").configure(family="Segoe UI", size=10)
    except Exception:  # noqa: BLE001
        pass
    MailerApp(root)
    root.mainloop()

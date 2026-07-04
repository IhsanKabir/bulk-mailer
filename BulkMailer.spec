# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Bulk Mailer.

A lightweight single-window app — no bundled browser or ML model, so the
.exe is small (~30-40 MB). Use: pyinstaller --noconfirm BulkMailer.spec
"""

from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    "openpyxl",
    "openpyxl.cell._writer",
    "tkinter",
    # Windows Credential Manager backend for `keyring` (SMTP password store).
    "keyring.backends.Windows",
    # Outlook COM automation — lazily imported in mailer_client, so the
    # static analyzer can't see it; the bundle must carry it.
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    # Microsoft Graph device-code auth (graph_mailer), lazily imported.
    "msal",
    "requests",
]

# Pull in keyring backends + sv_ttk's Tcl theme resources (collect_all gets
# the .tcl files PyInstaller's analyzer misses, so set_theme() works frozen).
# Playwright ships a node "driver" as package data the analyzer can't see —
# collect_all bundles it so the WhatsApp Blast can drive the system browser.
for pkg in ["keyring", "sv_ttk", "playwright"]:
    pkg_datas, pkg_bins, pkg_hi = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_bins
    hiddenimports += pkg_hi


a = Analysis(
    ["run_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="BulkMailer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

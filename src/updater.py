"""Lightweight update check for the standalone Bulk Mailer.

Unlike the combined console (which downloads + swaps the exe in place), the
standalone just *tells* you a newer release exists and opens its download page —
no self-replacement, so nothing here can corrupt a running exe. Every failure is
swallowed and reported as "couldn't check" so a flaky network never breaks the
button.
"""

from __future__ import annotations

import json
import logging
import re
import urllib.request
import webbrowser
from dataclasses import dataclass

from . import __version__

log = logging.getLogger(__name__)

GITHUB_REPO = "IhsanKabir/bulk-mailer"
GITHUB_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"
USER_AGENT = f"BulkMailer-Updater/{__version__}"

_VERSION_RE = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)")


@dataclass(frozen=True)
class UpdateInfo:
    latest_version: str
    is_newer: bool
    notes: str
    download_url: str        # a .exe asset if present, else the releases page
    page_url: str


def _parse_version(text: str) -> tuple[int, int, int]:
    m = _VERSION_RE.match((text or "").strip())
    if not m:
        return (0, 0, 0)
    return tuple(int(x) for x in m.groups())  # type: ignore[return-value]


def _is_newer(latest: str, current: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def check_for_update(timeout: int = 15) -> UpdateInfo | None:
    """Query GitHub for the latest release. Returns UpdateInfo, or None if the
    check itself failed (network/proxy/rate-limit/malformed) — never raises."""
    req = urllib.request.Request(
        GITHUB_API, headers={"Accept": "application/vnd.github+json",
                             "User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:  # noqa: BLE001 — any failure => "couldn't check"
        log.info("Update check failed: %s", exc)
        return None
    tag = str(data.get("tag_name") or "").strip()
    if not tag:
        return None
    exe = next((a.get("browser_download_url", "")
                for a in data.get("assets", [])
                if str(a.get("name", "")).lower().endswith(".exe")), "")
    return UpdateInfo(
        latest_version=tag.lstrip("v"),
        is_newer=_is_newer(tag, __version__),
        notes=str(data.get("body") or "").strip(),
        download_url=exe or RELEASES_PAGE,
        page_url=RELEASES_PAGE,
    )


def open_download(info: UpdateInfo) -> None:
    """Open the release/download page in the default browser."""
    webbrowser.open(info.download_url or info.page_url)

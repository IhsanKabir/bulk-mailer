"""Tests for the standalone update check (version compare + parse, no network)."""

from __future__ import annotations

from src import updater
from src.updater import UpdateInfo, _is_newer, _parse_version


def test_parse_version_handles_v_prefix_and_junk():
    assert _parse_version("v1.4.0") == (1, 4, 0)
    assert _parse_version("1.10.2") == (1, 10, 2)
    assert _parse_version("garbage") == (0, 0, 0)


def test_is_newer_semver_ordering():
    assert _is_newer("1.5.0", "1.4.0")
    assert _is_newer("1.10.0", "1.9.9")      # numeric, not lexical
    assert not _is_newer("1.4.0", "1.4.0")
    assert not _is_newer("1.3.9", "1.4.0")


def test_check_for_update_parses_release(monkeypatch):
    import io
    payload = (
        '{"tag_name":"v9.9.9","body":"- new stuff",'
        '"assets":[{"name":"BulkMailer.exe",'
        '"browser_download_url":"https://x/BulkMailer.exe"}]}'
    )

    class _Resp(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(updater.urllib.request, "urlopen",
                        lambda *a, **k: _Resp(payload.encode()))
    info = updater.check_for_update()
    assert isinstance(info, UpdateInfo)
    assert info.latest_version == "9.9.9" and info.is_newer
    assert info.download_url.endswith("BulkMailer.exe")


def test_check_for_update_returns_none_on_network_error(monkeypatch):
    def boom(*a, **k):
        raise OSError("no network")

    monkeypatch.setattr(updater.urllib.request, "urlopen", boom)
    assert updater.check_for_update() is None      # never raises

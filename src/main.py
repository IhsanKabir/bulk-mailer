"""Entry point.

No bundled browser or model — this is a lightweight single-window Tkinter
app. PyInstaller packs it into a small .exe (see build_exe.bat). No Python
install, no admin access needed on the target machine.
"""

from __future__ import annotations

import sys

from . import config


def main() -> int:
    config.APP_DIR.mkdir(parents=True, exist_ok=True)
    from .gui import run
    run()
    return 0


if __name__ == "__main__":
    sys.exit(main())

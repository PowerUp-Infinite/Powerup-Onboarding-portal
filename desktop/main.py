"""
main.py — PowerUp Portal desktop app entry point.

Usage (dev):
    cd desktop
    python main.py

Usage (bundled):
    double-click the PowerUp-Portal.exe (Windows) or
    open the PowerUp-Portal.app (macOS).
"""
from __future__ import annotations

import sys
import traceback


def _die(title: str, msg: str) -> None:
    """Show a friendly error popup and exit. Used when the app can't even
    construct itself (e.g. credentials.json missing from the bundle)."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, msg)
    except Exception:
        print(f"{title}\n\n{msg}", file=sys.stderr)
    sys.exit(1)


def main() -> None:
    # app_config.py loads env vars and bootstraps portal/ onto sys.path.
    # If credentials.json is missing, this will throw during module init —
    # catch it so the user sees a clear error instead of a silent crash.
    try:
        import app_config  # noqa: F401
    except Exception as e:
        _die(
            "PowerUp Portal — configuration error",
            f"The app could not load its configuration:\n\n{e}\n\n"
            f"If you built this app from source, make sure portal/.env "
            f"contains all required keys and that "
            f"desktop/resources/credentials.json exists.",
        )
        return

    try:
        import gui
        gui.run()
    except Exception as e:
        tb = traceback.format_exc()
        _die(
            "PowerUp Portal — crashed",
            f"The app crashed with:\n\n{type(e).__name__}: {e}\n\n"
            f"Traceback:\n{tb}",
        )


if __name__ == "__main__":
    main()

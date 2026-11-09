"""
eFBR CPR Receipt Downloader — Entry Point
==========================================
Run this file:  python main.py

On start-up the app checks for everything it needs and downloads whatever
is missing by itself (Python packages + Playwright's Chromium browser).
Nothing needs to be installed manually except Python itself.

Options:
    python main.py --setup-only    download/verify dependencies, then exit
"""

import sys


def main():
    # bootstrap uses only the standard library, so it can run before any
    # third-party package has been installed.
    from bootstrap import ensure_ready

    setup_only = "--setup-only" in sys.argv
    ensure_ready(show_gui=not setup_only)

    if setup_only:
        print("Setup complete. All dependencies are installed.")
        return

    from gui import CPRDownloaderApp
    app = CPRDownloaderApp()
    app.mainloop()


if __name__ == "__main__":
    main()

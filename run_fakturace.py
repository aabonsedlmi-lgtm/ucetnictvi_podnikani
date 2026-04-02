#!/usr/bin/env python3
from __future__ import annotations

import os
import threading
import webbrowser

import invoice_manager_gui


def open_browser_later() -> None:
    public_url = os.environ.get("APP_PUBLIC_URL", "").strip()
    if public_url:
        webbrowser.open(public_url)
        return
    host = os.environ.get("APP_HOST", "127.0.0.1").strip() or "127.0.0.1"
    port = os.environ.get("APP_PORT", "8000").strip() or "8000"
    webbrowser.open(f"http://{host}:{port}")


def main() -> None:
    timer = threading.Timer(2.0, open_browser_later)
    timer.daemon = True
    timer.start()
    invoice_manager_gui.main()


if __name__ == "__main__":
    main()

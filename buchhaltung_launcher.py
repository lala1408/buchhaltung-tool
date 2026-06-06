# -*- coding: utf-8 -*-
from __future__ import annotations

import socket
import threading
import webbrowser
import sys
from ctypes import windll
from http.server import ThreadingHTTPServer

from buchhaltung_app import AccountingRequestHandler, ensure_dirs


def find_free_port(start: int = 8501) -> int:
    for port in range(start, start + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("Kein freier lokaler Port gefunden.")


class BuchhaltungLauncher:
    def __init__(self) -> None:
        ensure_dirs()
        self.port = find_free_port()
        self.url = f"http://127.0.0.1:{self.port}"
        self.server = ThreadingHTTPServer(("127.0.0.1", self.port), AccountingRequestHandler)
        self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    def open_browser(self) -> None:
        webbrowser.open(self.url)

    def stop(self) -> None:
        self.server.shutdown()
        self.server.server_close()

    def show_control_message(self) -> None:
        message = (
            "Die lokale Buchhaltungs-App laeuft.\n\n"
            f"Browser-Adresse:\n{self.url}\n\n"
            "Dieses Fenster offen lassen, solange du arbeitest.\n"
            "OK beendet die App."
        )
        if sys.platform.startswith("win"):
            windll.user32.MessageBoxW(None, message, "SFK Buchhaltung", 0x40)
        else:
            print(message)
            input("\nEnter beendet die App.")

    def run(self) -> None:
        self.server_thread.start()
        self.open_browser()
        try:
            self.show_control_message()
        finally:
            self.stop()


def main() -> None:
    BuchhaltungLauncher().run()


if __name__ == "__main__":
    main()

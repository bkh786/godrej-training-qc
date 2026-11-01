#!/usr/bin/env python3
"""
Local Development Server for Godrej Dashboard.
Serves static HTML/CSS/JS files AND routes POST /api/chat to api/chat.py.
Usage:
    python3 dev_server.py [port]
"""

import os
import sys
from http.server import HTTPServer, SimpleHTTPRequestHandler

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT_DIR)

from api.chat import handler as ChatHandler  # noqa: E402


class GodrejDashboardDevHandler(SimpleHTTPRequestHandler):
    def do_OPTIONS(self):
        if self.path.startswith("/api/chat"):
            ChatHandler.do_OPTIONS(self)
        else:
            super().do_OPTIONS()

    def do_POST(self):
        if self.path.startswith("/api/chat"):
            ChatHandler._send_json = ChatHandler._send_json.__get__(self, self.__class__)
            ChatHandler.do_POST(self)
        else:
            self.send_error(404, "Endpoint not found")


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8090
    server = HTTPServer(("127.0.0.1", port), GodrejDashboardDevHandler)
    print(f"Godrej Dashboard + Live SharePoint AI Copilot running at http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        server.server_close()


import json
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from pathlib import Path

import analyzer

PORT = 8765


class Handler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)

        if parsed.path == "/run":
            params = parse_qs(parsed.query)
            price  = float(params.get("price", ["2.00"])[0])
            cost   = float(params.get("cost",  ["0.80"])[0])
            try:
                analyzer.run_analysis(price=price, cost=cost)
                self._json({"ok": True})
            except Exception as e:
                self._json({"ok": False, "error": str(e)})

        elif parsed.path == "/data":
            path = Path("output/dashboard_data.json")
            if path.exists():
                self._json(json.loads(path.read_text(encoding="utf-8")))
            else:
                self._json({"error": "no data yet"})

        else:
            super().do_GET()

    def _json(self, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        if any(x in args[0] for x in [".png", ".json", ".css"]):
            return
        super().log_message(fmt, *args)


if __name__ == "__main__":
    server = HTTPServer(("localhost", PORT), Handler)
    print(f"server running at http://localhost:{PORT}/dashboard.html")
    print("press Ctrl+C to stop")
    threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{PORT}/dashboard.html")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nserver stopped.")

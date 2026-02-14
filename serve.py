"""Serve the KPI dashboard frontend on http://localhost:8080"""
import http.server
import webbrowser
import threading
import os
from pathlib import Path

PORT = 8080
DIR = Path(__file__).resolve().parent
os.chdir(DIR)

def open_browser():
    import time
    time.sleep(0.8)
    webbrowser.open(f"http://localhost:{PORT}")

if __name__ == "__main__":
    server = http.server.HTTPServer(("", PORT), http.server.SimpleHTTPRequestHandler)
    print(f"Serving at http://localhost:{PORT}")
    print("Open the link in your browser. Press Ctrl+C to stop.")
    threading.Thread(target=open_browser, daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()

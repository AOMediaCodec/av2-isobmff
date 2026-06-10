"""Live preview HTTP server for spec development.

Serves the built HTML output on localhost and pairs with ``--watch`` to
provide instant visual feedback on source changes.
"""

from __future__ import annotations

import logging
import socket
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path


def find_free_port(preferred: int = 8080) -> int:
    """Return *preferred* if available, otherwise the next free port."""
    for port in range(preferred, preferred + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError(
        f"No free port found in range {preferred}–{preferred + 99}. "
        "Free up a port or configure a different starting port."
    )


def start_server(output_dir: Path, port: int = 8080) -> tuple[HTTPServer, str]:
    """Start an HTTP server serving *output_dir* in a daemon thread.

    Args:
        output_dir: Directory containing the built HTML to serve.
        port:       Preferred port number; a free port is found if taken.

    Returns:
        ``(httpd, url)`` — the HTTPServer instance and the root URL.
    """
    port = find_free_port(port)

    class _Handler(SimpleHTTPRequestHandler):
        def __init__(self, *a, **kw):
            super().__init__(*a, directory=str(output_dir), **kw)

        def log_message(self, fmt, *args):
            logging.debug("HTTP: " + fmt % args)

    httpd = HTTPServer(("127.0.0.1", port), _Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()

    url = f"http://localhost:{port}"
    return httpd, url


def open_browser(url: str) -> None:
    """Open *url* in the default browser (best-effort, no error on failure)."""
    try:
        import webbrowser

        webbrowser.open(url)
    except Exception:
        pass


def find_index_html(output_dir: Path) -> Path | None:
    """Return the main HTML file in *output_dir*, or None."""
    for candidate in ("index.html", "index.htm"):
        p = output_dir / candidate
        if p.exists():
            return p
    html_files = sorted(output_dir.glob("*.html"))
    return html_files[0] if html_files else None

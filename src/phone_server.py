"""
Tiny local HTTP server for receiving invoice photos from a mobile browser.

Usage:
    server = PhoneUploadServer(on_images_received=my_callback)
    server.start()   # non-blocking — runs in a daemon thread
    print(server.url)
    ...
    server.stop()

The callback is called from the server thread with a list of temp file paths.
The caller is responsible for deleting those files after use.
"""
from __future__ import annotations

import os
import socket
import tempfile
import threading
from typing import Callable, Optional


# ── Mobile-friendly upload page ───────────────────────────────────────────────

_UPLOAD_PAGE_TEMPLATE = """\
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Invoice Scan</title>
  <style>
    *{{box-sizing:border-box}}
    body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
         max-width:440px;margin:0 auto;padding:28px 20px;background:#f9fafb}}
    h2{{font-size:22px;color:#111827;margin:0 0 6px}}
    .sup{{display:inline-block;margin:0 0 16px;background:#dbeafe;color:#1e40af;
          font-size:14px;font-weight:600;padding:4px 12px;border-radius:20px}}
    p{{color:#6b7280;font-size:15px;margin:0 0 20px}}
    .drop{{display:block;background:#fff;border:2px dashed #d1d5db;
          border-radius:12px;padding:36px 20px;text-align:center;
          cursor:pointer;color:#6b7280;font-size:16px;transition:border-color .15s}}
    .drop:active{{border-color:#2563eb}}
    input[type=file]{{display:none}}
    .ct{{text-align:center;margin:10px 0;min-height:22px;font-size:16px;
        color:#059669;font-weight:600}}
    .btn{{display:block;width:100%;background:#2563eb;color:#fff;border:none;
         padding:16px;font-size:17px;font-weight:600;border-radius:10px;
         cursor:pointer;margin-top:8px}}
    .btn:disabled{{background:#93c5fd;cursor:not-allowed}}
    .btn:not(:disabled):active{{background:#1d4ed8}}
  </style>
</head>
<body>
  <h2>📷 Invoice Scan</h2>
  <div class="sup">Supplier: {supplier_name}</div>
  <p>Select one or more invoice photos, then tap <strong>Upload</strong>.<br>
     You can upload in multiple batches for multi-page invoices.</p>
  <form id="f" method="POST" enctype="multipart/form-data">
    <label class="drop" for="p">📁 Tap here to choose photos</label>
    <input type="file" id="p" name="photos" accept="image/*"
           multiple onchange="upd(this)">
    <div class="ct" id="ct"></div>
    <button class="btn" id="btn" type="submit" disabled>Upload Photos</button>
  </form>
  <script>
    function upd(i){{
      var n=i.files.length;
      document.getElementById("ct").textContent=
        n ? n+" photo"+(n>1?"s":"")+" selected" : "";
      document.getElementById("btn").disabled=!n;
    }}
  </script>
</body>
</html>
"""

_SUCCESS_PAGE = """\
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Uploaded</title>
  <style>
    body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
         max-width:440px;margin:0 auto;padding:60px 20px;text-align:center}
    .ic{font-size:60px;margin-bottom:8px}
    h2{color:#059669;font-size:24px;margin:0 0 8px}
    p{color:#6b7280;font-size:16px;margin:0}
    a{display:inline-block;margin-top:24px;background:#f3f4f6;color:#374151;
      text-decoration:none;padding:12px 28px;border-radius:8px;font-size:15px}
  </style>
</head>
<body>
  <div class="ic">✅</div>
  <h2>Photos received!</h2>
  <p>Return to the app and click <strong>Done</strong><br>when all pages are uploaded.</p>
  <a href="/">Upload more pages</a>
</body>
</html>
"""


# ── Server ────────────────────────────────────────────────────────────────────

class PhoneUploadServer:
    """
    HTTP server that serves a mobile-friendly upload page and accepts multipart
    image uploads. Each POST batch triggers on_images_received with a list of
    temporary file paths (caller is responsible for deleting them).
    """

    def __init__(
        self,
        on_images_received: Callable[[list[str]], None],
        supplier_name: str = "Unknown",
    ):
        self._callback = on_images_received
        self._supplier_name = supplier_name
        self._httpd: Optional[object] = None
        self._thread: Optional[threading.Thread] = None
        self.port: int = _find_free_port()
        self.local_ip: str = _get_local_ip()
        self.url: str = f"http://{self.local_ip}:{self.port}"

    def start(self) -> None:
        import http.server
        handler = _make_handler(self._callback, self._supplier_name)
        self._httpd = http.server.HTTPServer(("", self.port), handler)
        self._thread = threading.Thread(
            target=self._httpd.serve_forever,
            daemon=True,
            name="phone-upload-server",
        )
        self._thread.start()

    def stop(self) -> None:
        if self._httpd is not None:
            self._httpd.shutdown()
            self._httpd = None


# ── Helpers ───────────────────────────────────────────────────────────────────

def _find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("", 0))
        return s.getsockname()[1]


def _get_local_ip() -> str:
    """Return the machine's LAN IP (not loopback) by probing an external address."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"


def _guess_suffix(data: bytes) -> str:
    """Detect image type from magic bytes for a correct temp file extension."""
    if data[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return ".png"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return ".tif"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return ".gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    return ".jpg"  # safe default (covers HEIC and other formats)


def _make_handler(callback: Callable[[list[str]], None], supplier_name: str = "Unknown"):
    """Return a BaseHTTPRequestHandler class bound to the given callback."""
    import http.server
    upload_page = _UPLOAD_PAGE_TEMPLATE.format(supplier_name=supplier_name).encode()

    class _Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):  # noqa: N802
            body = upload_page
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self):  # noqa: N802
            saved: list[str] = []
            try:
                import cgi
                ctype, pdict = cgi.parse_header(
                    self.headers.get("Content-Type", "")
                )
                if ctype == "multipart/form-data":
                    boundary = pdict.get("boundary", "")
                    if isinstance(boundary, str):
                        boundary = boundary.encode("ascii")
                    pdict["boundary"] = boundary
                    pdict["CONTENT-LENGTH"] = int(
                        self.headers.get("Content-Length", 0)
                    )
                    fields = cgi.parse_multipart(self.rfile, pdict)
                    for data in fields.get("photos", []):
                        if isinstance(data, str):
                            data = data.encode("latin-1")
                        if data:
                            tmp = tempfile.NamedTemporaryFile(
                                suffix=_guess_suffix(data), delete=False
                            )
                            tmp.write(data)
                            tmp.close()
                            saved.append(tmp.name)
            except Exception:
                pass  # return the success page regardless to avoid confusing the user

            if saved:
                callback(saved)

            body = _SUCCESS_PAGE.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *_args):  # noqa: N802
            pass  # suppress console output

    return _Handler

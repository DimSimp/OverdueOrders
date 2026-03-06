"""
Dialog for scanning invoice pages with a phone camera.

Opens a local HTTP server, displays a QR code the user scans with their
phone, and collects uploaded photos. When the user clicks Done, the dialog
calls on_done(list_of_image_paths) with the received image temp-file paths.
The caller is responsible for processing and deleting those files.
"""
from __future__ import annotations

import os
import threading
from typing import Callable

import customtkinter as ctk

from src.phone_server import PhoneUploadServer


def _make_qr_image(url: str, size: int = 260) -> ctk.CTkImage | None:
    """Render url as a QR code CTkImage. Returns None if qrcode is not installed."""
    try:
        import qrcode
        from PIL import Image
    except ImportError:
        return None

    qr = qrcode.QRCode(version=1, box_size=7, border=3, error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url)
    qr.make(fit=True)
    pil_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    pil_img = pil_img.resize((size, size), Image.LANCZOS)
    return ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(size, size))


class PhoneScanDialog(ctk.CTkToplevel):
    """
    Modal that:
      1. Starts a local HTTP server on a random LAN port
      2. Shows a QR code pointing to the upload page
      3. Updates a photo count as uploads arrive
      4. Calls on_done(image_paths) when the user clicks Done
      5. Cleans up temp files if the user cancels
    """

    # Suppress CTkToplevel's deferred iconbitmap call (same fix as OrderDetailModal)
    def iconbitmap(self, *args, **kwargs):
        try:
            super().iconbitmap(*args, **kwargs)
        except Exception:
            pass

    def __init__(self, master, on_done: Callable[[list[str]], None], supplier_name: str = "Unknown"):
        super().__init__(master)
        self._on_done = on_done
        self._supplier_name = supplier_name
        self._received_paths: list[str] = []
        self._server: PhoneUploadServer | None = None

        self.title("Scan Invoice with Phone")
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self._build_ui()
        self.update_idletasks()
        self._start_server()
        self.after(150, self._activate)

    def _build_ui(self):
        # ── Heading ───────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Scan Invoice with Phone",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).pack(padx=28, pady=(22, 4))

        ctk.CTkLabel(
            self,
            text=f"Supplier: {self._supplier_name}",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("dodgerblue3", "dodgerblue2"),
        ).pack(padx=28, pady=(0, 6))

        ctk.CTkLabel(
            self,
            text="Scan the QR code with your phone to open the upload page.\n"
                 "Select photos from your library or take new ones.",
            font=ctk.CTkFont(size=13),
            text_color="gray60",
            justify="center",
        ).pack(padx=28, pady=(0, 14))

        # ── QR code ───────────────────────────────────────────────────────
        self._qr_label = ctk.CTkLabel(
            self, text="Starting server…",
            width=260, height=260,
        )
        self._qr_label.pack(padx=28)

        # ── URL (for manual entry if QR scan fails) ────────────────────────
        self._url_label = ctk.CTkLabel(
            self, text="",
            font=ctk.CTkFont(size=12, family="Courier New"),
            text_color="gray50",
        )
        self._url_label.pack(padx=28, pady=(8, 0))

        # ── Status ────────────────────────────────────────────────────────
        self._status_label = ctk.CTkLabel(
            self, text="Waiting for photos…",
            font=ctk.CTkFont(size=14),
            text_color="gray50",
        )
        self._status_label.pack(padx=28, pady=(10, 16))

        # ── Buttons ───────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=28, pady=(0, 24))

        self._done_btn = ctk.CTkButton(
            btn_row,
            text="Done — Process Photos",
            width=200,
            state="disabled",
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3"),
            command=self._done,
        )
        self._done_btn.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            btn_row,
            text="Cancel",
            width=90,
            fg_color="gray50",
            hover_color="gray40",
            command=self._cancel,
        ).pack(side="left")

    # ── Server lifecycle ──────────────────────────────────────────────────

    def _start_server(self):
        try:
            self._server = PhoneUploadServer(
                on_images_received=lambda paths: self.after(0, self._on_photos, paths),
                supplier_name=self._supplier_name,
            )
            self._server.start()
            self.after(0, self._show_qr)
        except Exception as exc:
            self._status_label.configure(
                text=f"Could not start server: {exc}",
                text_color="red",
            )

    def _show_qr(self):
        if self._server is None:
            return
        url = self._server.url
        self._url_label.configure(text=url)

        qr_img = _make_qr_image(url)
        if qr_img is not None:
            self._qr_label.configure(image=qr_img, text="")
            self._qr_label.image = qr_img  # prevent GC
        else:
            self._qr_label.configure(
                text=f"Install qrcode[pil] for a QR code.\n\nEnter this URL manually:\n{url}",
                wraplength=240,
            )

        # Snap window to natural size after inserting QR image
        self.update_idletasks()
        w = max(self.winfo_reqwidth(), 360)
        h = max(self.winfo_reqheight(), 480)
        self.geometry(f"{w}x{h}")

    def _stop_server(self):
        if self._server is not None:
            threading.Thread(target=self._server.stop, daemon=True).start()
            self._server = None

    # ── Photo events ─────────────────────────────────────────────────────

    def _on_photos(self, paths: list[str]):
        self._received_paths.extend(paths)
        n = len(self._received_paths)
        self._status_label.configure(
            text=f"✓  {n} photo{'s' if n != 1 else ''} received — click Done when finished",
            text_color=("green3", "green4"),
        )
        self._done_btn.configure(state="normal")

    # ── Actions ───────────────────────────────────────────────────────────

    def _done(self):
        paths = list(self._received_paths)
        self._stop_server()
        self.destroy()
        self._on_done(paths)

    def _cancel(self):
        self._stop_server()
        for p in self._received_paths:
            try:
                os.unlink(p)
            except OSError:
                pass
        self.destroy()

    def _activate(self):
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except Exception:
            pass

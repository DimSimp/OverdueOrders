"""Shared GUI dialogs used across multiple views."""
from __future__ import annotations

import webbrowser
from typing import Callable

import customtkinter as ctk


def show_ebay_sync_failed_dialog(
    parent,
    failed_orders: list[tuple[str, str]],
    on_close: Callable | None = None,
) -> ctk.CTkToplevel:
    """Show a modal dialog when one or more eBay orders could not be synced.

    Parameters
    ----------
    parent:
        Parent widget (used for positioning and grab_set).
    failed_orders:
        List of (ebay_order_id, error_message) tuples.
    on_close:
        Optional callback invoked after the dialog is dismissed.  Use this to
        trigger navigation (e.g. _on_fulfilled) once the user has acknowledged
        the failure.

    Returns the CTkToplevel so callers can configure it further if needed.
    """
    n = len(failed_orders)

    dlg = ctk.CTkToplevel(parent)
    dlg.title("eBay Update Required")
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.lift()

    def _close():
        dlg.destroy()
        if on_close:
            try:
                parent.after(0, on_close)
            except Exception:
                pass

    dlg.protocol("WM_DELETE_WINDOW", _close)

    # ── Header ────────────────────────────────────────────────────────────────
    header = ctk.CTkFrame(dlg, fg_color="#7D2323", corner_radius=0)
    header.pack(fill="x")
    ctk.CTkLabel(
        header,
        text="⚠  eBay order could not be updated automatically",
        font=ctk.CTkFont(size=13, weight="bold"),
        text_color="white",
    ).pack(padx=16, pady=10)

    # ── Body ──────────────────────────────────────────────────────────────────
    body = ctk.CTkFrame(dlg, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=20, pady=(14, 6))

    plural = "orders were" if n > 1 else "order was"
    ctk.CTkLabel(
        body,
        text=(
            f"{n} eBay {plural} dispatched in Neto but could not be updated\n"
            "via the eBay API. Please mark {'them' if n > 1 else 'it'} as dispatched manually."
        ),
        font=ctk.CTkFont(size=12),
        justify="left",
        wraplength=420,
    ).pack(anchor="w", pady=(0, 10))

    # One row per failed order
    scroll = ctk.CTkScrollableFrame(dlg, height=min(36 * n + 10, 160))
    scroll.pack(fill="x", padx=20, pady=(0, 6))

    for ebay_id, error_msg in failed_orders:
        url = f"https://www.ebay.com.au/sh/ord/details?orderid={ebay_id}"
        row = ctk.CTkFrame(scroll, fg_color="transparent")
        row.pack(fill="x", pady=3)

        ctk.CTkLabel(
            row,
            text=ebay_id,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(4, 0))

        if error_msg:
            ctk.CTkLabel(
                row,
                text=f"— {error_msg}",
                font=ctk.CTkFont(size=11),
                text_color="gray60",
            ).pack(side="left", padx=6)

        ctk.CTkButton(
            row,
            text="Open in eBay",
            width=110,
            height=28,
            command=lambda u=url: webbrowser.open(u),
        ).pack(side="right", padx=4)

    # ── Footer ────────────────────────────────────────────────────────────────
    ctk.CTkButton(
        dlg,
        text="Close",
        width=100,
        fg_color="gray50",
        hover_color="gray40",
        command=_close,
    ).pack(pady=(6, 16))

    # Size and centre over parent
    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    px = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
    dlg.geometry(f"{w}x{h}+{max(px, 0)}+{max(py, 0)}")

    return dlg

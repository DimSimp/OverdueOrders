from __future__ import annotations

import logging
import threading
from tkinter import messagebox

log = logging.getLogger(__name__)

import customtkinter as ctk

from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError

# Mirrors the channel map from options_view.py
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":          ["Website"],
    "eBay (via Neto)":  ["eBay"],
    "BigW":             ["BigW"],
    "Kogan":            ["Kogan"],
    "Amazon":           ["Amazon AU", "Amazon"],
    "Everydaymarket":   ["Everydaymarket"],
    "Control Panel":    ["Control Panel"],
    "Quote":            ["Quote"],
}


class AllOrdersFetchView(ctk.CTkFrame):
    """
    Step 2 — Fetch all pending orders from Neto and eBay with progress display.

    No books filter, no express/click-collect filter, no note phrase filter.
    Results are stored in window.all_orders_neto_orders / window.all_orders_ebay_orders.
    """

    def __init__(self, master, window, on_complete, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_complete = on_complete
        self._on_back = on_back
        self._fetch_done: dict[str, bool] = {}
        self._fetch_error: dict[str, str | None] = {}
        self._fetch_options: dict = {}
        self._build_ui()

    # ── UI ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=40, pady=20)

        self._heading_label = ctk.CTkLabel(
            center,
            text="Fetching orders…",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self._heading_label.pack(pady=(20, 12))

        self._progress = ctk.CTkProgressBar(center, mode="indeterminate")
        self._progress.pack(fill="x", pady=(0, 16))

        self._neto_status = ctk.CTkLabel(
            center, text="", font=ctk.CTkFont(size=13), anchor="w"
        )
        self._neto_status.pack(fill="x", pady=2)

        self._ebay_status = ctk.CTkLabel(
            center, text="", font=ctk.CTkFont(size=13), anchor="w"
        )
        self._ebay_status.pack(fill="x", pady=2)

        self._error_label = ctk.CTkLabel(
            center,
            text="",
            text_color="red",
            font=ctk.CTkFont(size=12),
            wraplength=700,
            justify="left",
            anchor="w",
        )
        self._error_label.pack(fill="x", pady=(8, 0))

        # Bottom row
        bottom = ctk.CTkFrame(center, fg_color="transparent")
        bottom.pack(fill="x", side="bottom", pady=(20, 0))

        self._back_btn = ctk.CTkButton(
            bottom,
            text="← Back to Options",
            width=160,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            state="disabled",
            command=self._on_back,
        )
        self._back_btn.pack(side="left")

        self._next_btn = ctk.CTkButton(
            bottom,
            text="Show Orders  →",
            width=180,
            state="disabled",
            command=self._on_next,
        )
        self._next_btn.pack(side="right")

    # ── Fetch ────────────────────────────────────────────────────────────

    def start_fetch(self, options: dict):
        """Called by DailyOpsWindow to begin the fetch with the given options."""
        self._fetch_options = options
        self._fetch_done = {"neto": False, "ebay": False}
        self._fetch_error = {"neto": None, "ebay": None}
        self._window.all_orders_neto_orders = []
        self._window.all_orders_ebay_orders = []

        self._back_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        self._error_label.configure(text="")
        self._heading_label.configure(text="Fetching orders…", text_color=("gray10", "gray90"))
        self._neto_status.configure(text="Fetching Neto orders…", text_color="gray60")
        self._progress.start()

        ebay_direct_on = options.get("ebay_direct", True)
        if ebay_direct_on:
            self._ebay_status.configure(text="Fetching eBay orders…", text_color="gray60")
        else:
            self._ebay_status.configure(text="eBay (direct): skipped", text_color="gray60")
            self._fetch_done["ebay"] = True

        ebay_via_neto = options.get("platforms", {}).get("eBay (via Neto)", True)

        threading.Thread(
            target=self._neto_worker,
            args=(options["date_from"], options["date_to"], ebay_via_neto),
            daemon=True,
        ).start()
        if ebay_direct_on:
            threading.Thread(
                target=self._ebay_worker,
                args=(options["date_from"], options["date_to"]),
                daemon=True,
            ).start()

    def _neto_worker(self, date_from, date_to, include_ebay_channel: bool):
        try:
            orders = self._window.neto_client.get_overdue_orders(
                date_from, date_to,
                include_ebay_channel=include_ebay_channel,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._neto_status.configure(
                        text=f"Fetching Neto orders… ({f}/{t})", text_color="gray60"
                    )
                ),
            )
            filtered = self._filter_neto_by_channel(orders)
            self.after(0, lambda n=len(filtered), raw=filtered: self._on_neto_done(n, raw))
        except NetoAPIError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("neto", f"Neto error: {msg}"))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("neto", f"Neto fetch failed: {msg}"))

    def _ebay_worker(self, date_from, date_to):
        try:
            orders = self._window.ebay_client.get_overdue_orders(
                date_from, date_to,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._ebay_status.configure(
                        text=f"Fetching eBay orders… ({f}/{t})", text_color="gray60"
                    )
                ),
            )
            warn = self._window.ebay_client.notes_warning
            self.after(0, lambda n=len(orders), raw=orders, w=warn: self._on_ebay_done(n, raw, w))
        except EbayAuthError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay auth error: {msg}"))
        except EbayAPIError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay API error: {msg}"))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay fetch failed: {msg}"))

    def _filter_neto_by_channel(self, orders: list) -> list:
        platforms = self._fetch_options.get("platforms", {})
        channel_enabled: dict[str, bool] = {}
        for label, channels in _NETO_CHANNEL_MAP.items():
            is_on = platforms.get(label, True)
            for ch in channels:
                channel_enabled[ch.lower()] = is_on

        result = []
        for order in orders:
            ch = (order.sales_channel or "").lower()
            if ch in channel_enabled:
                if channel_enabled[ch]:
                    result.append(order)
            else:
                result.append(order)
        return result

    # ── Callbacks ────────────────────────────────────────────────────────

    def _on_neto_done(self, count: int, orders: list):
        self._window.all_orders_neto_orders = orders
        self._fetch_done["neto"] = True
        self._neto_status.configure(
            text=f"Neto: {count} order{'s' if count != 1 else ''} fetched.",
            text_color="green",
        )
        self._check_both_done()

    def _on_ebay_done(self, count: int, orders: list, warning: str):
        self._window.all_orders_ebay_orders = orders
        self._fetch_done["ebay"] = True
        if warning:
            msg = f"eBay: {count} order{'s' if count != 1 else ''} fetched. ⚠ {warning}"
            self._ebay_status.configure(text=msg, text_color="orange")
        else:
            self._ebay_status.configure(
                text=f"eBay: {count} order{'s' if count != 1 else ''} fetched.",
                text_color="green",
            )
        self._check_both_done()

    def _on_platform_error(self, platform: str, message: str):
        self._fetch_done[platform] = True
        self._fetch_error[platform] = message

        if platform == "neto":
            self._neto_status.configure(text=message, text_color="red")
            self._check_both_done()
        else:
            self._ebay_status.configure(text=message, text_color="red")
            proceed = messagebox.askyesno(
                "eBay Unavailable",
                f"{message}\n\nContinue with Neto orders only?",
                parent=self,
            )
            if proceed:
                self._check_both_done()
            else:
                self._progress.stop()
                self._back_btn.configure(state="normal")
                self._window.all_orders_neto_orders = []
                self._ebay_status.configure(text="eBay fetch cancelled.", text_color="gray60")

    def _check_both_done(self):
        if not all(self._fetch_done.values()):
            return

        self._progress.stop()
        self._back_btn.configure(state="normal")

        neto = self._window.all_orders_neto_orders
        ebay = self._window.all_orders_ebay_orders
        total = len(neto) + len(ebay)

        if total == 0 and not any(self._fetch_error.values()):
            self._heading_label.configure(text="No orders found.", text_color=("gray50", "gray60"))
            self._error_label.configure(
                text="No orders found for the selected date range and platforms."
            )
        else:
            self._heading_label.configure(text="Orders fetched!", text_color="green")

        if total > 0 or any(self._fetch_done.values()):
            self._next_btn.configure(state="normal")

    def _on_next(self):
        self._on_complete()

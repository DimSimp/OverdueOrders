from __future__ import annotations

import logging
import threading
from tkinter import messagebox

log = logging.getLogger(__name__)

import customtkinter as ctk

from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError
from src.data_processor import exclude_phrases

# Mirrors the channel map from orders_tab.py
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":         ["Website"],
    "eBay (via Neto)": ["eBay"],
    "BigW":            ["BigW"],
    "Kogan":           ["Kogan"],
    "Amazon":          ["Amazon AU", "Amazon"],
}


class FetchView(ctk.CTkFrame):
    """
    Step 2 — Fetch orders from Neto and eBay with progress display.

    start_fetch(options) is called by DailyOpsWindow after the user confirms
    options. When both fetches are done (and filtering applied), on_complete()
    is called and results are stored in window.neto_orders / window.ebay_orders.
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
        pad = {"padx": 16, "pady": 8}

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

        self._filter_status = ctk.CTkLabel(
            center, text="", font=ctk.CTkFont(size=12), anchor="w",
            text_color=("gray50", "gray60"),
        )
        self._filter_status.pack(fill="x", pady=2)

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
            text="Next: Classify Envelopes →",
            width=200,
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
        self._window.neto_orders = []
        self._window.ebay_orders = []

        self._back_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        self._error_label.configure(text="")
        self._filter_status.configure(text="")
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
        self._window.neto_orders = orders
        self._fetch_done["neto"] = True
        self._neto_status.configure(
            text=f"Neto: {count} order{'s' if count != 1 else ''} fetched.",
            text_color="green",
        )
        self._check_both_done()

    def _on_ebay_done(self, count: int, orders: list, warning: str):
        self._window.ebay_orders = orders
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
                self._window.neto_orders = []
                self._ebay_status.configure(text="eBay fetch cancelled.", text_color="gray60")

    def _check_both_done(self):
        if not all(self._fetch_done.values()):
            return

        # Apply shipping-type filters first (no API call needed)
        include_express = self._fetch_options.get("include_express", True)
        include_cc = self._fetch_options.get("include_click_collect", False)

        def _keep(order) -> bool:
            st = getattr(order, "shipping_type", "") or ""
            if not include_express and st == "Express":
                return False
            if not include_cc and st == "Local Pickup":
                return False
            return True

        neto = [o for o in self._window.neto_orders if _keep(o)]
        ebay = [o for o in self._window.ebay_orders if _keep(o)]

        # Now look up ShippingCategory from Neto product catalogue to exclude Books (cat ID "4")
        # This requires an API call so run it in a background thread.
        self._neto_status.configure(
            text=self._neto_status.cget("text") + "  (checking categories…)",
            text_color="gray60",
        )
        threading.Thread(
            target=self._books_filter_worker,
            args=(neto, ebay),
            daemon=True,
        ).start()

    def _books_filter_worker(self, neto: list, ebay: list):
        """Background thread: look up ShippingCategory + Misc06 for all Neto SKUs.
        - Populates postage_type on each NetoLineItem from Misc06 (product-level field)
        - Filters out orders where any line item has ShippingCategory "4" (Books)
        """
        try:
            all_skus = list({
                li.sku for o in neto for li in o.line_items if li.sku
            })
            log.debug("Checking item attributes for %d SKUs: %s", len(all_skus), all_skus)
            attr_map = self._window.neto_client.get_item_attributes(all_skus)
            log.debug("Item attributes: %s", attr_map)

            # Populate postage_type on each line item from product-level Misc06
            for order in neto:
                for li in order.line_items:
                    attrs = attr_map.get(li.sku, {})
                    if attrs.get("postage_type"):
                        li.postage_type = attrs["postage_type"]

            # Category ID "4" = Books — filter out orders containing any Books SKU
            books_skus = {
                sku for sku, attrs in attr_map.items()
                if attrs.get("shipping_category") == "4"
            }
            log.debug("Books SKUs (cat 4): %s", books_skus)
            before = len(neto)
            if books_skus:
                neto = [
                    o for o in neto
                    if not any(li.sku in books_skus for li in o.line_items)
                ]
            excluded = before - len(neto)
            log.debug(
                "Books filter: %d order%s excluded, %d remain",
                excluded, "s" if excluded != 1 else "", len(neto),
            )
        except Exception as exc:
            log.warning("Item attribute lookup failed, keeping all orders: %s", exc)
        self.after(0, lambda n=neto, e=ebay: self._apply_final_filters(n, e))

    def _apply_final_filters(self, neto: list, ebay: list):
        self._progress.stop()
        self._back_btn.configure(state="normal")
        self._heading_label.configure(text="Orders fetched!", text_color="green")

        # Clean up Neto status — strip the "(checking categories…)" suffix
        neto_text = self._neto_status.cget("text")
        neto_text = neto_text.replace("  (checking categories…)", "").strip()
        self._neto_status.configure(text=neto_text, text_color="green")

        # Apply note-phrase exclusion filter
        phrases = self._fetch_options.get("note_filter_phrases", [])
        if phrases:
            neto_before, ebay_before = len(neto), len(ebay)
            neto = exclude_phrases(neto, phrases)
            ebay = exclude_phrases(ebay, phrases)
            excluded = (neto_before - len(neto)) + (ebay_before - len(ebay))
            if excluded:
                self._filter_status.configure(
                    text=f"Filtered out {excluded} order{'s' if excluded != 1 else ''} "
                         f"matching note phrases.",
                    text_color=("gray50", "gray60"),
                )

        self._window.neto_orders = neto
        self._window.ebay_orders = ebay

        total = len(neto) + len(ebay)
        if total == 0 and not any(self._fetch_error.values()):
            self._error_label.configure(
                text="No orders found for the selected date range and filters."
            )

        if total > 0 or any(self._fetch_done.values()):
            self._next_btn.configure(state="normal")

    def _on_next(self):
        self._on_complete()

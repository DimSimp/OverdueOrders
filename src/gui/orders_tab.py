from __future__ import annotations

import threading
from datetime import datetime, timedelta
from tkinter import messagebox, simpledialog

import customtkinter as ctk

from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError

# Neto SalesChannel values covered by each toggle label.
# Orders from channels NOT listed here are always included.
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":         ["Website"],
    "eBay (via Neto)": ["eBay"],
    "BigW":            ["BigW"],
    "Kogan":           ["Kogan"],
    "Amazon":          ["Amazon AU", "Amazon"],
}


class OrdersTab(ctk.CTkFrame):
    def __init__(self, master, app, on_complete, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._on_complete = on_complete
        self._fetch_done: dict[str, bool] = {"neto": False, "ebay": False}
        self._fetch_error: dict[str, str | None] = {"neto": None, "ebay": None}
        self._platform_switches: dict[str, ctk.CTkSwitch] = {}
        self._ebay_direct_switch: ctk.CTkSwitch | None = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        # ── Date range row ────────────────────────────────────────────────
        date_frame = ctk.CTkFrame(self, fg_color="transparent")
        date_frame.pack(fill="x", **pad)

        ctk.CTkLabel(date_frame, text="Orders from:", font=ctk.CTkFont(size=13)).pack(side="left")

        default_from = (datetime.today() - timedelta(days=self._app.config.app.order_lookback_days))
        default_to = datetime.today()

        self._from_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="YYYY-MM-DD")
        self._from_entry.insert(0, default_from.strftime("%Y-%m-%d"))
        self._from_entry.pack(side="left", padx=(6, 12))

        ctk.CTkLabel(date_frame, text="to:", font=ctk.CTkFont(size=13)).pack(side="left")

        self._to_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="YYYY-MM-DD")
        self._to_entry.insert(0, default_to.strftime("%Y-%m-%d"))
        self._to_entry.pack(side="left", padx=(6, 0))

        # ── Platform toggles ──────────────────────────────────────────────
        plat_frame = ctk.CTkFrame(self, fg_color="transparent")
        plat_frame.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(
            plat_frame, text="Platforms:", font=ctk.CTkFont(size=13)
        ).pack(side="left")

        for label in _NETO_CHANNEL_MAP:
            sw = ctk.CTkSwitch(plat_frame, text=label, font=ctk.CTkFont(size=12), width=50)
            sw.select()
            sw.pack(side="left", padx=(14, 0))
            self._platform_switches[label] = sw

        # eBay direct API toggle (separate — controls whether to call eBay Fulfillment API)
        ebay_sw = ctk.CTkSwitch(plat_frame, text="eBay (direct)", font=ctk.CTkFont(size=12), width=50)
        ebay_sw.select()
        ebay_sw.pack(side="left", padx=(14, 0))
        self._ebay_direct_switch = ebay_sw

        # ── Action row ────────────────────────────────────────────────────
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.pack(fill="x", **pad)

        self._fetch_btn = ctk.CTkButton(
            action_frame,
            text="Fetch Overdue Orders",
            width=180,
            command=self._fetch_orders,
        )
        self._fetch_btn.pack(side="left")

        # eBay auth status indicator
        self._ebay_auth_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        self._ebay_auth_frame.pack(side="left", padx=(20, 0))

        self._ebay_auth_label = ctk.CTkLabel(
            self._ebay_auth_frame,
            text="",
            font=ctk.CTkFont(size=12),
        )
        self._ebay_auth_label.pack(side="left")

        self._ebay_auth_btn = ctk.CTkButton(
            self._ebay_auth_frame,
            text="Authenticate eBay",
            width=140,
            height=26,
            font=ctk.CTkFont(size=12),
            command=self._do_ebay_auth,
        )
        self._ebay_auth_btn.pack(side="left", padx=(8, 0))

        self._update_ebay_status()

        # ── Progress bar ──────────────────────────────────────────────────
        self._progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(0, 4))
        self._progress.pack_forget()  # hidden until fetch

        # ── Status labels ─────────────────────────────────────────────────
        self._neto_status = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=13), anchor="w"
        )
        self._neto_status.pack(fill="x", padx=12, pady=2)

        self._ebay_status_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=13), anchor="w"
        )
        self._ebay_status_label.pack(fill="x", padx=12, pady=2)

        self._error_label = ctk.CTkLabel(
            self,
            text="",
            text_color="red",
            font=ctk.CTkFont(size=12),
            wraplength=800,
            justify="left",
            anchor="w",
        )
        self._error_label.pack(fill="x", padx=12, pady=4)

        # ── Bottom row ────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=12, pady=(4, 12), side="bottom")

        self._next_btn = ctk.CTkButton(
            bottom,
            text="Next: View Results →",
            width=180,
            state="disabled",
            command=self._on_complete,
        )
        self._next_btn.pack(side="right")

    # ── eBay auth ─────────────────────────────────────────────────────────

    def _update_ebay_status(self):
        if self._app.ebay_client.is_authenticated():
            self._ebay_auth_label.configure(text="eBay: Authenticated", text_color="green")
            self._ebay_auth_btn.configure(text="Re-authenticate eBay")
        else:
            self._ebay_auth_label.configure(text="eBay: Not authenticated", text_color="orange")
            self._ebay_auth_btn.configure(text="Authenticate eBay")

    def _do_ebay_auth(self):
        self._app.ebay_client.open_auth_in_browser()
        # Ask user to paste the redirect URL
        url = simpledialog.askstring(
            title="eBay Authentication",
            prompt=(
                "After approving access in your browser, paste the full URL\n"
                "from the browser address bar here (it will start with your redirect URI):"
            ),
            parent=self,
        )
        if not url:
            return
        try:
            self._app.ebay_client.exchange_code(url.strip())
            self._update_ebay_status()
        except EbayAuthError as e:
            self._set_error(str(e))
        except Exception as e:
            self._set_error(f"eBay authentication failed: {e}")

    # ── Fetch orders ──────────────────────────────────────────────────────

    def _fetch_orders(self):
        date_from, date_to = self._parse_dates()
        if date_from is None:
            return

        ebay_direct_on = self._ebay_direct_switch.get() == 1
        # Whether to include eBay-channel Neto orders (normally excluded)
        ebay_via_neto_on = self._platform_switches["eBay (via Neto)"].get() == 1

        self._fetch_done = {"neto": False, "ebay": False}
        self._fetch_error = {"neto": None, "ebay": None}
        self._app.neto_orders = []
        self._app.ebay_orders = []

        self._fetch_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        self._set_error("")
        self._neto_status.configure(text="Fetching Neto orders…", text_color="gray60")
        self._progress.pack(fill="x", padx=12, pady=(0, 4))
        self._progress.start()

        if ebay_direct_on:
            self._ebay_status_label.configure(text="Fetching eBay orders…", text_color="gray60")
        else:
            self._ebay_status_label.configure(text="eBay (direct): skipped", text_color="gray60")
            self._fetch_done["ebay"] = True  # no worker needed

        threading.Thread(
            target=self._neto_worker,
            args=(date_from, date_to, ebay_via_neto_on),
            daemon=True,
        ).start()
        if ebay_direct_on:
            threading.Thread(
                target=self._ebay_worker, args=(date_from, date_to), daemon=True
            ).start()

    def _neto_worker(self, date_from, date_to, include_ebay_channel: bool):
        try:
            orders = self._app.neto_client.get_overdue_orders(
                date_from, date_to,
                include_ebay_channel=include_ebay_channel,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._neto_status.configure(
                        text=f"Fetching Neto orders… ({f}/{t})", text_color="gray60"
                    )
                ),
            )
            filtered = self._filter_neto_orders(orders)
            self._app.neto_orders = filtered
            self.after(0, lambda n=len(filtered): self._on_platform_done(
                "neto", f"Neto: {n} order{'s' if n != 1 else ''} fetched.", "green"
            ))
        except NetoAPIError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("neto", f"Neto error: {msg}"))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("neto", f"Neto fetch failed: {msg}"))

    def _filter_neto_orders(self, orders: list) -> list:
        """Filter Neto orders to only include enabled platform channels."""
        # Build map: sales_channel (lowercase) → enabled
        channel_enabled: dict[str, bool] = {}
        for label, channels in _NETO_CHANNEL_MAP.items():
            sw = self._platform_switches.get(label)
            is_on = sw.get() == 1 if sw else True
            for ch in channels:
                channel_enabled[ch.lower()] = is_on

        result = []
        for order in orders:
            ch = order.sales_channel.lower()
            if ch in channel_enabled:
                if channel_enabled[ch]:
                    result.append(order)
            else:
                # Channel not covered by any toggle — always include
                result.append(order)
        return result

    def _ebay_worker(self, date_from, date_to):
        try:
            orders = self._app.ebay_client.get_overdue_orders(
                date_from, date_to,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._ebay_status_label.configure(
                        text=f"Fetching eBay orders… ({f}/{t})", text_color="gray60"
                    )
                ),
            )
            self._app.ebay_orders = orders
            self.after(0, lambda n=len(orders): self._on_platform_done(
                "ebay", f"eBay: {n} order{'s' if n != 1 else ''} fetched.", "green"
            ))
        except EbayAuthError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay auth error: {msg}"))
        except EbayAPIError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay API error: {msg}"))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay fetch failed: {msg}"))

    def _on_platform_done(self, platform: str, message: str, color: str):
        self._fetch_done[platform] = True
        if platform == "neto":
            self._neto_status.configure(text=message, text_color=color)
        else:
            self._ebay_status_label.configure(text=message, text_color=color)
        self._check_both_done()

    def _on_platform_error(self, platform: str, message: str):
        self._fetch_done[platform] = True
        self._fetch_error[platform] = message

        if platform == "neto":
            self._neto_status.configure(text=message, text_color="red")
            self._check_both_done()
        else:
            # eBay error — show the error then ask whether to continue
            self._ebay_status_label.configure(text=message, text_color="red")
            proceed = messagebox.askyesno(
                "eBay Unavailable",
                f"{message}\n\nContinue with Neto orders only?",
                parent=self,
            )
            if proceed:
                self._check_both_done()
            else:
                # User chose not to continue — reset fetch state
                self._progress.stop()
                self._progress.pack_forget()
                self._fetch_btn.configure(state="normal")
                self._app.neto_orders = []
                self._ebay_status_label.configure(
                    text="eBay fetch cancelled.", text_color="gray60"
                )

    def _check_both_done(self):
        if not all(self._fetch_done.values()):
            return
        self._progress.stop()
        self._progress.pack_forget()
        self._fetch_btn.configure(state="normal")

        errors = [e for e in self._fetch_error.values() if e]
        if errors:
            self._set_error("Some platforms had errors — see details above. You can still proceed with partial results.")

        # Enable Next if at least one platform succeeded
        if self._app.neto_orders or self._app.ebay_orders:
            self._next_btn.configure(state="normal")

    def _parse_dates(self):
        from_str = self._from_entry.get().strip()
        to_str = self._to_entry.get().strip()
        try:
            date_from = datetime.strptime(from_str, "%Y-%m-%d")
        except ValueError:
            self._set_error(f"Invalid 'From' date: '{from_str}'. Use YYYY-MM-DD format.")
            return None, None
        try:
            date_to = datetime.strptime(to_str, "%Y-%m-%d")
        except ValueError:
            self._set_error(f"Invalid 'To' date: '{to_str}'. Use YYYY-MM-DD format.")
            return None, None
        if date_from > date_to:
            self._set_error("'From' date must be before 'To' date.")
            return None, None
        return date_from, date_to

    def _set_error(self, text: str):
        self._error_label.configure(text=text)

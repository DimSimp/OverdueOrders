from __future__ import annotations

import os
import threading
from datetime import datetime, timedelta
from tkinter import messagebox

import customtkinter as ctk

from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError

# Neto SalesChannel values covered by each toggle label — kept in sync with orders_tab.py
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":         ["Website"],
    "eBay (via Neto)": ["eBay"],
    "BigW":            ["BigW"],
    "Kogan":           ["Kogan"],
    "Amazon":          ["Amazon AU", "Amazon"],
}


class AftNewSessionView(ctk.CTkFrame):
    """
    Combined options + automated-run view for starting a new afternoon session.

    Two internal panels:
      1. Options panel — date range + platform toggles
      2. Progress panel — inventory comparison → order fetch → auto-jump to Results
    """

    def __init__(self, master, app, on_complete, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._on_complete = on_complete
        self._on_back = on_back

        self._fetch_done: dict[str, bool] = {"neto": False, "ebay": False}
        self._fetch_error: dict[str, str | None] = {"neto": None, "ebay": None}
        self._platform_switches: dict[str, ctk.CTkSwitch] = {}
        self._ebay_direct_switch: ctk.CTkSwitch | None = None

        self._options_panel: ctk.CTkFrame | None = None
        self._progress_panel: ctk.CTkFrame | None = None

        self._build_options_panel()
        self._build_progress_panel()

        # Start on options panel
        self._options_panel.pack(fill="both", expand=True)

    # ── Options panel ─────────────────────────────────────────────────────

    def _build_options_panel(self):
        panel = ctk.CTkFrame(self, fg_color="transparent")
        self._options_panel = panel

        inner = ctk.CTkFrame(panel, fg_color="transparent")
        inner.pack(expand=True, fill="both", padx=80, pady=40)

        ctk.CTkLabel(
            inner,
            text="New Session — Options",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", pady=(0, 20))

        # ── Date range ────────────────────────────────────────────────────
        date_frame = ctk.CTkFrame(inner, fg_color="transparent")
        date_frame.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(date_frame, text="Orders from:", font=ctk.CTkFont(size=13)).pack(side="left")

        lookback = self._app.config.app.order_lookback_days
        default_from = datetime.today() - timedelta(days=lookback)
        default_to = datetime.today()

        self._from_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._from_entry.insert(0, default_from.strftime("%d/%m/%Y"))
        self._from_entry.pack(side="left", padx=(6, 12))

        ctk.CTkLabel(date_frame, text="to:", font=ctk.CTkFont(size=13)).pack(side="left")

        self._to_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._to_entry.insert(0, default_to.strftime("%d/%m/%Y"))
        self._to_entry.pack(side="left", padx=(6, 0))

        # ── Platform toggles ──────────────────────────────────────────────
        plat_label_frame = ctk.CTkFrame(inner, fg_color="transparent")
        plat_label_frame.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(plat_label_frame, text="Platforms:", font=ctk.CTkFont(size=13)).pack(side="left")

        plat_frame = ctk.CTkFrame(inner, fg_color="transparent")
        plat_frame.pack(fill="x", pady=(0, 24))

        saved_toggles = self._app.config._raw.get("app", {}).get("platform_toggles", {})

        for label in _NETO_CHANNEL_MAP:
            sw = ctk.CTkSwitch(plat_frame, text=label, font=ctk.CTkFont(size=12), width=50,
                               command=self._save_toggle_states)
            if saved_toggles.get(label, True):
                sw.select()
            else:
                sw.deselect()
            sw.pack(side="left", padx=(0, 14))
            self._platform_switches[label] = sw

        ebay_sw = ctk.CTkSwitch(plat_frame, text="eBay (direct)", font=ctk.CTkFont(size=12), width=50,
                                command=self._save_toggle_states)
        if saved_toggles.get("eBay (direct)", True):
            ebay_sw.select()
        else:
            ebay_sw.deselect()
        ebay_sw.pack(side="left", padx=(0, 14))
        self._ebay_direct_switch = ebay_sw

        # ── Error label ───────────────────────────────────────────────────
        self._options_error = ctk.CTkLabel(
            inner, text="", text_color="red", font=ctk.CTkFont(size=12),
            wraplength=700, justify="left", anchor="w",
        )
        self._options_error.pack(fill="x", pady=(0, 8))

        # ── Bottom buttons ────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(inner, fg_color="transparent")
        btn_row.pack(fill="x", side="bottom", pady=(0, 8))

        ctk.CTkButton(
            btn_row, text="← Back", width=100, command=self._on_back,
        ).pack(side="left")

        ctk.CTkButton(
            btn_row, text="Start →", width=120, command=self._start,
        ).pack(side="right")

    # ── Progress panel ────────────────────────────────────────────────────

    def _build_progress_panel(self):
        panel = ctk.CTkFrame(self, fg_color="transparent")
        self._progress_panel = panel

        inner = ctk.CTkFrame(panel, fg_color="transparent")
        inner.pack(expand=True, fill="both", padx=80, pady=40)

        ctk.CTkLabel(
            inner,
            text="Starting new session…",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", pady=(0, 16))

        self._progress_bar = ctk.CTkProgressBar(inner, mode="indeterminate")
        self._progress_bar.pack(fill="x", pady=(0, 16))

        self._inv_status = ctk.CTkLabel(
            inner, text="Comparing inventory reports…",
            font=ctk.CTkFont(size=13), anchor="w", text_color=("gray50", "gray60"),
        )
        self._inv_status.pack(fill="x", pady=2)

        self._neto_status = ctk.CTkLabel(
            inner, text="Neto: waiting…",
            font=ctk.CTkFont(size=13), anchor="w", text_color=("gray50", "gray60"),
        )
        self._neto_status.pack(fill="x", pady=2)

        self._ebay_status = ctk.CTkLabel(
            inner, text="eBay: waiting…",
            font=ctk.CTkFont(size=13), anchor="w", text_color=("gray50", "gray60"),
        )
        self._ebay_status.pack(fill="x", pady=2)

        self._progress_error = ctk.CTkLabel(
            inner, text="", text_color="red", font=ctk.CTkFont(size=12),
            wraplength=700, justify="left", anchor="w",
        )
        self._progress_error.pack(fill="x", pady=(8, 0))

        self._back_btn = ctk.CTkButton(
            inner, text="← Back to Menu", width=140, command=self._on_back,
        )
        # Hidden until a fatal error occurs
        self._back_btn.pack(anchor="w", pady=(12, 0))
        self._back_btn.pack_forget()

    # ── Logic ──────────────────────────────────────────────────────────────

    def _save_toggle_states(self):
        states = {label: (sw.get() == 1) for label, sw in self._platform_switches.items()}
        states["eBay (direct)"] = self._ebay_direct_switch.get() == 1
        self._app.config._raw.setdefault("app", {})["platform_toggles"] = states
        self._app.config.save()

    def _start(self):
        date_from, date_to = self._parse_dates()
        if date_from is None:
            return

        self._save_toggle_states()

        # Capture toggle states for use in background threads
        self._date_from = date_from
        self._date_to = date_to
        self._ebay_via_neto_on = self._platform_switches["eBay (via Neto)"].get() == 1
        self._ebay_direct_on = self._ebay_direct_switch.get() == 1
        self._toggle_snapshot = {
            label: (sw.get() == 1) for label, sw in self._platform_switches.items()
        }

        # Switch to progress panel
        self._options_panel.pack_forget()
        self._progress_panel.pack(fill="both", expand=True)

        # Reset progress state
        self._inv_status.configure(text="Comparing inventory reports…", text_color=("gray50", "gray60"))
        self._neto_status.configure(text="Neto: waiting…", text_color=("gray50", "gray60"))
        self._ebay_status.configure(text="eBay: waiting…", text_color=("gray50", "gray60"))
        self._progress_error.configure(text="")
        self._back_btn.pack_forget()
        self._fetch_done = {"neto": False, "ebay": False}
        self._fetch_error = {"neto": None, "ebay": None}
        self._app.neto_orders = []
        self._app.ebay_orders = []

        self._progress_bar.start()

        threading.Thread(target=self._run_session, daemon=True).start()

    def _run_session(self):
        """Background thread: compare inventory files, then launch order fetch."""
        from src.session import clear_overrides
        clear_overrides()

        ftp_cfg = self._app.config.ftp
        if ftp_cfg is None or not ftp_cfg.local_inventory_dir:
            self.after(0, lambda: self._fatal_error(
                'Local inventory directory not configured.\n'
                'Add "local_inventory_dir" to the "ftp" section of config.json.'
            ))
            return

        morning_path = os.path.join(ftp_cfg.local_inventory_dir, ftp_cfg.morning_filename)
        afternoon_path = os.path.join(ftp_cfg.local_inventory_dir, ftp_cfg.afternoon_filename)

        if not os.path.exists(morning_path):
            self.after(0, lambda: self._fatal_error(f"Morning report not found:\n{morning_path}"))
            return
        if not os.path.exists(afternoon_path):
            self.after(0, lambda: self._fatal_error(f"Afternoon report not found:\n{afternoon_path}"))
            return

        try:
            from src.ftp_inventory import compare_local_files
            received = compare_local_files(morning_path, afternoon_path)
        except Exception as e:
            self.after(0, lambda msg=str(e): self._fatal_error(f"Inventory comparison failed:\n{msg}"))
            return

        # Convert to InvoiceItem list
        from src.pdf_parser import InvoiceItem
        items = [
            InvoiceItem(
                sku=r.sku,
                sku_with_suffix=r.sku,
                description="",
                quantity=max(1, int(r.quantity)),
                source_page=0,
                supplier_name=r.supplier,
            )
            for r in received
        ]

        self.after(0, lambda i=items: self._after_inventory(i))

    def _after_inventory(self, items):
        """Main thread: inventory done — update UI and launch order fetch."""
        self._app.invoice_tab.set_invoice_items(items)
        n = len(items)
        self._inv_status.configure(
            text=f"Inventory: {n} item{'s' if n != 1 else ''} received",
            text_color="green",
        )

        # Launch Neto and eBay fetch in parallel
        if self._ebay_direct_on:
            self._ebay_status.configure(text="eBay: fetching…", text_color=("gray50", "gray60"))
        else:
            self._ebay_status.configure(text="eBay (direct): skipped", text_color=("gray50", "gray60"))
            self._fetch_done["ebay"] = True

        self._neto_status.configure(text="Neto: fetching…", text_color=("gray50", "gray60"))

        threading.Thread(
            target=self._neto_worker,
            args=(self._date_from, self._date_to, self._ebay_via_neto_on),
            daemon=True,
        ).start()

        if self._ebay_direct_on:
            threading.Thread(
                target=self._ebay_worker,
                args=(self._date_from, self._date_to),
                daemon=True,
            ).start()
        else:
            self._check_both_done()

    def _neto_worker(self, date_from, date_to, include_ebay_channel: bool):
        try:
            orders = self._app.neto_client.get_overdue_orders(
                date_from, date_to,
                include_ebay_channel=include_ebay_channel,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._neto_status.configure(
                        text=f"Neto: fetching… ({f}/{t})", text_color=("gray50", "gray60")
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

    def _ebay_worker(self, date_from, date_to):
        try:
            orders = self._app.ebay_client.get_overdue_orders(
                date_from, date_to,
                progress_callback=lambda f, t: self.after(
                    0, lambda: self._ebay_status.configure(
                        text=f"eBay: fetching… ({f}/{t})", text_color=("gray50", "gray60")
                    )
                ),
            )
            self._app.ebay_orders = orders
            warn = self._app.ebay_client.notes_warning
            if warn:
                msg = f"eBay: {len(orders)} order{'s' if len(orders) != 1 else ''} fetched. ⚠ {warn}"
                self.after(0, lambda m=msg: self._on_platform_done("ebay", m, "orange"))
            else:
                self.after(0, lambda n=len(orders): self._on_platform_done(
                    "ebay", f"eBay: {n} order{'s' if n != 1 else ''} fetched.", "green"
                ))
        except EbayAuthError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay auth error: {msg}"))
        except EbayAPIError as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay API error: {msg}"))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._on_platform_error("ebay", f"eBay fetch failed: {msg}"))

    def _filter_neto_orders(self, orders: list) -> list:
        channel_enabled: dict[str, bool] = {}
        for label, channels in _NETO_CHANNEL_MAP.items():
            is_on = self._toggle_snapshot.get(label, True)
            for ch in channels:
                channel_enabled[ch.lower()] = is_on
        result = []
        for order in orders:
            ch = order.sales_channel.lower()
            if ch in channel_enabled:
                if channel_enabled[ch]:
                    result.append(order)
            else:
                result.append(order)
        return result

    def _on_platform_done(self, platform: str, message: str, color: str):
        self._fetch_done[platform] = True
        if platform == "neto":
            self._neto_status.configure(text=message, text_color=color)
        else:
            self._ebay_status.configure(text=message, text_color=color)
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
                self._progress_bar.stop()
                self._app.neto_orders = []
                self._progress_error.configure(text="eBay fetch cancelled.")
                self._back_btn.pack(anchor="w", pady=(12, 0))

    def _check_both_done(self):
        if not all(self._fetch_done.values()):
            return
        self._progress_bar.stop()

        errors = [e for e in self._fetch_error.values() if e]
        if errors:
            self._progress_error.configure(
                text="Some platforms had errors — proceeding with available results."
            )

        if self._app.neto_orders or self._app.ebay_orders:
            self._on_complete()
        else:
            self._fatal_error("No orders were fetched. Check your date range and platform settings.")

    def _fatal_error(self, message: str):
        """Show a fatal error on the progress panel and stop — user must go back."""
        self._progress_bar.stop()
        self._progress_error.configure(text=message)
        self._back_btn.pack(anchor="w", pady=(12, 0))

    def _parse_dates(self):
        from_str = self._from_entry.get().strip()
        to_str = self._to_entry.get().strip()
        try:
            date_from = datetime.strptime(from_str, "%d/%m/%Y")
        except ValueError:
            self._options_error.configure(text=f"Invalid 'From' date: '{from_str}'. Use DD/MM/YYYY.")
            return None, None
        try:
            date_to = datetime.strptime(to_str, "%d/%m/%Y")
        except ValueError:
            self._options_error.configure(text=f"Invalid 'To' date: '{to_str}'. Use DD/MM/YYYY.")
            return None, None
        if date_from > date_to:
            self._options_error.configure(text="'From' date must be before 'To' date.")
            return None, None
        self._options_error.configure(text="")
        return date_from, date_to
